VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_09TNCN"
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
    Dim tuThang, denThang As Variant
    Dim varTuThang, varDenThang As Variant

    With fps
        .Sheet = 1 'To khai 09/TNCN
        ' Set ky cho thang bat dau tinh thue
        
        SetDateFormat fps, 1, 4, .ColLetterToNumber("L"), MMYYYY
        SetDateFormat fps, 1, 4, .ColLetterToNumber("N"), MMYYYY
        If TAX_Utilities_New.Year = 2009 Then
            .Row = 4
            .Col = .ColLetterToNumber("L")
            .Text = Format("07/2009", "mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
            
            ' Set ky cho thang ket thuc tinh thue
            .Row = 4
            .Col = .ColLetterToNumber("N")
            .Text = Format("12/2009", "mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
        Else
            .Row = 4
            .Col = .ColLetterToNumber("L")
            .Text = Format("01/" & TAX_Utilities_New.Year, "mm/yyyy")
            .TypeHAlign = TypeHAlignLeft

            .Row = 4
            .Col = .ColLetterToNumber("N")
            .Text = Format("12/" & TAX_Utilities_New.Year, "mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
        End If
                
        
        
        ' Tinh tong so thang theo dinh dang tu ky den ky
'        .Row = 8
'        .Col = .ColLetterToNumber("P")
'        .Text = DateDiff("M", Format("07/2009", "mm/yyyy"), Format("12/2009", "mm/yyyy")) + 1
        .GetText .ColLetterToNumber("L"), 4, tuThang
        .GetText .ColLetterToNumber("N"), 4, denThang
        
        'set so thang quyet toan
        .Row = 4
        .Col = .ColLetterToNumber("P")
        .Text = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy(CStr(denThang))) + 1

        
        ' Set format va gia tri cho ngay ky
        SetDateFormat fps, 1, 65, .ColLetterToNumber("M"), DDMMYYYY
                
        .Row = 65
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
               
        ' Dinh dang ngay thang cho cot ngay sinh trong phu luc 09C/TNCN
        SetDateFormat fps, 4, 22, .ColLetterToNumber("E"), DDMMYYYY
              
        .Sheet = 4
        ' Lay danh muc quan he voi NNT
        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_QHNNT.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        Dim strQuanHeNNTID As String
        Dim strQuanHeNNT As String
        ' Chuoi gia tri ve quan he NNT bao gom:
        ' 01-Con, 02-Vo, 03-Chong, 04-Cha, 05-Me, 06-Chau Ruot, 07-Ho Hang, 99-Khac
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            strQuanHeNNT = strQuanHeNNT + GetAttribute(xmlNode, "value") + Chr$(9)
        Next
        strQuanHeNNTID = "01" + Chr$(9) + "02" + Chr$(9) + "03" + Chr$(9) + "04" + Chr$(9) + "05" + Chr$(9) + "06" + Chr$(9) + "07" + Chr$(9) + "99" + Chr$(9)
        .Row = 22
        .Col = .ColLetterToNumber("I")
        
        .TypeComboBoxList = strQuanHeNNT
        
        .Col = .ColLetterToNumber("H")
        .TypeComboBoxList = strQuanHeNNTID
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim tuThang, denThang As Variant
    'bien lay gia tri chi tieu 29,30 lay luy tien cua chi tieu 31
    Dim lT31, oldValue As Variant
    Dim xmlNode   As MSXML.IXMLDOMNode
    Dim iCol      As Long, iRow As Long
    Dim varCT15   As Variant, varCT31 As Variant
    Dim soThangKG As Integer
    Dim viTriMSTDL As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If Strloaitk = "TKCT" And Trim(StrSolanBosung) = "" Then
            .Col = .ColLetterToNumber("C")
            .Row = 70
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
            .Col = .ColLetterToNumber("C")
            .Row = 70
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ''       lay gia tri CT32 da ghi trong datafile neu co
        For Each xmlNode In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlNode, "CellID"), iCol, iRow

            If iCol = .ColLetterToNumber("I") And iRow = 52 Then
                oldValue = GetAttribute(xmlNode, "Value")
                .Col = iCol
                .Row = iRow
                .Text = oldValue
'                .Formula = ""
            End If

            ' chi tieu 57
            ' dntai 17022012
            If iCol = .ColLetterToNumber("I") And iRow = 57 Then
                oldValue = GetAttribute(xmlNode, "Value")
                .Col = iCol
                .Row = iRow
                .Text = oldValue
'                .Formula = ""
            End If

        Next
        
        'set so thang tinh thue neu la nam 2011 thi co tinh so thang mien giam
        .GetText .ColLetterToNumber("L"), 4, tuThang
        .GetText .ColLetterToNumber("N"), 4, denThang

        'tinh so thang ke khai
        'ke khai nam 2011 se settext vao cell Q, <>2011 set vao cell P
        'neu ke khai nam 2009 thi so thang quyet toan = 6
        soThangKG = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1
        If TAX_Utilities_New.Year = "2011" Or TAX_Utilities_New.Year = "2012" Then
            'tinh so thang Ke khai
            soThangKG = getSoThangQTMienGiam(CStr(tuThang), CStr(denThang))
            .SetText .ColLetterToNumber("Q"), 4, soThangKG
            If isMienGiamPL26(CStr(tuThang), CStr(denThang)) Then
                .SetText .ColLetterToNumber("T"), 51, 1
            Else
                .SetText .ColLetterToNumber("T"), 51, 0
            End If
        ElseIf TAX_Utilities_New.Year = "2009" Then
            .SetText .ColLetterToNumber("Q"), 4, "6"
        Else
            .SetText .ColLetterToNumber("Q"), 4, "1"
        End If
        .Col = .ColLetterToNumber("P")
        .Row = 4
        .Text = DateDiff("M", Format(CStr(tuThang), "mm/yyyy"), Format(CStr(denThang), "mm/yyyy")) + 1
        
        'check CT 15 < 0 thi khong cho nhap CT 16 tren PL 09B
        .Sheet = 3
        .GetText .ColLetterToNumber("I"), 26, varCT15

        If Val(varCT15) < 0 Then
            .Col = .ColLetterToNumber("I")
            .Row = 27
            .Lock = True
        Else
            .Col = .ColLetterToNumber("I")
            .Row = 27
            .Lock = False
        End If
        'check thong tin ma so thue dai ly
        'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
        viTriMSTDL = "D_24~D_63~D_65"
        updateMSTDL fps, viTriMSTDL
        
        'set active vao cell tu ngay
        .Sheet = 1
        .SetActiveCell .ColLetterToNumber("L"), 4
        
        .EventEnabled(EventAllEvents) = True
    End With

End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_New.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
End Function

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim intIndexCombo As Integer
    Dim varTemp       As Variant, varTemp2 As Variant
    Dim tuThang, denThang As Variant
    ' get so thang quyet toan
    Dim soThangKG As Integer
    Dim CT22      As Variant, CT26 As Variant, varCT31 As Variant
    Dim varCT15   As Variant
    Dim mSheet    As Long
    
    Dim chiTieu22 As Variant
    
    With fps
    
        .EventEnabled(EventAllEvents) = False
        
        If Col = .ColLetterToNumber("I") And Row = 51 Then
            .GetText .ColLetterToNumber("I"), 51, varCT31
            If varCT31 <= 0 Then
                .SetText .ColLetterToNumber("I"), 52, 0
            End If
        End If
        
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("I") And Row = 52 Then
                .Formula = ""
            End If

            If Col = .ColLetterToNumber("I") And (Row = 47 Or Row = 48 Or Row = 49 Or Row = 50) Then
                .Col = .ColLetterToNumber("I")
                .Row = 52
                .Formula = "ROUND(R57,0)"
            End If
            
           

            
'            ' Set ct khi thay doi cac chi tieu 32,24,22  - 31072012
            If Col = .ColLetterToNumber("I") And Row = 57 Then
                .Col = .ColLetterToNumber("I")
                .Row = 42
                .GetText .Col, .Row, chiTieu22
                If Val(chiTieu22) > 0 Then
                    .Col = .ColLetterToNumber("I")
                    .Row = 57
                    .Formula = ""
                Else
                    .Col = .ColLetterToNumber("I")
                    .Row = 57
                    .Formula = "ROUND(S57,0)"
                End If
            End If
            
            If Col = .ColLetterToNumber("I") And (Row = 52 Or Row = 44 Or Row = 42) Then
                .Col = .ColLetterToNumber("I")
                .Row = 57
                .Formula = "ROUND(S57,0)"
            End If
            
'            If Col = .ColLetterToNumber("I") And (Row = 52 Or Row = 44 Or Row = 42) Then
'                .Col = .ColLetterToNumber("I")
'                .Row = 57
'                .Formula = "ROUND(S57,0)"
'            End If
'            ' end
            
        End If

        If .ActiveSheet = 2 Then
            If .Col = .ColLetterToNumber("I") And (Row = 17 Or Row = 19) Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 52
                .Formula = "ROUND(R57,0)"
            End If
            
            ' Set ct khi thay doi cac chi tieu 23,25  - 31072012
            If .Col = .ColLetterToNumber("I") And (Row = 17 Or Row = 18 Or Row = 19) Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 57
                .Formula = "ROUND(S57,0)"
            End If
            ' End
            
            
        End If

            'check CT15 <0 thi khong cho nhap CT16 tren PL 09B
            If .ActiveSheet = 3 Then
                .Sheet = 3
                fps.EventEnabled(EventAllEvents) = False
                If Col = .ColLetterToNumber("I") And Row = 26 Then
                    .GetText .ColLetterToNumber("I"), 26, varCT15
                    If Val(varCT15) < 0 Then
                        .Col = .ColLetterToNumber("I")
                        .Row = 27
                        .Text = "0"
                        .Lock = True
                    Else
                        .Col = .ColLetterToNumber("I")
                        .Row = 27
                        .Lock = False
                    End If

                End If
            ' Set ct khi thay doi cac chi tieu 23,25  - 31072012
            If .Col = .ColLetterToNumber("I") And (Row = 26 Or Row = 27 Or Row = 30) Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 57
                .Formula = "ROUND(S57,0)"
            End If
            ' End
                fps.EventEnabled(EventAllEvents) = True
            End If
        
            If .ActiveSheet = 4 Then
                If Col = .ColLetterToNumber("I") And .Text <> "" Then
                    intIndexCombo = .TypeComboBoxCurSel
                    .Row = Row
                    .Col = .ColLetterToNumber("H")
                    .TypeComboBoxCurSel = intIndexCombo
                    UpdateCell fps, .Col, .Row, .value
                End If
                fps.EventEnabled(EventAllEvents) = False
                If Col = .ColLetterToNumber("K") And Row = 4 Then
                    mSheet = .Sheet
                    'lay tong chi tieu 15 de so sanh
                    .GetText Col, Row, varTemp2
                    .Sheet = 1
                    .GetText .ColLetterToNumber("Q"), 48, varTemp
                    If Trim(varTemp) <> vbNullString Then
                        If Val(Trim(varTemp)) <> Val(Trim(varTemp2)) Then
                            .Col = .ColLetterToNumber("I")
                            .Row = 52
                            .Formula = "ROUND(R57,0)"
                        End If
                    End If
                    .Sheet = mSheet
                End If
                fps.EventEnabled(EventAllEvents) = True
            End If
            
        .EventEnabled(EventAllEvents) = True
        End With

    End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim intIndexCombo As Integer
    With fps
        If .ActiveSheet = 4 Then
            If Col = .ColLetterToNumber("I") And .Text <> "" Then
                intIndexCombo = .TypeComboBoxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("H")
                 .TypeComboBoxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
        End If
    End With
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
            If .ActiveSheet = 4 Then
                    fps.EventEnabled(EventAllEvents) = False
                    
                    .Col = .ColLetterToNumber("B")
                    i = 1
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
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

 Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
                          
    Dim varTemp As Variant
    Dim tuThang, denThang As Variant
    ' get so thang quyet toan
    Dim soThangQT As Integer, soThangKG As Integer
    Dim CT22      As Variant, CT26 As Variant
    
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet

        If .ActiveSheet = 1 Then

            'check tu thang
            If Col = .ColLetterToNumber("L") And Row = 4 Then
                .GetText Col, Row, varTemp

                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                        tuThang = varTemp
                        .GetText .ColLetterToNumber("N"), 4, denThang
                        
                        'check tu thang phai nam trong ky tinh thue
                        If Val(Right(CStr(tuThang), 4)) <> Val(TAX_Utilities_New.Year) Then
                            DisplayMessage "0252", msOKOnly, miWarning
'                            .SetFocus
                            .SetActiveCell Col, Row
                        Else
                            'check tu thang phai nho hon den thang
                            If (DateDiff("M", Format_mmyyyy(CStr(varTemp)), Format_mmyyyy(CStr(denThang))) + 1) <= 0 Then
                                DisplayMessage "0151", msOKOnly, miWarning
'                                .SetFocus
                                .SetActiveCell Col, Row
                            ElseIf Month(Format_mmyyyy(CStr(varTemp))) < 7 And TAX_Utilities_New.Year = 2009 Then
                                DisplayMessage "0152", msOKOnly, miWarning
'                                .SetFocus
                                .SetActiveCell Col, Row
                            Else
                                .Row = 4
                                .Col = .ColLetterToNumber("P")
                                .Text = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy(CStr(denThang))) + 1
                            End If
                        End If

                    Else
'                        .SetFocus
                        .SetActiveCell Col, Row
                    End If

                Else
                    DisplayMessage "0149", msOKOnly, miWarning
'                    .SetFocus
                    .SetActiveCell Col, Row
                    '.SetText Col, Row, ""
                End If

                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If

            ' check den thang
            If Col = .ColLetterToNumber("N") And Row = 4 Then
                .GetText Col, Row, varTemp

                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                        
                        denThang = varTemp
                        .GetText .ColLetterToNumber("L"), 4, tuThang

                        If (DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy(CStr(denThang))) + 1) <= 0 Then
                            DisplayMessage "0151", msOKOnly, miWarning
'                            .SetFocus
                            .SetActiveCell Col, Row
                        Else
                            .Row = 4
                            .Col = .ColLetterToNumber("P")
                            .Text = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy(CStr(denThang))) + 1
                            
                        End If

                    Else
'                        .SetFocus
                        .SetActiveCell Col, Row
                    End If

                Else
                    DisplayMessage "0150", msOKOnly, miWarning
                    '                    .SetFocus
                    .SetActiveCell Col, Row
                    '.SetText Col, Row, ""
                End If

                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If
            
            'set lai so thang tinh thue
            If (Col = .ColLetterToNumber("L") Or Col = .ColLetterToNumber("N")) And Row = 4 Then
                'set so thang tinh thue neu la nam 2011 thi co tinh so thang mien giam
                .GetText .ColLetterToNumber("L"), 4, tuThang
                .GetText .ColLetterToNumber("N"), 4, denThang

                'tinh so thang ke khai
                'ke khai nam 2011 se settext vao cell Q, <>2011 set vao cell P
                'ke khai nam 2009 thi so thang quyet toan = 6
                soThangKG = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1

                If TAX_Utilities_New.Year = "2011" Or TAX_Utilities_New.Year = "2012" Then
                    'tinh so thang ke khai
                    soThangKG = getSoThangQTMienGiam(CStr(tuThang), CStr(denThang))
                    .SetText .ColLetterToNumber("Q"), 4, soThangKG
                    
                    If isMienGiamPL26(CStr(tuThang), CStr(denThang)) Then
                        .SetText .ColLetterToNumber("T"), 51, 1
                    Else
                        .SetText .ColLetterToNumber("T"), 51, 0
                    End If
                ElseIf TAX_Utilities_New.Year = "2009" Then
                    .SetText .ColLetterToNumber("Q"), 4, "6"
                Else
                    .SetText .ColLetterToNumber("Q"), 4, "1"
                End If

                .Col = .ColLetterToNumber("P")
                .Row = 4
                .Text = DateDiff("M", Format(CStr(tuThang), "mm/yyyy"), Format(CStr(denThang), "mm/yyyy")) + 1
            End If
                        
            'check date
            If Col = .ColLetterToNumber("M") And Row = 65 Then
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
        
            If (Col = .ColLetterToNumber("L") Or Col = .ColLetterToNumber("N")) And Row = 4 Then
                .GetText .ColLetterToNumber("L"), 4, tuThang
                .GetText .ColLetterToNumber("N"), 4, denThang

                If DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) > 12 Then
                    DisplayMessage "0253", msOKOnly, miWarning
                    .SetActiveCell Col, Row
                Else
                    'get CT31 / so thang quyet toan
                    'kekhai nam 2009 thi so thang qt = 6
                    .GetText .ColLetterToNumber("Q"), 57, varTemp

                    If varTemp <= 5000000 And (TAX_Utilities_New.Year = "2011" Or TAX_Utilities_New.Year = "2012") Then
                        'tinh so thang ke khai
                        soThangQT = getSoThangQTMienGiam(tuThang, denThang)
                        .SetText .ColLetterToNumber("Q"), 4, soThangQT
                        If isMienGiamPL26(CStr(tuThang), CStr(denThang)) Then
                            .SetText .ColLetterToNumber("T"), 51, 1
                        Else
                            .SetText .ColLetterToNumber("T"), 51, 0
                        End If

'                    ElseIf TAX_Utilities_New.Year = "2009" Then
'                        .SetText .ColLetterToNumber("Q"), 4, "6"
'                    Else
'                        .SetText .ColLetterToNumber("Q"), 4, "1"
                    End If
                    .Col = .ColLetterToNumber("I")
                    .Row = 52
                    .Formula = "ROUND(R57,0)"
                End If

            End If
        End If
            
        
        If .ActiveSheet = 4 Then
            If Col = .ColLetterToNumber("E") And Row >= 22 Then
                       .GetText Col, Row, varTemp
                       If varTemp <> "" And varTemp <> "../../...." Then
                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                               '.SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                               If CDate(Format_ddmmyyyy(CStr(varTemp))) < CDate(Format_ddmmyyyy(Format(Date, "dd/mm/yyyy"))) Then
                                     .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                                Else
                                     DisplayMessage "0175", msOKOnly, miCriticalError
                                     .SetText Col, Row, ""
                                     .SetActiveCell Col, Row
                                End If
                           Else
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
        
        .EventEnabled(EventAllEvents) = True
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
'        SetData
        
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
    'chuoi vitriheader tren to khai, chuoi vi tri tren sheet header de check lai
    Dim viTriHeaderTk As String, viTriSheetHeader As String, mangMessage As String, xmlNode As MSXML.IXMLDOMNode, strTenNVDDL, strCCHN As String
    Dim MSTDTNT As Variant, tSheet As Integer
    Dim ct12   As Variant, cT40 As Variant, mErrCT12 As Variant, iFlagErrCT12 As Integer
    
    With fps
        CheckErrorMST
        'check thong tin header
        tSheet = .Sheet
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe
        viTriHeaderTk = "D_24~D_22~D_26~D_28~M_28~D_34~M_34~D_63~D_65"
        viTriSheetHeader = "B_68~B_69~B_70~B_71~B_72~B_73~B_74~B_75~B_76"
        mangMessage = "E_68~E_69~E_70~E_71~E_72~E_73~E_74~E_75~E_76"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        'dntai comment rang buoc nay tam thoi k dung
'        'check neu CT 40 > 0 thi bat buoc phai nhap CT 12
'        'get message loi
'        .Sheet = .SheetCount
'        .GetText .ColLetterToNumber("E"), 77, mErrCT12
'        .Sheet = 1
'        'get ct40 & ct 12
'        .GetText .ColLetterToNumber("I"), 60, cT40
'        .GetText .ColLetterToNumber("D"), 20, ct12
'        iFlagErrCT12 = 0
'        If cT40 > 0 And Trim(ct12) = vbNullString Then
'            .Col = .ColLetterToNumber("D")
'            .Row = 20
'            .CellNote = mErrCT12
'            .BackColor = mErrorColor
'            iFlagErrCT12 = 1
'        Else
'            .Col = .ColLetterToNumber("D")
'            .Row = 20
'            .CellNote = ""
'            .BackColor = mFormColor
'        End If
'        'set lai sheet header
'        .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), 77, IIf(iFlagErrCT12 = 1, "0", "1")
'        'end check ct40 nhap thi bat buoc nhap ct12
        
        'end check ct40 nhap thi bat buoc nhap ct12
        
        .Sheet = tSheet
        ' Neu ton tai file data phu luc 09C thi kiem tra tren phu luc 09C
        If IIf(TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet4
        Else
            clearCheckSheet4
        End If
        
    End With
End Sub

Private Sub checkSheet4()
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq As Variant
    Dim vCMT As Variant, vErrMsgCMT As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vNgaySinh As Variant, vErrMsgNgaySinh As Variant
    Dim vSoThangGiamTru As Variant, vErrMsgSoThangGiamTru As Variant
    Dim vQuanHeNNT As Variant, vErrMsgQuanHeNNT As Variant
    Dim vLastRow As Variant
    Dim vTuThang, vDenThang, vSoThangQuyetToan As Variant
    Dim vErrMSGSOCMND As Variant, vCMND As Variant, iFlagCMDN_HC As Integer
    Dim vErrEmpty As Variant, iFlagEmty As Integer
    Dim vErrMSTSoCMND As Variant, iFlagCheckEmptyMSTCMND As Integer
    Dim vErrCheckCT15, vCT15 As Variant, iFlagCheckCT15 As Integer
    Dim iFlagvHoVaTen As Integer, iFlagvMst As Integer, iFlagvNgaySinh As Integer, iFlagvSoThanhGiamTru As Integer, iFocusFlag As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim mErrNgaySinhPhuThuoc As Variant, iFlagNgaySinhPhuThuoc As Integer
    Dim mErrMSTCT5Trung As Variant, iFlagMSTCT5Trung As Integer, rowMSTTrung As Integer
    
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5 As Boolean
    With fps
        iCurrentSheet = .Sheet
        
        .Sheet = 1
        .GetText .ColLetterToNumber("L"), 4, vTuThang
        .GetText .ColLetterToNumber("N"), 4, vDenThang
        vSoThangQuyetToan = DateDiff("M", Format_mmyyyy(CStr(vTuThang)), Format_mmyyyy(CStr(vDenThang))) + 1
        
        .Sheet = .SheetCount
        
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 17, vErrMsgHoVaTen
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 18, vErrMsgNgaySinh
        ' Lay thong bao ve ko nhap ma so thue trong sheet header
        '.GetText .ColLetterToNumber("E"), 19, vErrMsgMSTReq
        
        ' Lay thong bao ve ko chon quan he voi NNT trong sheet header
        .GetText .ColLetterToNumber("E"), 19, vErrMsgQuanHeNNT
        ' Lay thong bao ve ko nhap ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 20, vErrMsgCMT
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 21, vErrMsgMST
        ' Lay thong bao ve sai ve so thang duoc giam tru thue cua nguoi phu thuoc trong sheet header
        .GetText .ColLetterToNumber("E"), 22, vErrMsgSoThangGiamTru
        'lay thong bao ve nhap sai cmnd/ho chieu khi nguoi dung nhap dau cach
        .GetText .ColLetterToNumber("E"), 61, vErrMsgCMT
        'lay thong bao ve truong bat buoc phai nhap
        .GetText .ColLetterToNumber("E"), 62, vErrEmpty
        'lay thong bao ve truong nhap ma so thue hoac so cmnd
        .GetText .ColLetterToNumber("E"), 63, vErrMSTSoCMND
        'lay thong bao ve chi tieu 15 < chi tieu 14 * 1600000
        .GetText .ColLetterToNumber("E"), 64, vErrCheckCT15
        'lay thong bao ve ngay sinh cua nguoi phu thuoc phai truoc ky tinh thue
        .GetText .ColLetterToNumber("E"), 66, mErrNgaySinhPhuThuoc
        'lay thong bao ve MST cua vo hoac chong trung voi MST  nguoi phu thuoc
        .GetText .ColLetterToNumber("E"), 67, mErrMSTCT5Trung
        
        .Sheet = 4 'Check tinh xac thuc cua cac thong tin tren bang ke 05B/TNCN
        
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        iFlagNgaySinhPhuThuoc = 0
        iFlagMSTCT5Trung = 0
                
        ' Lay ma so thue tren phu luc 09C/TNCN chi tieu D11 (Ma so thue cua vo hoac chong)
        .GetText .ColLetterToNumber("E"), 11, vMST
        'lay ho va ten nguoi phu thuoc
        .GetText .ColLetterToNumber("D"), 22, vHoVaTen
        ' Check ma so thue tren phu luc 09C/TNCN chi tieu D11 (Ma so thue cua vo hoac chong)
        If Trim(vMST) <> vbNullString Then
            vMST = Left(vMST, 10)
           strCheck = ""
                   
           If Len(vMST) = 10 Or Len(vMST) = 13 Then
                strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
           End If
           If Mid(strCheck, 2, 1) <> "0" Then
                .Col = .ColLetterToNumber("E")
                .Row = 11
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgMST   'static
                .BackColor = mErrorColor
                iFlagvMst = 1
                blCheck_S5 = False
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
            Else
                If Trim(vHoVaTen) <> vbNullString Then
                    rowMSTTrung = .SearchCol(.ColLetterToNumber("F"), 21, -1, vMST, SearchFlagsNone)
                    If rowMSTTrung <> -1 Then
                        .Col = .ColLetterToNumber("E")
                        .Row = 11
                        .CellNote = mErrMSTCT5Trung
                        .BackColor = mErrorColor
                        iFlagMSTCT5Trung = 1
                    Else
                        .Col = .ColLetterToNumber("E")
                        .Row = 11
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                 .Col = .ColLetterToNumber("E")
                 .Row = 11
                 .CellNote = ""
                 .BackColor = mNonErrorColor
                End If
            End If
        Else
            .Col = .ColLetterToNumber("E")
            .Row = 11
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
        'dntai 25/08/2011
        'Lay gia so cmnd , so ho chieu nguoi dung nhap vao
        .GetText .ColLetterToNumber("E"), 13, vCMND
        If InStr(1, CStr(vCMND), " ") <> 0 Then
            .Col = .ColLetterToNumber("E")
            .Row = 13
            .BackColor = mErrorColor
            .CellNote = vErrMsgCMT
            iFlagCMDN_HC = 1
        Else
            .Col = .ColLetterToNumber("E")
            .Row = 13
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        'end
        
        For j = 22 To .MaxRows
            .Row = j
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vNgaySinh
            .GetText .ColLetterToNumber("F"), .Row, vMST
            .GetText .ColLetterToNumber("G"), .Row, vCMT
            .GetText .ColLetterToNumber("I"), .Row, vQuanHeNNT
            .GetText .ColLetterToNumber("J"), .Row, vSoThangGiamTru
            .GetText .ColLetterToNumber("K"), .Row, vCT15
            
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            'reset du lieu loi
            .Col = .ColLetterToNumber("F")
            .CellNote = ""
            .BackColor = mNonErrorColor

            
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            
            If Trim(vHoVaTen) = vbNullString And Trim(vNgaySinh) = vbNullString And Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString Then
            Else
                'Check Ho va ten tren phu luc 09C/TNCN
                If Trim(vHoVaTen) = vbNullString Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgHoVaTen    'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S1 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("D")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'Check Ngay sinh tren phu luc 09C/TNCN
                If Trim(vNgaySinh) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgNgaySinh     'static
                    .BackColor = mErrorColor
                    iFlagvNgaySinh = 1
                    blCheck_S2 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("E")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'check ngay sinh phai sau ky tinh thue
                If Val(Right(CStr(Trim(vNgaySinh)), 4)) > Val(TAX_Utilities_New.Year) Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = mErrNgaySinhPhuThuoc
                    iFlagNgaySinhPhuThuoc = 1
'                Else
'                    .Col = .ColLetterToNumber("E")
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
                End If
                
                'check khong dc de trong dong thoi 2 truong MST va so CMND
'                If Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString Then
'                    .Col = .ColLetterToNumber("F")
'                    .BackColor = mErrorColor
'                    .CellNote = vErrMSTSoCMND
'                    .Col = .ColLetterToNumber("G")
'                    .BackColor = mErrorColor
'                    .CellNote = vErrMSTSoCMND
'                    iFlagCheckEmptyMSTCMND = 1
'                Else
'                    .Col = .ColLetterToNumber("F")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
                
                
                
                'Check ma so thue tren phu luc 09C/TNCN

                If Trim(vMST) <> vbNullString Then
'                    vMST = Left(vMST, 10)
                   
                   strCheck = ""
                   If Len(vMST) = 14 And InStr(1, vMST, "-") = 11 Then
                    vMST = Replace(vMST, "-", "")
                   End If
                   
                   If Len(vMST) = 10 Or Len(vMST) = 13 And InStr(1, vMST, "-") = 0 Then
                        strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                   End If
                   If Mid(strCheck, 2, 1) <> "0" Then
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMST   'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        blCheck_S3 = False
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
                         
                         .Col = .ColLetterToNumber("G")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
'                ElseIf (Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString) Then
'                    .Col = .ColLetterToNumber("F")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTReq    'static
'                    .BackColor = mErrorColor
'
'                    .Col = .ColLetterToNumber("G")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgCMT    'static
'                    .BackColor = mErrorColor
'
'
'                    iFlagvMst = 1
'                    blCheck_S3 = False
'                    'capture the first error cell
'                    If iFocusFlag = 0 Then
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        iFocusFlag = 1
'                    End If
                Else
                        ' check mst k dc de trong
'                        .Col = .ColLetterToNumber("F")
'                        .CellNote = vErrEmpty
'                        .BackColor = mErrorColor
'                        iFlagEmty = 1
                End If
                
                
                
                'check SoCMND khong dc co dau cach
                If Trim(vCMT) <> vbNullString Then
                    If InStr(1, CStr(vCMT), " ") <> 0 Then
                        .Col = .ColLetterToNumber("G")
                        .BackColor = mErrorColor
                        .CellNote = vErrMsgCMT
                        iFlagCMDN_HC = 1
                    Else
                        .Col = .ColLetterToNumber("G")
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                    End If
                End If
                ' Check quan he voi NNT
                'Check Ho va ten tren phu luc 09C/TNCN
                If Trim(vQuanHeNNT) = vbNullString Then
                    .Col = .ColLetterToNumber("I")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgQuanHeNNT     'static
                    .BackColor = mErrorColor
                    blCheck_S3 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("I")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'check chi tieu [15] <= chitieu [14] * 1600000
                vCT15 = Replace(CStr(vCT15), ".", "")
                If (Val(vCT15) > Val(vSoThangGiamTru) * 1600000) Then
                    .Col = .ColLetterToNumber("K")
                    .BackColor = mErrorColor
                    .CellNote = vErrCheckCT15
                    iFlagCheckCT15 = 1
                Else
                    .Col = .ColLetterToNumber("K")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                'Check So thang duoc giam tru thue cua nguoi phu thuoc tren phu luc 09C/TNCN
                
                If IIf(vSoThangGiamTru = vbNullString, 0, vSoThangGiamTru) > Val(vSoThangQuyetToan) Then
                    .Col = .ColLetterToNumber("J")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThangGiamTru    'static
                    .BackColor = mErrorColor
                    iFlagvSoThanhGiamTru = 1
                    blCheck_S4 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("J")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
            
        'ket thuc vong lap
        Next
        ' Trong truong hop blCheck_S1 khong co Ho va ten nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "0"
        End If
        ' Trong truong hop blCheck_S2 khong co Ngay sinh bi loi thi set lai gia tri cua Header la 1
        If blCheck_S2 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "0"
        End If
        ' Trong truong hop blCheck_S3 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S3 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "1"
            .SetText .ColLetterToNumber("B"), 20, "1"
            .SetText .ColLetterToNumber("B"), 21, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "0"
            .SetText .ColLetterToNumber("B"), 20, "1"
            .SetText .ColLetterToNumber("B"), 21, "0"
            
        End If
        ' Trong truong hop NSD k nhap mst
        If iFlagEmty = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 62, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 62, "1"
        End If
        ' Trong truong hop blCheck_S4 so thang duoc giam tru ko sai thi set lai gia tri cua Header la 1
        If blCheck_S4 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 22, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 22, "0"
        End If
        
        ' Trong truong hop blCheck_S5 ma so thue cua vo hoac chong ko sai thi set lai gia tri cua Header la 1
        If blCheck_S5 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 23, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 23, "0"
        End If
        
        ' Trong truong hop nguoi su dung nhap dau cach khi nhap cmnd/ ho chieu thi set lai gia tri tren Header
        If iFlagCMDN_HC = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 61, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 61, "1"
        End If
        'Trong truong hop nguoi su dung khong nhap ca MST va so CMND thi set lai gia tri tren Header
        If iFlagCheckEmptyMSTCMND = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 63, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 63, "1"
        End If
        'trong truong hop ct [15] > ct [14] * 1600000 thi set lai gia tri tren header
        If iFlagCheckCT15 = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 64, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 64, "1"
        End If
        ' Trong truong hop Ngay sinh nguoi phu thuoc sau ky tinh thue set lai header
        If iFlagNgaySinhPhuThuoc = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 66, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 66, "1"
        End If
       ' Trong truong hop MST cha me trung voi mst nguoi phu thuoc thi set lai header
        If iFlagMSTCT5Trung = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 67, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 67, "1"
        End If
        .Sheet = iCurrentSheet
    End With
End Sub

Private Sub clearCheckSheet4()
    With fps
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 17, "1"
        .SetText .ColLetterToNumber("B"), 18, "1"
        .SetText .ColLetterToNumber("B"), 19, "1"
        .SetText .ColLetterToNumber("B"), 20, "1"
        .SetText .ColLetterToNumber("B"), 21, "1"
        .SetText .ColLetterToNumber("B"), 22, "1"
        .SetText .ColLetterToNumber("B"), 27, "1"
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
        .Row = 3
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
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode, xmlCellNodeChild As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "09A"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and I column
'                    If (lCol = fps.ColLetterToNumber("I") And lRow = 16) Or (lCol = fps.ColLetterToNumber("I") And lRow = 18) Or (lCol = fps.ColLetterToNumber("I") And lRow = 20) Or (lCol = fps.ColLetterToNumber("I") And lRow = 23) Or (lCol = fps.ColLetterToNumber("I") And lRow = 26) Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next

                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If

             Case "09B"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and I column
'                    If (lCol = fps.ColLetterToNumber("I") And lRow = 11) Or (lCol = fps.ColLetterToNumber("I") And lRow = 17) _
'                            Or (lCol = fps.ColLetterToNumber("I") And lRow = 26) Or (lCol = fps.ColLetterToNumber("I") And lRow = 27) _
'                                    Or (lCol = fps.ColLetterToNumber("I") And lRow = 28) Or (lCol = fps.ColLetterToNumber("I") And lRow = 29) Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next

                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If

            Case "09C"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(3).getElementsByTagName("Cell")
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
            Case "26MT"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(4).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("I") Then
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
    Set xmlCellNodeChild = Nothing
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim isheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
            ' set gia tri CT8 tren 09C
            isheet = mCurrentSheet
            .Sheet = 1
            .SetText .ColLetterToNumber("D"), 20, ""
            UpdateCell fps, .ColLetterToNumber("D"), 20, ""
            .SetText .ColLetterToNumber("M"), 20, ""
            UpdateCell fps, .ColLetterToNumber("M"), 20, ""
            mCurrentSheet = isheet
        .EventEnabled(EventAllEvents) = True
    End With
End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
Dim isheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
            ' set gia tri CT8 tren 09C
            isheet = mCurrentSheet
            .Sheet = 4
            .SetText .ColLetterToNumber("B"), 22, 1
            mCurrentSheet = isheet
        .EventEnabled(EventAllEvents) = True
    End With
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
        
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'               If lCol = .ColLetterToNumber("P") And lRow = 52 Then
'
'               Else
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
'               End If
             Next
        End If
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(1), "Active") <> "0") Then
        .Sheet = 2
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(2), "Active") <> "0") Then
        .Sheet = 3
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(2).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(3), "Active") <> "0") Then
        .Sheet = 4
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(3).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(4), "Active") <> "0") Then
        .Sheet = 5
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(4).getElementsByTagName("Cell")
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

Public Sub FinishImport()

End Sub


'Private Function getSoThangQTMienGiam(ByVal tuThang As String, ByVal denThang As String) As Integer
'    Dim soThang As Integer
'    Dim tuThangTmp As Variant, denThangTmp As Variant
'    Dim hieuTuThang As Integer, hieuDenThang As Integer
'    'set Thang mien giam thue tu thang 08/2011 den thang 12/2011
'    tuThangTmp = "08/2011"
'    denThangTmp = "12/2011"
'    'set so Thang quyet toan Thue
'    soThang = 1
'    hieuTuThang = DateDiff("M", Format(tuThangTmp, "mm/yyyy"), Format(tuThang, "mm/yyyy"))
'    hieuDenThang = DateDiff("M", Format(denThangTmp, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    If hieuTuThang < 0 And hieuDenThang <= 0 And DateDiff("M", Format(tuThangTmp, "mm/yyyy"), Format(denThang, "mm/yyyy")) >= 0 Then ' tinh so thang quyet toan trong TH : tuthang < 08/2011 va den thang < 12/2011
'        denThang = tuThangTmp
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    ElseIf hieuTuThang >= 0 And DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThangTmp, "mm/yyyy")) >= 0 And hieuDenThang > 0 Then ' tinh so thang quyet toan trong TH : tuThang > 08/2011 va denThang > 12/2011
'        tuThang = denThangTmp
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    ElseIf hieuTuThang < 0 And hieuDenThang > 0 Then  ' tinh so thang quyet toan trong TH : tu Thang < 08 / 2011 va denThang > 12/2011
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(tuThangTmp, "mm/yyyy"))
'        soThang = soThang + DateDiff("M", Format(denThangTmp, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    ElseIf hieuTuThang >= 0 And hieuDenThang <= 0 Then   ' tinh so thang quyet toan trong TH : thThang > 08/2011 va denThang < 12/2011
'        soThang = 0
'    Else   ' tinh so thang trong truong hop thang ke khai khong nam trong khoang thang mien giam
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1
'    End If
'
'    ' tinh so thang mien giam cho 2012
'
'
'
'    getSoThangQTMienGiam = soThang
'End Function


Private Function getSoThangQTMienGiam(ByVal tuThang As String, ByVal denThang As String) As Integer
    Dim soThang As Integer
    'set so Thang quyet toan Thue
    soThang = 0

    ' tinh so thang mien giam cho 2012
    Dim arrDanhsach() As String
    Dim thangMg As String
    Dim countMonth As Integer
    Dim tempDate As Date
    Dim strDataFileName As String
    Dim i As Integer
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    countMonth = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1
    If countMonth > 1 Then
        ReDim arrDanhsach(countMonth - 1)
        arrDanhsach(0) = tuThang
        For i = 1 To countMonth - 1
            tempDate = DateAdd("M", i, Format(tuThang, "mm/yyyy"))
            arrDanhsach(i) = Format(tempDate, "mm/yyyy")
        Next
    Else
        ReDim arrDanhsach(1)
        arrDanhsach(0) = tuThang
    End If

    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_Thang_MT.xml")
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                thangMg = Trim$(GetAttribute(xmlNode, "Value"))
                For i = 0 To countMonth - 1
                    If thangMg = arrDanhsach(i) Then
                        soThang = soThang + 1
                        Exit For
                    End If
                Next
            End If
        Next
    End If


    getSoThangQTMienGiam = soThang
End Function

Private Sub fps_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    With fps
    
'        If OldSheet = 3 Then
'
'        End If
        If NewSheet = 1 And OldSheet = 3 Then
            .Sheet = 3
            .GetText .ColLetterToNumber("I"), 33, varTemp1
            .Sheet = 1
            .GetText .ColLetterToNumber("R"), 56, varTemp2
            If (varTemp2 <> varTemp1) Then
                .Col = .ColLetterToNumber("I")
                .Row = 52
                .Formula = "ROUND(R57,0)"
                .SetText .ColLetterToNumber("R"), 56, varTemp1
            End If
        End If
        If NewSheet = 4 Then
            .Sheet = 1
            .GetText .ColLetterToNumber("I"), 48, varTemp1
            .SetText .ColLetterToNumber("Q"), 48, varTemp1
            .Sheet = NewSheet
        End If
    End With
End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 20
    denRowDL = 25
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "10_12_2~10_12_3~18_20_4"
    ' end
End Sub


' Ham kiem tra thang Mien giam thuoc 7/2012 - 12/2012
Private Function isMienGiamPL26(ByVal tuThang As String, ByVal denThang As String) As Boolean
    Dim arrThang() As String
    Dim result As Boolean
    arrThang = Split(tuThang, "/")
    result = False
    ' kiem tra tu thang thuoc 7-12/2012
    If Val(arrThang(1)) = 2012 Then
        If Val(arrThang(0)) >= 7 And Val(arrThang(0)) <= 12 Then
            result = True
        End If
    End If
    
    ' kiem tra den thang thuoc 7-12/2012
    arrThang = Split(denThang, "/")
     If Val(arrThang(1)) = 2012 Then
        If Val(arrThang(0)) >= 7 And Val(arrThang(0)) <= 12 Then
            result = True
        End If
    End If
    
    isMienGiamPL26 = result
End Function
