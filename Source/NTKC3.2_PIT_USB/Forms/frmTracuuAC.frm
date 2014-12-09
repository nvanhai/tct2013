VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTraCuuAC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnThoat 
      Caption         =   "&Tho¸t"
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton btnTraCuu 
      Caption         =   "Tra &cøu"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin FPUSpreadADO.fpSpread fpsKetQua 
      Height          =   4050
      Left            =   360
      TabIndex        =   2
      Top             =   1830
      Width           =   10335
      _Version        =   458752
      _ExtentX        =   18230
      _ExtentY        =   7144
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   8
      MaxRows         =   13
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   14
      SpreadDesigner  =   "frmTracuuAC.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän ngµy nép"
      Height          =   1215
      Left            =   5760
      TabIndex        =   5
      Top             =   360
      Width           =   5055
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   585
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _Version        =   458752
         _ExtentX        =   8070
         _ExtentY        =   1032
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   7
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmTracuuAC.frx":0584
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   1590
      Width           =   10575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chän lo¹i b¸o c¸o Ên chØ"
      Height          =   1185
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   5415
      Begin VB.OptionButton OptToLoi 
         Caption         =   "Tê khai lçi"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5280
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.OptionButton optRecv 
         Caption         =   "Ch­a nhËn sang QLAC"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   2145
      End
      Begin VB.OptionButton optQLT 
         Caption         =   "NhËn sang QLAC"
         Height          =   255
         Left            =   750
         TabIndex        =   10
         Top             =   840
         Value           =   -1  'True
         Width           =   1845
      End
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   525
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   4965
         _Version        =   458752
         _ExtentX        =   8758
         _ExtentY        =   926
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmTracuuAC.frx":0A81
         UserResize      =   1
      End
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1965
      ForeColor       =   -2147483634
      Size            =   "3466;450"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   315
      Left            =   120
      Top             =   0
      Width           =   9195
   End
   Begin MSForms.Label LblSoBG 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   6120
      Width           =   3015
      BackColor       =   -2147483648
      VariousPropertyBits=   8388627
      Size            =   "5318;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTraCuuAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const f1dteTuNRow = 2
Private Const f1dteTuNCol = "C"
Private Const f1dteDeNRow = 2
Private Const f1dteDeNCol = "F"
Private Const f4cboLoTRow = 2
Private Const f4cboLoTCol = "C"
Private Const mFormColor = -2147483633
Private Const lminDate = "01/01/1900"
Private Const lmaxDate = "31/12/3000"
Private Const lmaxSoTk = 1000

Private lngRowFocus As Long
Private DteTuN As Date
Private DteDeN As Date
Private lmaTK As Long
Private lSoBG As Long
Private larrId(lmaxSoTk) As Long
Private lerror As Boolean
Private strDaNhan As String
Private Sub btnThoat_Click(Index As Integer)
    Unload Me
End Sub
Private Sub btnTraCuu_Click(Index As Integer)
    
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    Dim strMST As String
    'Xoa bo ket qua cu tren Grid
    fpsKetQua.ClearRange 1, 1, fpsKetQua.MaxCols, lSoBG, True
    If Not KiemTraDKngay Then
         Exit Sub
    End If
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    'Lay cau lenh truy van
    xmlSQL.Load App.path & "\SQL.xml"
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraCuuAc")
    strSQL = Replace(strSQL, "nhohon=", "<=")
    strSQL = Replace(strSQL, "ma_tkhai", "" & changeLoaiToKhai(lmaTK) & "")
    strSQL = Replace(strSQL, "strDa_Nhan", "" & strDaNhan & "")
    strSQL = Replace(strSQL, "ngay_nop_dau", "To_date('" & format(DteTuN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    strSQL = Replace(strSQL, "ngay_nop_cuoi", "To_date('" & format(DteDeN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    'Khoi tao bien
    lCol = 2
    lRow = 1
    lSoBG = 0
    'Thuc hien cau lenh sql
    Set rsReturn = clsDAO.Execute(strSQL)
    'Hien thi du lieu len grid
    If rsReturn.Fields.Count > 0 Then
        Do While Not rsReturn.EOF
            fpsKetQua.MaxRows = lRow + 1
            fpsKetQua.InsertRows lRow, 1
            fpsKetQua.SetText 1, lRow, lSoBG + 1
            For lIndex = 1 To rsReturn.Fields.Count
                If Not (IsNull(rsReturn.Fields(lIndex - 1).Value)) Then
                    If UCase(Trim(rsReturn.Fields(lIndex - 1).Name)) = "TIN" Then
                        strMST = Trim(rsReturn.Fields(lIndex - 1).Value)
                    End If
                    If UCase(Trim(rsReturn.Fields(lIndex - 1).Name)) = "NGUOI_DAI_DIEN" Then
                        fpsKetQua.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(getTENDTNT(strMST), TCVN, UNICODE)
                    Else
                        fpsKetQua.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
                    End If
                    lCol = lCol + 1
                    fpsKetQua.RowHeight(lRow) = fpsKetQua.MaxTextRowHeight(lRow)
                Else
                    fpsKetQua.SetText lCol, lRow, rsReturn.Fields(lIndex - 1).Value
                    lCol = lCol + 1
                    fpsKetQua.RowHeight(lRow) = fpsKetQua.MaxTextRowHeight(lRow)
                End If
            Next lIndex
            lCol = 2
            lRow = lRow + 1
            
            rsReturn.MoveNext
            lSoBG = lSoBG + 1
        Loop
        
    Else
        lSoBG = 0
        DisplayMessage "0072", msOKOnly
    End If
    LblSoBG.Visible = True
    LblSoBG.TextAlign = fmTextAlignLeft
    LblSoBG.caption = "Sè tê khai t×m thÊy: " & lSoBG
    'Chinh sua lai Grid ket cho dep
    With fpsKetQua
        If lSoBG <= 15 Then
            Dim i As Integer
            For i = 1 To 15 - lRow
                .MaxRows = lSoBG + i + 1
                .InsertRows lSoBG + i, 1
            Next
        Else
            .RowHeight(lRow) = 0
        End If
        
        If lngRowFocus = 0 Then
            lngRowFocus = SetRowFocus(1, 1, True)
        ElseIf lngRowFocus > .MaxRows Then
            lngRowFocus = SetRowFocus(1, .MaxRows, True)
        Else
            lngRowFocus = SetRowFocus(1, lngRowFocus, True)
        End If
        .SetFocus
    End With

End Sub

Private Function changeLoaiToKhai(ByVal strLoaiMaToKhai As String) As String
    If strLoaiMaToKhai = "102" Then changeLoaiToKhai = "'%GTGT%'"
    If strLoaiMaToKhai = "103" Then changeLoaiToKhai = "'%TNDN%'"
    If strLoaiMaToKhai = "104" Then changeLoaiToKhai = "'%TNCN%'"
    If strLoaiMaToKhai = "105" Then changeLoaiToKhai = "'%TAIN%'"
    If strLoaiMaToKhai = "106" Then changeLoaiToKhai = "'%TTDB%'"
    If strLoaiMaToKhai = "101" Then changeLoaiToKhai = "'%NTNN%'"
    
    If strLoaiMaToKhai = "108" Then changeLoaiToKhai = "'%15%'"
    If strLoaiMaToKhai = "109" Then changeLoaiToKhai = "'%48%'"
    If strLoaiMaToKhai = "110" Then changeLoaiToKhai = "'%16%'"
    If strLoaiMaToKhai = "111" Then changeLoaiToKhai = "'%99%'"
    If strLoaiMaToKhai = "64" Then changeLoaiToKhai = "'01_TBAC'"
    If strLoaiMaToKhai = "18" Then changeLoaiToKhai = "'BC26_AC_SL'"
    If strLoaiMaToKhai = "65" Then changeLoaiToKhai = "'01_AC'"
    If strLoaiMaToKhai = "66" Then changeLoaiToKhai = "'BC21_AC'"
    If strLoaiMaToKhai = "67" Then changeLoaiToKhai = "'03_TBAC'"
    If strLoaiMaToKhai = "91" Then changeLoaiToKhai = "'04_TBAC'"
    If strLoaiMaToKhai = "68" Then changeLoaiToKhai = "'BC26_AC'" 'fix (same %BC26_AC_SL%)
    If strLoaiMaToKhai = "27" Then changeLoaiToKhai = "'01_BK_BC26_AC'"
    If (strLoaiMaToKhai = "07" Or strLoaiMaToKhai = "7") Then changeLoaiToKhai = "'01_TBAC_BLP'"
    If strLoaiMaToKhai = "13" Then changeLoaiToKhai = "'%01_AC_BLP%'"
    If (strLoaiMaToKhai = "09" Or strLoaiMaToKhai = "9") Then changeLoaiToKhai = "'BC21_AC_BLP'"
    If strLoaiMaToKhai = "14" Then changeLoaiToKhai = "'BC26_AC_BLP'"
    If strLoaiMaToKhai = "10" Then changeLoaiToKhai = "'03_TBAC_BLP'"
    If strLoaiMaToKhai = "0" Then changeLoaiToKhai = "'%'"
End Function


Private Sub Form_Load()
    SetControlCaption Me, "frmTraCuuAC"
    FormatGrid
    SetupData
    'lerror = False
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
End Sub
Sub FormatGrid()
    With fpsDkNgay
        .BackColor = mFormColor
      'fpSpread1.BorderStyle = BorderStyleNone
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        '.BorderStyle = BorderStyleNone
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        .BackColor = vbWhite
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        .BackColor = vbWhite
    End With
    With fpsKetQua
        .EditModePermanent = True
        .EditModeReplace = True
        .ColHeadersShow = True
        .RowHeadersShow = False
        .AllowColMove = False
        .MaxRows = 17
        .TextTip = TextTipFloating
        .CursorType = CursorTypeLockedCell
        .CursorStyle = CursorStyleArrow
        .SetTextTipAppearance "Tahoma", 8, False, False, RGB(255, 255, 235), &H0
        
    End With
    With fpsLoaiTK
        .BackColor = mFormColor
        'fpSpread4.BorderStyle = BorderStyleNone
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        .Col = .ColLetterToNumber(f4cboLoTCol)
        .Row = f4cboLoTRow
        .BackColor = vbWhite
    End With
    LblSoBG.Visible = False
    
End Sub
Sub SetupData()
        With fpsDkNgay
            Dim vdtehientai As Variant
            vdtehientai = Date
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .SetText .Col, .Row, vdtehientai
            
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .SetText .Col, .Row, vdtehientai
        End With
    'Lay du lieu cho cbo
        With fpsLoaiTK
            .Col = .ColLetterToNumber(f4cboLoTCol)
            .Row = f4cboLoTRow
            Dim xmlDocument As New MSXML.DOMDocument
            Dim xmlNode As MSXML.IXMLDOMNode
            Dim strDataFileName As String
            Dim i As Integer
            xmlDocument.Load TAX_Utilities_Srv_New.GetAbsolutePath("menu.xml")
            Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            i = 0
            larrId(0) = 0
            For Each xmlNode In xmlNodeListMenu
                Dim LoaiTk As String
                Dim Parentid As String
                Parentid = GetAttribute(xmlNode, "ParentID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                '.TypeComboBoxIndex = 0
                If Parentid = "112" Or Parentid = "114" Then
                    i = i + 1
                    .TypeComboBoxIndex = -1
                    .TypeComboBoxString = LoaiTk
                    larrId(i) = Val(GetAttribute(xmlNode, "ID"))
                End If
            Next
            .TypeComboBoxCurSel = 0
            Set xmlNode = Nothing
            Set xmlDocument = Nothing
        End With
        strDaNhan = "='Y'"
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    frmTraCuuAC.Top = (frmSystem.Height - frmTraCuuAC.Height) / 4
    frmTraCuuAC.Left = (frmSystem.Width - frmTraCuuAC.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lngRowFocus = 0
End Sub

Private Sub fpsDkNgay_GotFocus()
    'btnTraCuu(0).Default = True
End Sub
Private Sub fpsDkNgay_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab And Shift = 0 Then 'And Not lerror
    If fpsDkNgay.ActiveRow = f1dteDeNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteDeNCol) Then
            fpsKetQua.SetFocus
    End If
End If
If KeyCode = vbKeyTab And Shift = 1 Then
    If fpsDkNgay.ActiveRow = f1dteTuNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteTuNCol) Then
        fpsLoaiTK.SetFocus
        With fpsLoaiTK
            .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
        End With
    End If
    If fpsDkNgay.ActiveRow = f1dteDeNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteDeNCol) Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteTuNCol), f1dteTuNRow
        End With
    End If
End If
End Sub
Private Sub fpsDkNgay_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With fpsDkNgay
         .Col = .ColLetterToNumber(f1dteTuNCol)
         .Row = f1dteTuNRow
           Dim vdtetun As Variant
            .GetText .Col, .Row, vdtetun
            If Not IsDate(vdtetun) And vdtetun <> "" Then
                .SetText .Col, .Row, ""
                DisplayMessage "0073", msOKOnly, miInformation
                Cancel = True
                'lerror = True
                .SetActiveCell .Col, .Row
                Exit Sub
            End If
            
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
           Dim vdteden As Variant
            .GetText .Col, .Row, vdteden
            If Not IsDate(vdteden) And vdteden <> "" Then
                .SetText .Col, .Row, ""
                DisplayMessage "0073", msOKOnly, miInformation
                Cancel = True
                'lerror = True
                .SetFocus
                .SetActiveCell .Col, .Row
                'lerror = True
                Exit Sub
            End If
    End With
    'lerror = False
End Sub

Private Sub fpsKetQua_Click(ByVal Col As Long, ByVal Row As Long)
    With fpsKetQua
        .ReDraw = False
        lngRowFocus = SetRowFocus(lngRowFocus, Row, True)
        .ReDraw = True
    End With
End Sub

Private Sub fpsKetQua_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 1 Then
       fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteDeNCol), f1dteDeNRow
        End With
    End If
    If KeyCode = vbKeyTab And Shift = 0 Then
       btnTraCuu(0).SetFocus
    End If
    If KeyCode = vbKeyDown And lngRowFocus < fpsKetQua.MaxRows Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus + 1)
    ElseIf KeyCode = vbKeyUp And lngRowFocus > 1 Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus - 1)
    End If
End Sub

Private Sub fpsLoaiTK_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    With fpsLoaiTK
        lmaTK = larrId(.TypeComboBoxCurSel)
    End With
End Sub
Private Sub fpsLoaiTK_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 0 Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteTuNCol), f1dteTuNRow
        End With
    End If
    If KeyCode = vbKeyTab And Shift = 1 Then
        btnThoat(1).SetFocus
    End If
End Sub
Function KiemTraDKngay() As Boolean
With fpsDkNgay
        KiemTraDKngay = True
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        Dim vdtetun As Variant
        .GetText .Col, .Row, vdtetun
        If vdtetun = "" Then
            DteTuN = format(lminDate, "DD/MM/YYYY")
        Else
            If IsDate(vdtetun) Then
                DteTuN = format(vdtetun, "DD/MM/YYYY")
            Else
                KiemTraDKngay = False
                Exit Function
            End If
        End If
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        Dim vdteden As Variant
        .GetText .Col, .Row, vdteden
        If vdteden = "" Then
            DteDeN = format(lmaxDate, "DD/MM/YYYY")
        Else
            If IsDate(vdteden) Then
                DteDeN = format(vdteden, "DD/MM/YYYY")
            Else
                KiemTraDKngay = False
                Exit Function
            End If
        End If
        If DteTuN > DteDeN Then
            DisplayMessage "0071", msOKOnly
            KiemTraDKngay = False
            .SetActiveCell .ColLetterToNumber(f1dteDeNCol), f1dteDeNRow
        End If
    End With
End Function

Private Function SetRowFocus(ByVal lngRow As Long, ByVal lngNewRow As Long, Optional ByVal blnClickEvent As Boolean = False) As Long
    With fpsKetQua
        .Col = -1
        .Row = lngRow
        .BackColor = vbWhite
        
        .Row = lngNewRow
        .BackColor = RGB(212, 343, 423)
        
        If blnClickEvent Then
            .SetActiveCell 8, lngNewRow
        Else
            .SetActiveCell 8, lngRow
        End If
        
    End With
    SetRowFocus = lngNewRow
End Function

Private Sub optQLT_Click()
    strDaNhan = " = 'Y'"
End Sub

Private Sub optRecv_Click()
    strDaNhan = " IS NULL"
End Sub

Private Sub OptToLoi_Click()
     strDaNhan = " = 'E'"
End Sub

' Ham tra ve ten cua doi tuong nop thue
Private Function getTENDTNT(ByVal maDTNT As String) As String
    Dim rs As ADODB.Recordset
    Dim strSQL As String
     'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    ' Tach ma so thue 13 thanh 14
    If Len(Trim(maDTNT)) = 13 Then
        maDTNT = Left(Trim(maDTNT), 10) & "-" & Right(Trim(maDTNT), 3)
    End If
    'Lay cau lenh truy van
    strSQL = "select ten_dtnt from rcv_v_dtnt where tin='" & Trim(maDTNT) & "'"
    Set rs = clsDAO.Execute(strSQL)
    If Not rs Is Nothing Then
        If Not IsNull(Trim(rs.Fields("TEN_DTNT"))) Then
            getTENDTNT = rs.Fields("TEN_DTNT")
        Else
            getTENDTNT = ""
        End If
    End If
End Function

