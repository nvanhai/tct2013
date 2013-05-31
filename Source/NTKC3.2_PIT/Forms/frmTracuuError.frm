VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTraCuuError 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "DS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Chän tÊt c¶"
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton btnNhanTk 
      Caption         =   "NhËn l¹i TK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnThoat 
      Caption         =   "&Tho¸t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton btnTraCuu 
      Caption         =   "Tra &cøu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Top             =   1800
      Width           =   10335
      _Version        =   458752
      _ExtentX        =   18230
      _ExtentY        =   7144
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
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
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   14
      SpreadDesigner  =   "frmTracuuError.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän ngµy nép"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         SpreadDesigner  =   "frmTracuuError.frx":06A9
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   1590
      Width           =   10575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chän lo¹i tê khai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   5415
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   435
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   4005
         _Version        =   458752
         _ExtentX        =   7064
         _ExtentY        =   767
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
         SpreadDesigner  =   "frmTracuuError.frx":0BA6
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
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   6120
      Width           =   2655
      BackColor       =   -2147483648
      VariousPropertyBits=   8388627
      Size            =   "4683;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTraCuuError"
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
'dhdang dk nut nhan lai tk
Private Sub btnNhanTk_Click()
    Dim tes As Boolean
    Dim strSQL As String
    Dim TuNgay As Date
    Dim DenNgay As Date
    Dim IdTK As String
    Dim i As Variant
    Dim check As Variant
    
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    If Not KiemTraDKngay Then
         Exit Sub
    End If
    TuNgay = format(DteTuN, "dd/mm/yyyy")
    DenNgay = format(DteDeN, "dd/mm/yyyy")
    If fpsKetQua.MaxRows > 1 Then
        For i = 1 To fpsKetQua.MaxRows - 1
            fpsKetQua.Col = 7
            fpsKetQua.Row = i
            check = fpsKetQua.Text
            If check = "1" Then
                fpsKetQua.Col = 8
                fpsKetQua.Row = i
                IdTK = fpsKetQua.Text
                strSQL = "update rcv_tkhai_hdr t set t.da_nhan = null where t.id = '" & IdTK & "'"
                'strSQL = "update rcv_tkhai_hdr t set t.da_nhan = null where t.da_nhan ='E' and t.ngay_nop >= to_date('" & TuNgay & "','dd/mm/yyyy')  and t.ngay_nop <= to_date('" & DenNgay & "','dd/mm/yyyy') and t.id = '" & IdTK & "'"
                tes = clsDAO.ExecuteQuery(strSQL)
'                If tes = False Then
'                    Exit Sub
'                End If
                If tes = True And i = fpsKetQua.MaxRows Then
                    tes = True
                End If
            End If
        Next
    End If
'load lai ket qua tren Grid
        LoadTK
    If tes = True Then

        DisplayMessage "0095", msOKOnly

    Else
        DisplayMessage "0096", msOKOnly
    End If
    Check1.Visible = btnNhanTk.Visible
    clsDAO.Disconnect
End Sub

Private Sub btnThoat_Click(Index As Integer)
    Unload Me
End Sub
Private Sub btnTraCuu_Click(Index As Integer)
    
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
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
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraCuuError")
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
                    fpsKetQua.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
                    lCol = lCol + 1
                    fpsKetQua.RowHeight(lRow) = fpsKetQua.MaxTextRowHeight(lRow)
                End If
            Next lIndex
            'dhdang edit for check colum
            'begin
            fpsKetQua.SetText 7, lRow, 0
            'end
            lCol = 2
            lRow = lRow + 1
            
            rsReturn.MoveNext
            lSoBG = lSoBG + 1
        Loop
        '    dhdang dieu khien an hien nut nhan lai to khai
       
            btnNhanTk.Visible = True
        Else
        lSoBG = 0
        DisplayMessage "0072", msOKOnly
        'dhdang begin edit
        btnNhanTk.Visible = False
        'dhdang end edit
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
    Check1.Visible = btnNhanTk.Visible
End Sub

Private Function changeLoaiToKhai(ByVal strLoaiMaToKhai As String) As String
    If strLoaiMaToKhai = "102" Then changeLoaiToKhai = "'%GTGT%'"
    If strLoaiMaToKhai = "103" Then changeLoaiToKhai = "'%TNDN%'"
    If strLoaiMaToKhai = "104" Then changeLoaiToKhai = "'%TNCN%'"
    If strLoaiMaToKhai = "105" Then changeLoaiToKhai = "'%TAIN%'"
    If strLoaiMaToKhai = "106" Then changeLoaiToKhai = "'%TTDB%'"
    If strLoaiMaToKhai = "101" Then changeLoaiToKhai = "'%NTNN%'"
    
    If strLoaiMaToKhai = "113" Then changeLoaiToKhai = "'%BVMT%' or hdr.loai_tkhai like '%PHXD%'"
    
    If strLoaiMaToKhai = "108" Then changeLoaiToKhai = "'%15%'"
    If strLoaiMaToKhai = "109" Then changeLoaiToKhai = "'%48%'"
    If strLoaiMaToKhai = "110" Then changeLoaiToKhai = "'%16%'"
    If strLoaiMaToKhai = "111" Then changeLoaiToKhai = "'%99%'"
    
    
    If strLoaiMaToKhai = "0" Then changeLoaiToKhai = "'%'"
End Function


Private Sub Check1_Click()
    Dim lRow As Long
    Dim varTemp, varTemp1 As Variant
    If Check1.Value = 1 Then
        Check1.caption = "Bá chän tÊt c¶"
    Else
        Check1.caption = "Chän tÊt c¶"
    End If
    With fpsKetQua
        For lRow = 1 To .MaxRows
            .Row = lRow
            .Col = .ColLetterToNumber("B")
            .GetText .Col, .Row, varTemp
            .Col = .ColLetterToNumber("D")
            .GetText .Col, .Row, varTemp1
            
            .Col = .ColLetterToNumber("G")
            
            If Check1.Value = 1 Then
                If (Trim(varTemp) <> vbNullString) And (Trim(varTemp1) <> vbNullString) Then
                    .Text = "1"
                Else
                    .Text = "0"
                End If
            Else
                .Text = "0"
            End If
        Next
    End With
End Sub

Private Sub Form_Load()
    SetControlCaption Me, "frmTraCuuError"
    FormatGrid
    SetupData
    'lerror = False
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
    Check1.Visible = btnNhanTk.Visible
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
                If Parentid = "101" And Val(GetAttribute(xmlNode, "ID")) <> 112 Then
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
        strDaNhan = "='E'"
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    frmTraCuuError.Top = (frmSystem.Height - frmTraCuuError.Height) / 3
    frmTraCuuError.Left = (frmSystem.Width - frmTraCuuError.Width) / 2
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

Private Sub fpsKetQua_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
     'MsgBox "tess"
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
Private Function LoadTK()
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    'Xoa bo ket qua cu tren Grid
    fpsKetQua.ClearRange 1, 1, fpsKetQua.MaxCols, lSoBG, True
    If Not KiemTraDKngay Then
         Exit Function
    End If
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    'Lay cau lenh truy van
    xmlSQL.Load App.path & "\SQL.xml"
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraCuuError")
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
                    fpsKetQua.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
                    lCol = lCol + 1
                    fpsKetQua.RowHeight(lRow) = fpsKetQua.MaxTextRowHeight(lRow)
                End If
            Next lIndex
            'dhdang edit for check colum
            'begin
            fpsKetQua.SetText 7, lRow, 0
            'end
            lCol = 2
            lRow = lRow + 1
            
            rsReturn.MoveNext
            lSoBG = lSoBG + 1
        Loop
        '    dhdang dieu khien an hien nut nhan lai to khai
       
            btnNhanTk.Visible = True
    Else
        lSoBG = 0
        'DisplayMessage "0072", msOKOnly
        'dhdang begin edit
        btnNhanTk.Visible = False
        'dhdang end edit
    End If
    LblSoBG.Visible = True
    LblSoBG.TextAlign = fmTextAlignLeft
    LblSoBG.caption = "Sè tê khai t×m thÊy: " & lSoBG
    
End Function

Private Sub optQLT_Click()
    strDaNhan = "='E'"
    btnNhanTk.Visible = False
End Sub

Private Sub optRecv_Click()
    'strDaNhan = " IS NULL"
    strDaNhan = "='E'"
    btnNhanTk.Visible = False
End Sub

Private Sub OptToLoi_Click()
     strDaNhan = "='E'"
End Sub
