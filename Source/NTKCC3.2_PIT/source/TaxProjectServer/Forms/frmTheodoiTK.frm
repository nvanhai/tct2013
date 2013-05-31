VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTheodoiTK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "M· sè thuÕ"
      Height          =   615
      Left            =   90
      TabIndex        =   13
      Top             =   1290
      Width           =   3465
      Begin VB.TextBox txtMST 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   930
         MaxLength       =   15
         TabIndex        =   14
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tr¹ng th¸i tê khai"
      Height          =   615
      Left            =   3630
      TabIndex        =   10
      Top             =   1290
      Width           =   7335
      Begin VB.OptionButton optIHTKK 
         Caption         =   "Tê khai ch­a göi lªn tæng Côc"
         Height          =   255
         Left            =   540
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optRecv 
         Caption         =   "Tê khai ®· göi lªn tæng Côc"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   270
         Width           =   2865
      End
   End
   Begin VB.CommandButton btnThoat 
      Caption         =   "&Tho¸t"
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   4
      Top             =   6150
      Width           =   1215
   End
   Begin VB.CommandButton btnTraCuu 
      Caption         =   "Tra cøu"
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   3
      Top             =   6150
      Width           =   1215
   End
   Begin FPUSpreadADO.fpSpread fpsKetQua 
      Height          =   3675
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   10635
      _Version        =   458752
      _ExtentX        =   18759
      _ExtentY        =   6482
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
      MaxCols         =   9
      MaxRows         =   13
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmTheodoiTK.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän ngµy nép"
      Height          =   885
      Left            =   5580
      TabIndex        =   5
      Top             =   360
      Width           =   5385
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
         SpreadDesigner  =   "frmTheodoiTK.frx":07EB
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      Height          =   3975
      Left            =   90
      TabIndex        =   6
      Top             =   1890
      Width           =   10875
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chän lo¹i tê khai"
      Height          =   885
      Left            =   90
      TabIndex        =   7
      Top             =   360
      Width           =   5445
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   525
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5205
         _Version        =   458752
         _ExtentX        =   9181
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
         SpreadDesigner  =   "frmTheodoiTK.frx":0D12
         UserResize      =   1
      End
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   90
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
      Top             =   6000
      Width           =   3015
      BackColor       =   -2147483648
      VariousPropertyBits=   8388627
      Size            =   "5318;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTheodoiTK"
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
Private Const mHeaderColor = 16709097

Private lngRowFocus As Long
Private DteTuN As Date
Private DteDeN As Date
Private lmaTK As Long
Private lSoBG As Long
Private larrId(lmaxSoTk) As Long
Private lerror As Boolean
Private strDaNhan As String
Private strTT_Gui As String
Private strTracuu As String

Private Sub btnThoat_Click(Index As Integer)
    Unload Me
End Sub
Public Sub traCuuToKhai()
'    Dim rsReturn As New ADODB.Recordset
'    Dim strSQL As String
'    Dim strMa_Cqthue As String
'    Dim lIndex As Long, lCtrl As Long
'    Dim lCol As Long, lRow As Long
'    'Xoa bo ket qua cu tren Grid
'    'fpsKetQua.ClearRange 1, 1, fpsKetQua.MaxCols, lSoBG, True
'    If Not KiemTraDKngay Then
'         Exit Sub
'    End If
'    'connect to database QLT
'    If Not clsDAO.ConnectedTG Then
'        clsDAO.CreateConnectionStringTg [MSDAORA.1], "QLT", "QLT_NTK", "QLT_NTK"
'        clsDAO.ConnectTG
'    End If
'    'Lay cau lenh truy van
'    xmlSQL.Load App.path & "\SQL.xml"
'    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraCuuTC")
'    strSQL = Replace(strSQL, "nhohon=", "<=")
'    strSQL = Replace(strSQL, "ma_tkhai", "" & changeLoaiToKhai(lmaTK) & "")
'    strSQL = Replace(strSQL, "strTT_Gui", "" & strTT_Gui & "")
'    strSQL = Replace(strSQL, "strMa_Cqthue", "" & strMa_Cqthue & "")
'    strSQL = Replace(strSQL, "ngay_nop_dau", "To_date('" & format(DteTuN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
'    strSQL = Replace(strSQL, "ngay_nop_cuoi", "To_date('" & format(DteDeN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
'    'Khoi tao bien
'    lCol = 2
'    lRow = 1
'    lSoBG = 0
'    'Thuc hien cau lenh sql
'    If clsDAO.ConnectedTG Then
'        Set rsReturn = clsDAO.ExecuteTg(strSQL)
'        'Hien thi du lieu len grid
'        If rsReturn.Fields.Count > 0 Then
'            Do While Not rsReturn.EOF
'                fpsKetQua.MaxRows = lRow + 1
'                fpsKetQua.InsertRows lRow, 1
'                fpsKetQua.SetText 1, lRow, lSoBG + 1
'                For lIndex = 1 To rsReturn.Fields.Count
'                    If Not (IsNull(rsReturn.Fields(lIndex - 1).Value)) Then
'                        fpsKetQua.SetText lCol, lRow, TAX_Utilities_Svr_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
'                        lCol = lCol + 1
'                        fpsKetQua.RowHeight(lRow) = fpsKetQua.MaxTextRowHeight(lRow)
'                    End If
'                Next lIndex
'                lCol = 2
'                lRow = lRow + 1
'
'                rsReturn.MoveNext
'                lSoBG = lSoBG + 1
'            Loop
'
'        Else
'            lSoBG = 0
'            DisplayMessage "0072", msOKOnly
'        End If
'        LblSoBG.Visible = True
'        LblSoBG.TextAlign = fmTextAlignLeft
'        LblSoBG.caption = "Sè tê khai t×m thÊy: " & lSoBG
'        'Chinh sua lai Grid ket cho dep
'        With fpsKetQua
'            If lSoBG <= 15 Then
'                Dim i As Integer
'                For i = 1 To 15 - lRow
'                    .MaxRows = lSoBG + i + 1
'                    .InsertRows lSoBG + i, 1
'                Next
'            Else
'                .RowHeight(lRow) = 0
'            End If
'
'            If lngRowFocus = 0 Then
'                lngRowFocus = SetRowFocus(1, 1, True)
'            ElseIf lngRowFocus > .MaxRows Then
'                lngRowFocus = SetRowFocus(1, .MaxRows, True)
'            Else
'                lngRowFocus = SetRowFocus(1, lngRowFocus, True)
'            End If
'            .SetFocus
'        End With
'    Else
'        DisplayMessage "0131", msOKOnly, miCriticalError
'        Exit Sub
'    End If
End Sub


Private Function changeLoaiToKhai(ByVal strLoaiMaToKhai As String) As String
    If strLoaiMaToKhai = "15" Then changeLoaiToKhai = " = '02A_TNCN10'"
    If strLoaiMaToKhai = "16" Then changeLoaiToKhai = " = '02B_TNCN10'"
    If strLoaiMaToKhai = "50" Then changeLoaiToKhai = " ='03A_TNCN10'"
    If strLoaiMaToKhai = "51" Then changeLoaiToKhai = "  ='03B_TNCN10'"
    If strLoaiMaToKhai = "36" Then changeLoaiToKhai = " ='07_TNCN10'"
    
    If strLoaiMaToKhai = "0" Then changeLoaiToKhai = " LIKE '%' "
    If Trim(txtMST.Text) <> vbNullString Then
        changeLoaiToKhai = changeLoaiToKhai & ") And (TIN like '%" & UCase(Trim(txtMST.Text)) & "%'"
    End If
End Function

Private Function validToKhaiiHTKK(ByVal maSoThue As String, ByVal maToKhai As String, ByVal kyKeKhai As String, ByVal lanNop As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    validToKhaiiHTKK = False
    
    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'    ' Lay tat ca cac to khai trong cung mot ky tinh thue, cua mot loai to khai, cua mot nguoi nop thue va co so lan nop den thoi diem kiem tra
'    strSQL = "select ID from RCV_IHTKK_MVACH where TIN = '" & maSoThue & "' and MA_TKHAI = '" & maToKhai & "' and KY_KKHAI = '" & kyKeKhai & "' and DA_NHAN IS NULL and LAN_NOP < " & lanNop
'    Set rs = clsDAO.Execute(strSQL)
'
'    'Neu con to khac lan nop truoc do chua duoc chuyen vao QLT_NTK thi tra lai ket qua la true
'    If rs.Fields.Count > 0 Then
'        validToKhaiiHTKK = True
'    End If
End Function





Private Sub btnTraCuu_Click(Index As Integer)
    traCuuToKhai
End Sub

Private Sub Form_Load()
    SetControlCaption Me, "frmTheodoiTK"
    FormatGrid
    SetupData
    'lerror = False
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
End Sub
Sub FormatGrid()
    Dim i As Integer
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
            
            xmlDocument.Load TAX_Utilities_Svr_New.GetAbsolutePath("menu.xml")
            Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            i = 0
            larrId(0) = 0
            For Each xmlNode In xmlNodeListMenu
                Dim LoaiTk As String
                Dim Parentid As String
                Dim strID As Variant
                Parentid = GetAttribute(xmlNode, "ParentID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                strID = Val(GetAttribute(xmlNode, "ID"))
                '.TypeComboBoxIndex = 0
                If Parentid = "101_11" Then
                    If strID = 15 Or strID = 16 Or strID = 50 Or strID = 51 Or strID = 36 Then
                        i = i + 1
                        .TypeComboBoxIndex = -1
                        .TypeComboBoxString = LoaiTk
                        larrId(i) = Val(GetAttribute(xmlNode, "ID"))
                    End If
                End If
            Next
            .TypeComboBoxCurSel = 0
            Set xmlNode = Nothing
            Set xmlDocument = Nothing
        End With
        strDaNhan = " = 'N' "
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    frmTheodoiTK.Top = (frmSystem.Height - frmTheodoiTK.Height) / 2
    frmTheodoiTK.Left = (frmSystem.Width - frmTheodoiTK.Width) / 2
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
    Dim vxoatat As Variant, varValue As Variant
    Dim i As Integer, blnCheck As Boolean
    
    With fpsKetQua
        GetCellSpan fpsKetQua, Col, Row
        If Col = 2 Then
            .EventEnabled(EventAllEvents) = False
            If Row = 1 Then
                .GetText Col, Row, vxoatat
                If vxoatat = "1" Then
                    For i = 2 To lSoBG + 1
                        .SetText Col, i, "1"
                        blnCheck = True
                    Next
                Else
                    For i = 2 To lSoBG + 1
                        .SetText Col, i, "0"
                    Next
                    blnCheck = False
                End If
            Else
                blnCheck = False
                .SetText Col, 1, "1"
                For i = 2 To lSoBG + 1
                    .GetText Col, i, varValue
                    If CStr(varValue) = "0" Or CStr(varValue) = "" Then
                        .SetText Col, 1, "0"
                        'Exit For
                    Else
                        blnCheck = True
                    End If
                Next i
            End If
            .EventEnabled(EventAllEvents) = True
        End If
        
        If Row > 1 Then
            .ReDraw = False
            lngRowFocus = SetRowFocus(lngRowFocus, Row, True)
            .ReDraw = True
        End If
        
        
    End With
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

Private Sub optIHTKK_Click()
    strDaNhan = " ='N'"
End Sub

Private Sub optRecv_Click()
    strDaNhan = " = 'Y'"
End Sub

' Ham tra ve ten cua doi tuong nop thue
Private Function getTENDTNT(ByVal madtnt As String) As String
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    If Not clsConn.Connected Then
        clsConn.CreateConnectionString spathVat & "\DTNT\"
        clsConn.Connect
    End If
    'Lay cau lenh truy van
    strSQL = "select TENGOI from DTNT2 where MADTNT ='" & Trim(madtnt) & "'"
    Set rs = clsConn.Execute(strSQL)
    If Not rs Is Nothing Then
        If Not IsNull(Trim(rs.Fields("TENGOI"))) Then
            getTENDTNT = rs.Fields("TENGOI")
        Else
            getTENDTNT = ""
        End If
    End If
    clsConn.Disconnect
End Function
