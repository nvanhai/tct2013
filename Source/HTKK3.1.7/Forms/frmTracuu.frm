VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmTraCuu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnTracuu 
      Caption         =   "Tra &cøu"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   3
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton btnXoa 
      Caption         =   "&Xo¸"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6510
      TabIndex        =   5
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton btnThoat 
      Caption         =   "&Tho¸t"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7770
      TabIndex        =   6
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton btnMo 
      Caption         =   "&Më"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      TabIndex        =   4
      Top             =   4950
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän lo¹i tê khai"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   420
      Width           =   4455
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   495
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   4215
         _Version        =   458752
         _ExtentX        =   7435
         _ExtentY        =   873
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
         SpreadDesigner  =   "frmTracuu.frx":0000
         UserResize      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän kú kª khai"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   4530
      TabIndex        =   8
      Top             =   420
      Width           =   4425
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   495
         Left            =   300
         TabIndex        =   1
         Top             =   240
         Width           =   4005
         _Version        =   458752
         _ExtentX        =   7064
         _ExtentY        =   873
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
         MaxCols         =   12
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmTracuu.frx":044F
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   30
      TabIndex        =   9
      Top             =   1320
      Width           =   8925
      Begin FPUSpreadADO.fpSpread fpSKetQua 
         Height          =   3225
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   8745
         _Version        =   458752
         _ExtentX        =   15425
         _ExtentY        =   5689
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
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
         MaxCols         =   19
         MaxRows         =   18
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowsFrozen      =   1
         SpreadDesigner  =   "frmTracuu.frx":0A50
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Tê khai cã mµu ®á lµ tê khai kh«ng hîp lÖ."
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Tra cøu th«ng tin tê khai"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image imgCaption 
      Height          =   345
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmTraCuu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmAddSheet
' Descriptions      : Report sh
' Start date        : 10/08/2007 (dd/mm/yyyy)
' Finish date       :
' Coder             : htphuong
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************

Option Explicit
Private Const f1dteTuNRow = 2
Private Const f1dteTuNCol = "F"
Private Const f1dteDeNRow = 2
Private Const f1dteDeNCol = "K"
Private Const f4cboLoTRow = 2
Private Const f4cboLoTCol = "C"
Private Const mFormColor = -2147483633
Private Const mHeaderColor = 16709097
Private Const LOI_KY_HIEU_LUC = "1"
Private Const LOI_NGAY_BAT_DAU_NAM_TAI_CHINH = "2"
Private Const LOI_TU_NGAY_DEN_NGAY = "3"
Private lstryear As String
Private lstrMonth As String
Private lstrThreemonths As String
Private arrStrId() As String
Private Dtetun As String
Private Dteden As String
Private lSoBG As Long
Private arrCheckStatus() As Long
Private blnOpenInterfaces As Boolean
Private blnOnExit As Boolean                    'Bat su kien Click vao nut Thoat
Private arrStrXMLFileNames() As String
Private arrLngErrRows() As Long
Private lngRowFocus As Long
Private blnDKienTraCuu As Boolean               'Kiem tra dieu kien tra cuu co hop le hay ko

Private Sub btnMo_Click()
    Dim frmTK As frmInterfaces
    Dim strxoa As Variant, varErrDesc As Variant
    Dim varId As Variant, varPeriod As Variant
    Dim varFirstDay As Variant, varLastDay As Variant, vCheckStatus As Variant
    Dim varFileName As Variant
    Dim i, j As Integer
    Dim varDateKHBS As Variant, strFileName As Variant
    With fpSKetQua
        If IsErrRow(lngRowFocus) Then
            .GetText 12, lngRowFocus, varErrDesc
            Select Case Trim(CStr(varErrDesc))
                Case LOI_KY_HIEU_LUC
                    DisplayMessage "0096", msOKOnly, miCriticalError
                Case LOI_TU_NGAY_DEN_NGAY
                    DisplayMessage "0101", msOKOnly, miCriticalError
                Case LOI_NGAY_BAT_DAU_NAM_TAI_CHINH
                    DisplayMessage "0100", msOKOnly, miCriticalError
            End Select
            
            Exit Sub
        End If
        
        .GetText 3, lngRowFocus, varId        ' Get ID
        .GetText 5, lngRowFocus, varDateKHBS        ' Get DateKHBS
        .GetText 6, lngRowFocus, varPeriod    ' Get period
        .GetText 7, lngRowFocus, varFirstDay    ' Get first day
        .GetText 8, lngRowFocus, varLastDay    ' Get last day
        .GetText 11, lngRowFocus, varFileName    ' Get File name
        
        
        If Left(varId, 4) = "KHBS" Then
            varId = Right(varId, 2)
            strKHBS = "frmKHBS_BS"
            TAX_Utilities_New.DateKHBS = varDateKHBS
        End If
        If Left(varFileName, 2) = "bs" Then
            strKHBS = "TKBS"
            strSolanBS = Right(Split(varFileName, "_")(0), Len(Split(varFileName, "_")(0)) - 2)
            'TAX_Utilities_New.DateKHBS = varDateKHBS
        Else
            ' Neu la loai to khai TNCN thi dat trang thai cua strKHBS ="TKCT"
            If Trim(varId) = "46" Or Trim(varId) = "47" Or Trim(varId) = "48" Or Trim(varId) = "49" Or Trim(varId) = "15" Or Trim(varId) = "16" _
                Or Trim(varId) = "53" Or Trim(varId) = "37" Or Trim(varId) = "50" Or Trim(varId) = "51" Or Trim(varId) = "54" Or Trim(varId) = "38" _
                    Or Trim(varId) = "39" Or Trim(varId) = "40" Or Trim(varId) = "36" Or Trim(varId) = "70" Or Trim(varId) = "17" Or Trim(varId) = "41" Or Trim(varId) = "42" Or Trim(varId) = "43" Then
                strKHBS = "TKCT"
            End If
            
        End If
        
        TAX_Utilities_New.NodeMenu = getNode(CStr(varId))
        ' 12110211 xu ly to khai BS
        If strKHBS = "TKBS" Then
            For i = 1 To TAX_Utilities_New.NodeMenu.childNodes(0).childNodes.length - 1
                If i = TAX_Utilities_New.NodeMenu.childNodes(0).childNodes.length - 1 Then
                    SetAttribute TAX_Utilities_New.NodeMenu.childNodes(0).childNodes(i), "Active", "1"
                Else
                    SetAttribute TAX_Utilities_New.NodeMenu.childNodes(0).childNodes(i), "Active", "0"
                End If
            Next
            
        End If
        
        
        If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
            TAX_Utilities_New.month = Mid$(CStr(varPeriod), 1, 2)
            TAX_Utilities_New.Year = Mid$(CStr(varPeriod), 4, 4)
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
            TAX_Utilities_New.ThreeMonths = CInt(Mid$(CStr(varPeriod), 1, 2))
            TAX_Utilities_New.Year = Mid$(CStr(varPeriod), 4, 4)
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" Then
            If varId = "80" Or varId = "82" Then
                TAX_Utilities_New.Year = Right(CStr(varPeriod), 4)
            Else
                TAX_Utilities_New.Year = CStr(varPeriod)
            End If
            TAX_Utilities_New.FirstDay = CStr(varFirstDay)
            TAX_Utilities_New.LastDay = CStr(varLastDay)
        Else
            TAX_Utilities_New.Year = CStr(varPeriod)
        End If
        
        j = 1
        ReDim Preserve arrCheckStatus(0)
        arrCheckStatus(0) = 0
        For i = 2 To lSoBG + 1
            .GetText 2, i, vCheckStatus
            If vCheckStatus = 1 Or vCheckStatus = "1" Then
                ReDim Preserve arrCheckStatus(j)
                arrCheckStatus(j) = i
                j = j + 1
            End If
        Next

        DoEvents
        strHiddenFormName = Me.Name
        Me.Hide
        
        Set frmTK = New frmInterfaces
        frmTK.Show
        
    End With
   
End Sub

Private Sub btnThoat_Click()
    Unload Me
    strHiddenFormName = vbNullString
End Sub

Private Sub btnThoat_LostFocus()
    blnOnExit = False
End Sub

Private Sub btnThoat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnOnExit = True
End Sub

'********************************************
'Description: Hien thi thong tin tra cuu len bang ket qua theo dieu
'             kien tra cuu.
'Author:
'Date: 25/04/06
'Return:
'********************************************
Private Sub btnTracuu_Click()
    TraCuu
End Sub

Private Sub btnXoa_Click()
With fpSKetQua
    Dim i As Integer
    Dim strxoa As Variant
    Dim strDataFile As Variant
    Dim fso As New FileSystemObject
    If DisplayMessage("0094", msYesNo, miQuestion, , mrNo) = mrYes Then
        For i = lSoBG To 1 Step -1
            .GetText 2, i + 1, strxoa
            If strxoa = "1" Then
                .GetText 11, i + 1, strDataFile
                DeleteDataFiles (strDataFile)
'                If fso.FileExists(GetAbsolutePath(TAX_Utilities_New.DataFolder & strDataFile & ".xml")) Then
'                    fso.DeleteFile GetAbsolutePath(TAX_Utilities_New.DataFolder & strDataFile & ".xml"), True
'                    .DeleteRows i + 1, 1
'                End If
            End If
        Next
        TraCuu
    End If
End With
End Sub
Private Sub Form_Activate()
    Dim iSoBG, i, j As Integer
    Dim irowfocus As Long
    Dim vstatus, vcheckall As Variant
    Dim finish As Boolean
    
    DoEvents
    If strInterfaceUnloadEventName <> "" Then
        iSoBG = lSoBG
        irowfocus = lngRowFocus
        'Lay vi tri focus
        fpSKetQua.GetText 2, irowfocus, vstatus
        fpSKetQua.GetText 2, 1, vcheckall
        TraCuu
        DoEvents
        btnMo.Default = True
        
        If UBound(arrCheckStatus) > 0 Then
            btnXoa.Enabled = True
        End If
        
        With fpSKetQua
            .EventEnabled(EventAllEvents) = False
            If iSoBG = lSoBG Then
                'Ko xoa
                For i = 1 To UBound(arrCheckStatus)
                    .SetText 2, arrCheckStatus(i), "1"
                Next
            Else
                'Xoa
                If UBound(arrCheckStatus) > 0 Then
                    i = 1
                    finish = False
                    While Not finish And arrCheckStatus(i) < irowfocus
                        .SetText 2, arrCheckStatus(i), "1"
                        i = i + 1
                        If i > UBound(arrCheckStatus) Then
                            i = i - 1
                            finish = True
                        End If
                    Wend
                    j = i + 1
                    
                    If vstatus = "1" Then
                        For j = i + 1 To UBound(arrCheckStatus)
                            .SetText 2, arrCheckStatus(j) - 1, "1"
                        Next
                    Else
                        For j = i To UBound(arrCheckStatus)
                            If arrCheckStatus(j) >= irowfocus Then
                                .SetText 2, arrCheckStatus(j) - 1, "1"
                            End If
                        Next
                    End If
                
                End If
            End If
            If lSoBG > 0 Then
                vcheckall = "1"
            Else
                vcheckall = "0"
            End If
            For i = 2 To lSoBG + 1
                .GetText 2, i, vcheckall
                If vcheckall = "0" Or vcheckall = "" Then
                    Exit For
                End If
            Next
            .SetText 2, 1, vcheckall
            .EventEnabled(EventAllEvents) = True
        End With
    End If
    If lngRowFocus = 0 Then
        lngRowFocus = SetRowFocus(2, 2, True)
    ElseIf lngRowFocus > fpSKetQua.MaxRows Then
        lngRowFocus = SetRowFocus(2, fpSKetQua.MaxRows, True)
    Else
        lngRowFocus = SetRowFocus(2, lngRowFocus, True)
    End If
End Sub

Private Sub Form_Load()
    'SetControlCaption Me, "frmTraCuu"
    FormatGrid
    SetupData
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
    
    blnOnExit = False
    
    Me.Top = (frmSystem.ScaleHeight - Me.Height) / 2 - 250
    Me.Left = (frmSystem.ScaleWidth - Me.Width) / 2
    
    btnMo.Enabled = False
    btnXoa.Enabled = False
    
    strInterfaceUnloadEventName = ""

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
        .CellType = CellTypePic
        .TypePicMask = "9999"
        
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        .BackColor = vbWhite
        .CellType = CellTypePic
        .TypePicMask = "9999"
        
            fpsDkNgay.Col = 2
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 3
            fpsDkNgay.ColHidden = False
            fpsDkNgay.Col = 4
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 5
            fpsDkNgay.ColHidden = True
            
            fpsDkNgay.Col = 7
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 8
            fpsDkNgay.ColHidden = False
            fpsDkNgay.Col = 9
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 10
            fpsDkNgay.ColHidden = True
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
    With fpSKetQua
        .MaxCols = 12
        .EditModePermanent = True
        .EditModeReplace = True
        .CursorType = CursorTypeLockedCell
        .CursorStyle = CursorStyleArrow
        .TypeNumberNegStyle = TypeNumberNegStyle1
        .ColWidth(11) = 0
        .Row = 1
        .RowHeight(1) = 25

        For i = 1 To .MaxCols
            .Col = i
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignCenter
            .BackColor = mHeaderColor
        Next

    End With
End Sub
Sub SetupData()
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    
        With fpsDkNgay
            Dim vdtehientai As String
            vdtehientai = format(Date, "dd/mm/yyyy")
            Dim strarrdate() As String
            formatPrefix vdtehientai, strarrdate
            
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .Text = strarrdate(2)
            '.Text = strarrdate(0) & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .Text = strarrdate(2)
            '.Text = strarrdate(0) & "/" & strarrdate(2)
        End With
    'Lay du lieu cho cbo
        With fpsLoaiTK
            .Col = .ColLetterToNumber(f4cboLoTCol)
            .Row = f4cboLoTRow
            Dim xmlDocument As New MSXML.DOMDocument
            Dim xmlNode As MSXML.IXMLDOMNode
            Dim strDataFileName As String
            Dim i As Integer

            xmlDocument.Load TAX_Utilities_New.GetAbsolutePath("Map.xml")
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            ReDim Preserve arrStrId(0)
            arrStrId(0) = "00"
            For Each xmlNode In xmlNodeListMap
                Dim id As String
                Dim LoaiTk As String
                Dim Parentid As String
                'Parentid = GetAttribute(xmlNode, "ParentID")
                   '.TypeComboBoxIndex = 0
                'If Parentid = "101" Then
                i = i + 1
                ReDim Preserve arrStrId(i)
                arrStrId(i) = GetAttribute(xmlNode, "ID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = LoaiTk
                'End If
            Next
            'lintSoTk = i
            .TypeComboBoxCurSel = 0
            lstryear = "1"
            lstrMonth = "0"
            lstrThreemonths = "0"
            Set xmlNode = Nothing
            Set xmlDocument = Nothing
        End With
End Sub
Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not blnOpenInterfaces Then
        frmTreeviewMenu.Show
    End If
    Set frmTraCuu = Nothing
    lngRowFocus = 0
End Sub

Private Sub fpsDkNgay_GotFocus()
    btnTracuu.Default = True
End Sub

'Private Sub fpsDkNgay_Click(ByVal Col As Long, ByVal Row As Long)
'With fpsDkNgay
'    MsgBox "Col: " & .Col & "            Row:" & .Row
'End With
'End Sub
Private Sub fpsDkNgay_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab And Shift = 0 Then
    With fpsDkNgay
        If fpsDkNgay.ActiveRow = f1dteDeNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteDeNCol) Then
            fpSKetQua.SetFocus
        End If
    End With
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
Private Sub fpsDkNgay_KeyPress(KeyAscii As Integer)
    'Chan ko cho nhap dau "."
    If KeyAscii = Asc(".") Then
        KeyAscii = 0
    End If
End Sub
Private Sub fpsDkNgay_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With fpsDkNgay
        Dim strarrdate() As String
        Dim strPrefix As String
        Dim vdtehientai As String
        Static blnOnDKienTraCuu_LeaveCell As Boolean          'Kiem tra su kien LeaveCell dang dc goi ???
        
        'Khi nguoi dung bam nut Thoat -> Bo qua ham nay.
        If blnOnExit Then Exit Sub
        
        '*******************************
        ' added
        'Date: 13/05/06
        'Su kien nay dang dc goi
        If blnOnDKienTraCuu_LeaveCell = True Then Exit Sub
        
        'Khoi tao gia tri dieu kien tra cuu
        blnOnDKienTraCuu_LeaveCell = True
        '*******************************
        
        'Khoi tao gia tri dieu kien tra cuu
        blnDKienTraCuu = False
        
        vdtehientai = format(Date, "dd/mm/yyyy")
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        
        If Trim(Replace(.Text, "/", "")) <> "" Then
            formatPrefix .Text, strarrdate
            'Bat dk thang
            If (Val(strarrdate(0)) > 12 Or Val(strarrdate(0)) <= 0) And lstrMonth = "1" Then
                .Text = ""
                DisplayMessage "0090", msOKOnly, miInformation
                blnOnDKienTraCuu_LeaveCell = False
                Exit Sub
            End If
            'Bat dk quy
            If (Val(strarrdate(0)) > 4 Or Val(strarrdate(0)) <= 0) And lstrThreemonths = "1" Then
                .Text = ""
                DisplayMessage "0091", msOKOnly, miInformation
                blnOnDKienTraCuu_LeaveCell = False
                Exit Sub
            End If
                'bat dk nam
            If lstrMonth = "0" And lstrThreemonths = "0" Then
                If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Then
                Else
                    Select Case Len(Trim(strarrdate(0)))
                        Case 1
                            strarrdate(0) = Year(vdtehientai)
                        Case 2
                            strarrdate(0) = "20" & Trim(strarrdate(0))
                        Case 3
                            strarrdate(0) = "2" & Trim(strarrdate(0))
                        Case 4
                            If Val(strarrdate(0)) < 2000 Then
                                strarrdate(0) = "2000"
                            End If
                    End Select
               End If
            Else
                Select Case Len(Trim(strarrdate(1)))
                Case 1
                    strarrdate(1) = Year(vdtehientai)
                Case 2
                    strarrdate(1) = "20" & Trim(strarrdate(1))
                Case 3
                    strarrdate(1) = "2" & Trim(strarrdate(1))
                Case 4
                    If Val(strarrdate(1)) < 2000 Then
                          strarrdate(1) = "2000"
                    End If
                End Select
            End If
            'Hien thi lai kq
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            If lstrMonth = "0" And lstrThreemonths = "0" Then
                If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Then
                    
                Else
                    .SetText .Col, .Row, strarrdate(0)
                End If
            Else
                .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1)
            End If
        End If

        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        If Trim(Replace(.Text, "/", "")) <> "" Then
            formatPrefix .Text, strarrdate
            If (Val(strarrdate(0)) > 12 Or Val(strarrdate(0)) <= 0) And lstrMonth = "1" Then
                DisplayMessage "0090", msOKOnly, miInformation
                Cancel = True
                .SetFocus
                .Text = ""
                .SetActiveCell .Col, .Row
                blnOnDKienTraCuu_LeaveCell = False
                Exit Sub
            End If
            If (Val(strarrdate(0)) > 4 Or Val(strarrdate(0)) <= 0) And lstrThreemonths = "1" Then
                DisplayMessage "0091", msOKOnly, miInformation
                Cancel = True
                .SetFocus
                .Text = ""
                .SetActiveCell .Col, .Row
                blnOnDKienTraCuu_LeaveCell = False
                Exit Sub
            End If
            'bat dk nam
            If lstrMonth = "0" And lstrThreemonths = "0" Then
                If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Then
                    
                Else
                    Select Case Len(Trim(strarrdate(0)))
                    Case 1
                        strarrdate(0) = Year(vdtehientai)
                    Case 2
                        strarrdate(0) = "20" & Trim(strarrdate(0))
                    Case 3
                        strarrdate(0) = "2" & Trim(strarrdate(0))
                    Case 4
                        If Val(strarrdate(0)) < 2000 Then
                            strarrdate(0) = "2000"
                        End If
                    End Select
                End If
            Else
                Select Case Len(Trim(strarrdate(1)))
                Case 1
                     strarrdate(1) = Year(vdtehientai)
                Case 2
                    strarrdate(1) = "20" & Trim(strarrdate(1))
                Case 3
                    strarrdate(1) = "2" & Trim(strarrdate(1))
                Case 4
                    If Val(strarrdate(1)) < 2000 Then
                          strarrdate(1) = "2000"
                    End If
                End Select
            End If
            'Hien thi lai kq
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            If lstrMonth = "0" And lstrThreemonths = "0" Then
                If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Then
                    
                Else
                    .SetText .Col, .Row, strarrdate(0)
                End If
            Else
                .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1)
            End If
        End If
    End With
    
    blnOnDKienTraCuu_LeaveCell = False
    blnDKienTraCuu = True
End Sub

Private Sub fpSKetQua_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim vxoatat As Variant, varValue As Variant
    Dim i As Integer, blnCheck As Boolean
    
    With fpSKetQua
        GetCellSpan fpSKetQua, Col, Row
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
        '********************************************
        ' added
        'Date: 04/05/06
        
        ' Set row focus to fpsKetQua
        If Row > 1 Then
            .ReDraw = False
            lngRowFocus = SetRowFocus(lngRowFocus, Row, True)
            .ReDraw = True
        End If
        
        'Set enable to btnXoa
        If blnCheck Then
            btnXoa.Enabled = True
        Else
            btnXoa.Enabled = False
        End If
        '********************************************
    End With
End Sub

Private Sub fpSKetQua_Click(ByVal Col As Long, ByVal Row As Long)

    If Col = 2 Then Exit Sub
    If Row = 1 Then Exit Sub

    With fpSKetQua
        .ReDraw = False
        lngRowFocus = SetRowFocus(lngRowFocus, Row, True)
        .ReDraw = True
    End With
End Sub

Private Sub fpSKetQua_DblClick(ByVal Col As Long, ByVal Row As Long)
    'btnMo_Click
End Sub

Private Sub fpSKetQua_GotFocus()
    btnMo.Default = True
End Sub

Private Sub fpSKetQua_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab And Shift = 1 Then
       fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteDeNCol), f1dteDeNRow
        End With
    End If
    If KeyCode = vbKeyDown And lngRowFocus < fpSKetQua.MaxRows Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus + 1)
    ElseIf KeyCode = vbKeyUp And lngRowFocus > 2 Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus - 1)
    End If
    
End Sub

Private Sub fpsLoaiTK_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    With fpsLoaiTK
        .Col = .ColLetterToNumber(f4cboLoTCol)
        .Row = f4cboLoTRow
        If .TypeComboBoxCurSel = 0 Then
            With fpsDkNgay
                Dim vdtehientai As String
                vdtehientai = format(Date, "dd/mm/yyyy")
                Dim strarrdate() As String
                formatPrefix vdtehientai, strarrdate
            
                .Col = .ColLetterToNumber(f1dteTuNCol)
                .Row = f1dteTuNRow
                .CellType = CellTypePic
                .TypePicMask = "9999"
                '.Text = strarrdate(0) & "/" & strarrdate(2)
                .Text = strarrdate(2)
                
                .Col = .ColLetterToNumber(f1dteDeNCol)
                .Row = f1dteDeNRow
                .CellType = CellTypePic
                .TypePicMask = "9999"
                .Text = strarrdate(2)
                lstryear = "1"
                lstrMonth = "0"
                lstrThreemonths = "0"
                '.Text = strarrdate(0) & "/" & strarrdate(2)
                fpsDkNgay.Col = 2
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 3
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 4
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 5
                fpsDkNgay.ColHidden = True
                
                fpsDkNgay.Col = 7
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 8
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 9
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 10
                fpsDkNgay.ColHidden = True
            End With
            Exit Sub
        End If
        'xmlDocument.Load TAX_Utilities_New.GetAbsolutePath("menu.xml")
        'Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            For Each xmlNode In xmlNodeListMenu
            If arrStrId(.TypeComboBoxCurSel) = "KHBS" Then
                lstryear = "1"
                lstrMonth = "1"
                lstrThreemonths = "0"
                Exit For
            End If
                Dim Parentid As String
                Parentid = GetAttribute(xmlNode, "PopID")
                '.TypeComboBoxIndex = 0
                If Parentid = "101" Then
                    If arrStrId(.TypeComboBoxCurSel) = GetAttribute(xmlNode, "ID") Then
                        lstryear = GetAttribute(xmlNode, "Year")
                        lstrMonth = GetAttribute(xmlNode, "Month")
                        lstrThreemonths = GetAttribute(xmlNode, "ThreeMonth")
                        Exit For
                    End If
                End If
            Next
        CreateDkKy
    End With
End Sub

Private Sub fpsLoaiTK_GotFocus()
    btnTracuu.Default = True
End Sub

Private Sub fpsLoaiTK_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 0 Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteTuNCol), f1dteTuNRow
        End With
    End If
    If KeyCode = vbKeyTab And Shift = 1 Then
        btnThoat.SetFocus
    End If
End Sub
Sub formatPrefix(strDate As String, strarrdate() As String)
    Dim lstrdate As Variant
    Dim i As Integer
    i = 0
    strarrdate = Split(strDate, "/")
    For Each lstrdate In strarrdate
        If Len(Trim(lstrdate)) < 2 Then
                strarrdate(i) = "0" & Trim(lstrdate)
         End If
         i = i + 1
    Next
End Sub
Sub CreateDkKy()
    Dim vdtehientai As String
    vdtehientai = format(Date, "dd/mm/yyyy")
    Dim strarrdate() As String
    formatPrefix vdtehientai, strarrdate
    
    With fpsDkNgay
        If lstrMonth = "0" And lstrThreemonths = "0" Then
            If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Then
                .Col = .ColLetterToNumber(f1dteTuNCol)
                .Row = f1dteTuNRow
                .CellType = CellTypePic
                .TypePicMask = "99//99//9999"
                .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
                
                .Col = .ColLetterToNumber(f1dteDeNCol)
                .Row = f1dteDeNRow
                .CellType = CellTypePic
                .TypePicMask = "99//99//9999"
                .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
                fpsDkNgay.Col = 2
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 3
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 4
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 5
                fpsDkNgay.ColHidden = True
                
                fpsDkNgay.Col = 7
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 8
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 9
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 10
                fpsDkNgay.ColHidden = True
                Exit Sub
            
            Else
                .Col = .ColLetterToNumber(f1dteTuNCol)
                .Row = f1dteTuNRow
                .CellType = CellTypePic
                .TypePicMask = "9999"
                .Text = strarrdate(2)
                
                .Col = .ColLetterToNumber(f1dteDeNCol)
                .Row = f1dteDeNRow
                .CellType = CellTypePic
                .TypePicMask = "9999"
                .Text = strarrdate(2)
                
                fpsDkNgay.Col = 2
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 3
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 4
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 5
                fpsDkNgay.ColHidden = True
                
                fpsDkNgay.Col = 7
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 8
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 9
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 10
                fpsDkNgay.ColHidden = True
                Exit Sub
             End If
        End If
        If lstrMonth = "1" Then
             .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strarrdate(1) & "/" & strarrdate(2)
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strarrdate(1) & "/" & strarrdate(2)
            
            fpsDkNgay.Col = 2
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 3
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 4
            fpsDkNgay.ColHidden = False
            fpsDkNgay.Col = 5
            fpsDkNgay.ColHidden = True
            
            fpsDkNgay.Col = 7
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 8
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 9
            fpsDkNgay.ColHidden = False
            fpsDkNgay.Col = 10
            fpsDkNgay.ColHidden = True
        
            Exit Sub
        End If
        If lstrThreemonths = "1" Then
            Dim strQuy As String
            If Val(strarrdate(0)) < 4 Then
                strQuy = "01"
            ElseIf Val(strarrdate(0)) >= 4 And Val(strarrdate(0)) < 7 Then
                 strQuy = "02"
            ElseIf Val(strarrdate(0)) >= 7 And Val(strarrdate(0)) < 10 Then
                strQuy = "03"
            Else
                strQuy = "04"
            End If
             .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            
            If Val(strQuy) = "04" Then
                .Text = strQuy & "/" & Val(strarrdate(2)) - 1
            Else
                .Text = strQuy & "/" & strarrdate(2)
            End If
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            
            If Val(strQuy) = "04" Then
                .Text = strQuy & "/" & Val(strarrdate(2)) - 1
            Else
                .Text = strQuy & "/" & strarrdate(2)
            End If
            
            fpsDkNgay.Col = 2
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 3
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 4
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 5
            fpsDkNgay.ColHidden = False
            
            fpsDkNgay.Col = 7
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 8
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 9
            fpsDkNgay.ColHidden = True
            fpsDkNgay.Col = 10
            fpsDkNgay.ColHidden = False
            Exit Sub
        End If
    End With
End Sub
Private Function KiemTraDKngay() As Boolean
Dim strarrdate() As String
With fpsDkNgay
        KiemTraDKngay = True
        
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        If .Text = "" Then
            Dtetun = "199000"
        Else
            formatPrefix .Text, strarrdate
            If lstrMonth <> "0" Or lstrThreemonths <> 0 Then
                Dtetun = strarrdate(1) & strarrdate(0)
            Else
                Dtetun = strarrdate(0) & "00"
            End If
        End If
        
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        
        If Trim(Replace(.Text, "/", "")) = "" Then
            Dteden = "999099"
        Else
            formatPrefix .Text, strarrdate
            If lstrMonth <> "0" Or lstrThreemonths <> "0" Then
                Dteden = strarrdate(1) & strarrdate(0)
            Else
                Dteden = strarrdate(0) & "99"
            End If
        End If
        If Dtetun > Dteden Then
            If lstrMonth = "0" And lstrThreemonths = "0" Then
                DisplayMessage "0093", msOKOnly, miInformation
            End If
            If lstrMonth = "1" Then
                DisplayMessage "0097", msOKOnly, miInformation
            End If
            If lstrThreemonths = "1" Then
                DisplayMessage "0098", msOKOnly, miInformation
            End If
            KiemTraDKngay = False
            .SetFocus
            .SetActiveCell .ColLetterToNumber(f1dteTuNCol), f1dteDeNRow
        End If
    End With
End Function

'*********************************************************
' Description: Lay thong tin ve tat ca cac to khai theo
'              kieu kien tra cuu dua vao va ma to khai.
'       strId: Ma to khai
'       strPeriodFrom: Ky ke khai tu ky.
'       strPeriodTo: Ky ke khai den ky.
'       strPeriods(): Thong tin ve tat ca cac to khai, bao gom:
'           + Ma to khai
'           + Ten to khai
'           + Ky hieu luc
'           + Ky ke khai
'           + Ky ke khai tu ngay (neu co)
'           + Ky ke khai den ngay (neu co)
'           + Ten cac file du lieu
'           + Gia tri thue phai nop(neu co).
'           + Gia tri thue khau tru(neu co).
' Return: True neu ton tai to khai tuong ung voi ma,
'         False neu nguoc lai
'*********************************************************
Private Function GetTaxReportsById(ByVal strId As String, ByVal strPeriodFrom As String, ByVal strPeriodTo As String, _
        ByRef strPeriods() As String) As Boolean
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Dim xmlNodeMenu As MSXML.IXMLDOMNode
    Dim xmlNodeValid As MSXML.IXMLDOMNode
    Dim strReturn() As String, strPeriodReturn As String
    Dim lngIndex As Long, lngIndex2 As Long
    Dim blnReturn As Boolean, blnValidFinanceYear As Boolean
    Dim strKieu_Ky As String, strNextPeriod As String
    Dim strThueKhauTruId As String, strThuePhaiNopId As String
    Dim strDataFile As String
    
    blnReturn = False
    
    xmlDocument.Load TAX_Utilities_New.GetAbsolutePath("map.xml")
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
    
    'Khoi tao gia tri khoang tra cuu
    If Len(strPeriodFrom) = 4 Then
        strPeriodFrom = "01/01/" & strPeriodFrom
    Else
        If strId = "80" Or strId = "82" Then
        Else
            strPeriodFrom = "01/" & strPeriodFrom
        End If
    End If
    
    If Len(strPeriodTo) = 4 Then
        strPeriodTo = "01/12/" & strPeriodTo
    Else
        If strId = "80" Or strId = "82" Then
        Else
            strPeriodTo = "01/" & strPeriodTo
        End If
    End If
    
    
    
    
    'Lay menu node
    For Each xmlNodeMenu In xmlNodeListMap
        If GetAttribute(xmlNodeMenu, "ID") = strId Then
            blnReturn = True
            Exit For
        End If
    Next
    
    If GetAttribute(xmlNodeMenu, "ID") = "KHBS" Then
        blnReturn = SearchKHBS(strPeriodFrom, strPeriodTo, strPeriods())
        Exit Function
    End If
    
    'Lay kieu ky ke khai
    If GetAttribute(xmlNodeMenu, "Month") = "1" Then
        strKieu_Ky = KIEU_KY_THANG
    ElseIf GetAttribute(xmlNodeMenu, "ThreeMonth") = "1" Then
        strKieu_Ky = KIEU_KY_QUY
    ElseIf GetAttribute(xmlNodeMenu, "Day") = "1" Then
        strKieu_Ky = KIEU_KY_NGAY_NAM
    Else
        strKieu_Ky = KIEU_KY_NAM
    End If
    
    'Kiem tra ngay bat dau nam tai chinh
    If GetAttribute(xmlNodeMenu, "FinanceYear") = "1" Then
        strNgayTaiChinh = GetNgayBatDauNamTaiChinh
        If Not KiemTraNgayTaiChinh(strNgayTaiChinh, False) Then
            blnValidFinanceYear = False
        Else
            iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
            iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
            blnValidFinanceYear = True
        End If
    Else
        strNgayTaiChinh = "01/01"
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
        blnValidFinanceYear = True
    End If
    
    'Khoi  tao gia tri cho bien lngIndex2
    lngIndex2 = UBound(strPeriods()) + 1
    
    'lay validity node
    For lngIndex = 0 To UBound(arrStrXMLFileNames)
        For Each xmlNodeValid In xmlNodeMenu.childNodes
            
            'Lay gia tri ky hieu luc tiep theo
            If Not xmlNodeValid.nextSibling Is Nothing Then
                strNextPeriod = GetAttribute(xmlNodeValid.nextSibling, "StartDate")
            Else
                strNextPeriod = vbNullString
            End If
            
            'Lay ten mau data file
            strDataFile = GetAttribute(xmlNodeValid, "DataFile")
            
            'Lay cac Id chua gia tri thue khau tru va thue phai nop
            strThueKhauTruId = GetAttribute(xmlNodeValid, "ThueKhauTru")
            strThuePhaiNopId = GetAttribute(xmlNodeValid, "ThuePhaiNop")
            
            If InStr(1, arrStrXMLFileNames(lngIndex), Split(strDataFile, ",")(0)) = 1 Then
                    If IsValidPeriod(strKieu_Ky, strDataFile, strNextPeriod, arrStrXMLFileNames(lngIndex), strPeriodFrom, strPeriodTo, blnValidFinanceYear, strPeriodReturn) Then
                        ReDim Preserve strPeriods(lngIndex2)
                        strPeriods(lngIndex2) = strId & _
                            "~" & GetAttribute(xmlNodeMenu, "Caption") & _
                            "~" & GetAttribute(xmlNodeValid, "StartDate") & _
                            "~" & strPeriodReturn & _
                            "~" & GetTaxValue(arrStrXMLFileNames(lngIndex), strThueKhauTruId, False) & _
                            "~" & GetTaxValue(arrStrXMLFileNames(lngIndex), strThuePhaiNopId, False)
                            '"~" & arrStrXMLFileNames(lngIndex)
                        lngIndex2 = lngIndex2 + 1
                    End If
            End If
            If InStr(1, arrStrXMLFileNames(lngIndex), Split(strDataFile, ",")(0)) = 5 And Left(arrStrXMLFileNames(lngIndex), 2) = "bs" Then
                    strDataFile = Split(arrStrXMLFileNames(lngIndex), "_")(0) & "_" & GetAttribute(xmlNodeValid, "DataFile")
                    If IsValidPeriod(strKieu_Ky, strDataFile, strNextPeriod, arrStrXMLFileNames(lngIndex), strPeriodFrom, strPeriodTo, blnValidFinanceYear, strPeriodReturn) Then
                        ReDim Preserve strPeriods(lngIndex2)
                        strPeriods(lngIndex2) = strId & _
                            "~" & GetAttribute(xmlNodeMenu, "Caption") & " - BS lan " & Right(Split(strDataFile, "_")(0), Len(Split(strDataFile, "_")(0)) - 2) & _
                            "~" & GetAttribute(xmlNodeValid, "StartDate") & _
                            "~" & GetDataFileNamesBS(strPeriodReturn) & _
                            "~" & GetTaxValue(arrStrXMLFileNames(lngIndex), strThueKhauTruId, False) & _
                            "~" & GetTaxValue(arrStrXMLFileNames(lngIndex), strThuePhaiNopId, False)
                            '"~" & arrStrXMLFileNames(lngIndex)
                        lngIndex2 = lngIndex2 + 1
                    End If
            End If
        Next
    Next lngIndex
    
    GetTaxReportsById = blnReturn
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetTaxReportsById", Err.Number, Err.Description
End Function

Private Sub LoadXMLFileNames()
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As file
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_New.DataFolder)).Files
        If Right$(fle.Name, 4) = ".xml" Then
            ReDim Preserve arrStrXMLFileNames(lngIndex)
            arrStrXMLFileNames(lngIndex) = Mid$(fle.Name, 1, Len(fle.Name) - 4)
            lngIndex = lngIndex + 1
        End If
    Next
End Sub

Private Function IsValidPeriod(ByVal strKieu_Ky As String, ByVal strDataFile As String, _
                                ByVal strNextPeriod As String, ByVal strFileName As String, _
                                ByVal strPeriodFrom As String, ByVal strPeriodTo As String, _
                                ByVal blnValidFinanceYear As Boolean, ByRef strPeriod As String) As Boolean
    Dim lStrPeriod As String
    Dim dPeriod As Date, dNextPeriod As Date
    Dim dPeriodFrom As Date, dPeriodTo As Date
    Dim dNgayDauQuy As Date, dNgayCuoiQuy As Date
    Dim dNgayDau As Date, dNgayCuoi As Date
    Dim objDateUtils As DateUtils
    Dim strDataFiles() As String                 'Luu ten mau cua cac sheet
    
    strDataFiles = Split(strDataFile, ",")
    
    ' TK 02/NTNN, 04/NTNN xu ly khac
    If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Or InStr(strFileName, "02_NTNN") > 0 Or InStr(strFileName, "04_NTNN") > 0 Then
        dPeriodFrom = DateSerial(CInt(Mid$(strPeriodFrom, 7, 4)), CInt(Mid$(strPeriodFrom, 4, 2)), CInt(Mid$(strPeriodFrom, 1, 2)))
        dPeriodTo = DateSerial(CInt(Mid$(strPeriodTo, 7, 4)), CInt(Mid$(strPeriodTo, 4, 2)), CInt(Mid$(strPeriodTo, 1, 2)))
    Else
        dPeriodFrom = DateSerial(CInt(Mid$(strPeriodFrom, 7, 4)), CInt(Mid$(strPeriodFrom, 4, 2)), CInt(Mid$(strPeriodFrom, 1, 2)))
        dPeriodTo = DateSerial(CInt(Mid$(strPeriodTo, 7, 4)), CInt(Mid$(strPeriodTo, 4, 2)), CInt(Mid$(strPeriodTo, 1, 2)))
    End If
    
    If strNextPeriod <> vbNullString Then
        dNextPeriod = DateSerial(CInt(Mid$(strNextPeriod, 7, 4)), CInt(Mid$(strNextPeriod, 4, 2)), CInt(Mid$(strNextPeriod, 1, 2)))
    Else
        dNextPeriod = DateSerial(9990, 12, 30)
    End If
    
    Select Case strKieu_Ky
        Case KIEU_KY_THANG
            lStrPeriod = Right$(strFileName, 6)
            
            If strDataFiles(0) <> Mid(strFileName, 1, Len(strFileName) - Len(lStrPeriod) - 1) Then
            ' This is not valid file name
                Exit Function
            End If
            
            If IsNumeric(Left(lStrPeriod, 2)) And Val(Left(lStrPeriod, 2)) >= 1 And Val(Left(lStrPeriod, 2)) <= 12 Then
                If IsNumeric(Right(lStrPeriod, 4)) And Val(Right(lStrPeriod, 4)) >= 2000 Then
                    dPeriod = DateSerial(CInt(Right(lStrPeriod, 4)), CInt(Left(lStrPeriod, 2)), 1)
                    If dPeriod >= dPeriodFrom And _
                            dPeriod <= dPeriodTo Then
                        'Lay ky ke khai
                        strPeriod = Mid$(lStrPeriod, 1, 2) & "/" & Mid$(lStrPeriod, 3, 4) & "~" & "~"
                        
                        'Lay trang thai to khai
                        dPeriod = GetNgayCuoiThang(CInt(Right(lStrPeriod, 4)), CInt(Left(lStrPeriod, 2)))
                        If dPeriod >= dNextPeriod Then
                            strPeriod = strPeriod & "~False~" & LOI_KY_HIEU_LUC
                        Else
                            strPeriod = strPeriod & "~True~"
                        End If
                        
                        'Lay danh sach ten cac file data
                        strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod)
                        IsValidPeriod = True
                        Exit Function
                    End If
                End If
            End If
        Case KIEU_KY_QUY
            lStrPeriod = Right$(strFileName, 6)
            
            If strDataFiles(0) <> Mid(strFileName, 1, Len(strFileName) - Len(lStrPeriod) - 1) Then
            ' This is not valid file name
                Exit Function
            End If
            
            If IsNumeric(Left(lStrPeriod, 2)) And Val(Left(lStrPeriod, 2)) >= 1 And Val(Left(lStrPeriod, 2)) <= 4 Then
                If IsNumeric(Right(lStrPeriod, 4)) And Val(Right(lStrPeriod, 4)) >= 2000 Then
                    dPeriod = DateSerial(CInt(Right(lStrPeriod, 4)), CInt(Left(lStrPeriod, 2)), 1)
                    If dPeriod >= dPeriodFrom And _
                            dPeriod <= dPeriodTo Then
                        strPeriod = Mid$(lStrPeriod, 1, 2) & "/" & Mid$(lStrPeriod, 3, 4) & "~" & "~" '& "~True"
                        
                        If blnValidFinanceYear Then
                            'Lay trang thai to khai
                            dPeriod = GetNgayCuoiQuy(CInt(Left(lStrPeriod, 2)), CInt(Right(lStrPeriod, 4)), iNgayTaiChinh, iThangTaiChinh)
                            If dPeriod >= dNextPeriod Then
                                strPeriod = strPeriod & "~False~" & LOI_KY_HIEU_LUC
                            Else
                                strPeriod = strPeriod & "~True~"
                            End If
                        Else
                            strPeriod = strPeriod & "~False~" & LOI_NGAY_BAT_DAU_NAM_TAI_CHINH
                        End If
                        
                        'Lay danh sach ten cac file data
                        strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod)
                        IsValidPeriod = True
                        Exit Function
                    End If
                End If
            End If
        Case KIEU_KY_NAM
            lStrPeriod = Right$(strFileName, 4)
            
            If strDataFiles(0) <> Mid(strFileName, 1, Len(strFileName) - Len(lStrPeriod) - 1) Then
            ' This is not valid file name
                Exit Function
            End If
            
            If IsNumeric(lStrPeriod) And Val(lStrPeriod) >= 2000 Then
                dPeriod = DateSerial(CInt(lStrPeriod), 1, 1)
                If dPeriod >= dPeriodFrom And _
                        dPeriod <= dPeriodTo Then
                    strPeriod = lStrPeriod & "~" & "~" '& "~True"
                    
                    If blnValidFinanceYear Then
                        'Lay trang thai to khai
                        dPeriod = NgayCuoiNamTaiChinh(CInt(lStrPeriod), iThangTaiChinh, iNgayTaiChinh)
                        If dPeriod >= dNextPeriod Then
                            strPeriod = strPeriod & "~False~" & LOI_KY_HIEU_LUC
                        Else
                            strPeriod = strPeriod & "~True~"
                        End If
                    Else
                        strPeriod = strPeriod & "~False~" & LOI_NGAY_BAT_DAU_NAM_TAI_CHINH
                    End If
                    
                    'Lay danh sach ten cac file data
                    strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod)
                    IsValidPeriod = True
                    Exit Function
                End If
            End If
        Case KIEU_KY_NGAY_NAM
            ' TK 02/NTNN, 04/NTNN xu ly khac
            If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "80" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "82" Or InStr(strFileName, "02_NTNN") > 0 Or InStr(strFileName, "04_NTNN") > 0 Then
                lStrPeriod = Right$(strFileName, 17)
                
                If strDataFiles(0) <> Mid(strFileName, 1, Len(strFileName) - Len(lStrPeriod) - 1) Then
                ' This is not valid file name
                    Exit Function
                End If
                    
                If IsValidDate(Mid$(lStrPeriod, 1, 8)) <> "" And IsValidDate(Mid$(lStrPeriod, 10, 8)) <> "" Then
                    dPeriod = DateSerial(CInt(Mid$(lStrPeriod, 5, 4)), CInt(Mid$(lStrPeriod, 3, 2)), CInt(Mid$(lStrPeriod, 1, 2)))
                    If dPeriod >= dPeriodFrom And _
                            dPeriod <= dPeriodTo Then
                        strPeriod = Mid$(lStrPeriod, 1, 2) & "/" & Mid$(lStrPeriod, 3, 2) & "/" & Mid$(lStrPeriod, 5, 4) & "~" & IsValidDate(Mid$(lStrPeriod, 1, 8)) & "~" & IsValidDate(Mid$(lStrPeriod, 10, 8))  '& "~True"
                        
                        If blnValidFinanceYear Then
                            'Lay trang thai to khai
                            dPeriod = NgayCuoiNamTaiChinh(CInt(Mid$(lStrPeriod, 5, 4)), iThangTaiChinh, iNgayTaiChinh)
                            If dPeriod >= dNextPeriod Then
                                strPeriod = strPeriod & "~False~" & LOI_KY_HIEU_LUC
                            Else
                                Set objDateUtils = New DateUtils
                                dNgayDauQuy = GetNgayDauQuy(4, CInt(Mid$(lStrPeriod, 5, 4)) - 1, iNgayTaiChinh, iThangTaiChinh)
                                dNgayCuoiQuy = GetNgayCuoiQuy(1, CInt(Mid$(lStrPeriod, 5, 4)) + 1, iNgayTaiChinh, iThangTaiChinh)
                                dNgayDau = objDateUtils.ToDate(IsValidDate(Mid$(lStrPeriod, 1, 8)), "DD/MM/YYYY")
                                dNgayCuoi = objDateUtils.ToDate(IsValidDate(Mid$(lStrPeriod, 10, 8)), "DD/MM/YYYY")
                                
                                'Kiem tra gia tri tu ngay va den ngay
                                If dNgayDau < dNgayDauQuy Or dNgayCuoi > dNgayCuoiQuy _
                                   Or DateDiff("M", dNgayDau, dNgayCuoi) + 1 > 15 Or dNgayCuoi < dNgayDau Then
                                     strPeriod = strPeriod & "~False~" & LOI_TU_NGAY_DEN_NGAY
                                Else
                                    strPeriod = strPeriod & "~True~"
                                End If
                            End If
                        Else
                            strPeriod = strPeriod & "~False~" & LOI_NGAY_BAT_DAU_NAM_TAI_CHINH
                        End If
                        
                        'Lay danh sach ten cac file data
                        strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod)
                        IsValidPeriod = True
                        Exit Function
                    End If
                End If
                
            Else
                lStrPeriod = Right$(strFileName, 22)
                
                If strDataFiles(0) <> Mid(strFileName, 1, Len(strFileName) - Len(lStrPeriod) - 1) Then
                ' This is not valid file name
                    Exit Function
                End If
                            
                If IsNumeric(Mid$(lStrPeriod, 1, 4)) And Val(Mid$(lStrPeriod, 1, 4)) >= 2000 Then
                    
                    If IsValidDate(Mid$(lStrPeriod, 6, 8)) <> "" And IsValidDate(Mid$(lStrPeriod, 15, 8)) <> "" Then
                        dPeriod = DateSerial(CInt(Mid$(lStrPeriod, 1, 4)), 1, 1)
                        If dPeriod >= dPeriodFrom And _
                                dPeriod <= dPeriodTo Then
                            strPeriod = Mid$(lStrPeriod, 1, 4) & "~" & IsValidDate(Mid$(lStrPeriod, 6, 8)) & "~" & IsValidDate(Mid$(lStrPeriod, 15, 8)) '& "~True"
                            
                            If blnValidFinanceYear Then
                                'Lay trang thai to khai
                                dPeriod = NgayCuoiNamTaiChinh(CInt(Mid$(lStrPeriod, 1, 4)), iThangTaiChinh, iNgayTaiChinh)
                                If dPeriod >= dNextPeriod Then
                                    strPeriod = strPeriod & "~False~" & LOI_KY_HIEU_LUC
                                Else
                                    Set objDateUtils = New DateUtils
                                    dNgayDauQuy = GetNgayDauQuy(4, CInt(Mid$(lStrPeriod, 1, 4)) - 1, iNgayTaiChinh, iThangTaiChinh)
                                    dNgayCuoiQuy = GetNgayCuoiQuy(1, CInt(Mid$(lStrPeriod, 1, 4)) + 1, iNgayTaiChinh, iThangTaiChinh)
                                    dNgayDau = objDateUtils.ToDate(IsValidDate(Mid$(lStrPeriod, 6, 8)), "DD/MM/YYYY")
                                    dNgayCuoi = objDateUtils.ToDate(IsValidDate(Mid$(lStrPeriod, 15, 8)), "DD/MM/YYYY")
                                    
                                    'Kiem tra gia tri tu ngay va den ngay
                                    If dNgayDau < dNgayDauQuy Or dNgayCuoi > dNgayCuoiQuy _
                                       Or DateDiff("M", dNgayDau, dNgayCuoi) + 1 > 15 Or dNgayCuoi < dNgayDau Then
                                         strPeriod = strPeriod & "~False~" & LOI_TU_NGAY_DEN_NGAY
                                    Else
                                        strPeriod = strPeriod & "~True~"
                                    End If
                                End If
                            Else
                                strPeriod = strPeriod & "~False~" & LOI_NGAY_BAT_DAU_NAM_TAI_CHINH
                            End If
                            
                            'Lay danh sach ten cac file data
                            strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod)
                            IsValidPeriod = True
                            Exit Function
                        End If
                    End If
                End If
          End If
    End Select
End Function

Private Function IsValidDate(ByVal strDate As String) As String
    Dim dDate As Variant
    
    On Error GoTo ErrHandle
    If Len(strDate) <> 8 Then Exit Function
    
    dDate = DateValue(Mid$(strDate, 1, 2) & "/" & Mid$(strDate, 3, 2) & "/" & Mid$(strDate, 5, 4))
    IsValidDate = Mid$(strDate, 1, 2) & "/" & Mid$(strDate, 3, 2) & "/" & Mid$(strDate, 5, 4)
    
    Exit Function
ErrHandle:
End Function

Private Function GetTaxValue(ByVal strDataFileName As String, ByVal strId As String, ByVal KHBS As Boolean) As String
    Dim xmlDom As New MSXML.DOMDocument
    
    xmlDom.Load TAX_Utilities_New.DataFolder & strDataFileName & ".xml"
    If KHBS = False Then
        GetTaxValue = GetAttribute(xmlDom.nodeFromID(strId), "Value")
    Else
        GetTaxValue = GetAttribute(xmlDom.childNodes(2).childNodes(2).lastChild.lastChild, "Value")
        If Val(GetTaxValue) < 0 Then GetTaxValue = 0
    End If
    
    Set xmlDom = Nothing
    
End Function


Private Function CheckErrTaxReport(ByVal strId As String, ByVal strKyHieuLuc As String, _
                                   ByVal strKyKeKhai As String, ByVal lRow As Long) As Boolean
    Dim lngCtrl As Long, lngCtrl2 As Long, lUbound As Long
    Dim varId As Variant, varKyHieuLuc As Variant, varKyKeKhai As Variant
    
    lUbound = UBound(arrLngErrRows)
    
    With fpSKetQua
        For lngCtrl = 1 To lRow
            .GetText 3, lngCtrl, varId
            .GetText 5, lngCtrl, varKyHieuLuc
            .GetText 6, lngCtrl, varKyKeKhai
            
            If CStr(varId) = strId And _
               CStr(varKyKeKhai) = strKyKeKhai And _
               CStr(varKyHieuLuc) <> strKyHieuLuc Then
                
               lUbound = lUbound + 1
               ReDim Preserve arrLngErrRows(lUbound)
               arrLngErrRows(lUbound) = lngCtrl
                For lngCtrl2 = 1 To .MaxCols
                    .Col = lngCtrl2
                    .Row = lngCtrl
                    .ForeColor = vbRed
                Next lngCtrl2
            End If
        Next lngCtrl
    End With
    
End Function

Private Function IsErrRow(ByVal lRow As Long) As Boolean
    Dim lCtrl As Long
    Dim blnReturn As Boolean
    
    blnReturn = False
    For lCtrl = 1 To UBound(arrLngErrRows)
        If lRow = arrLngErrRows(lCtrl) Then
            blnReturn = True
            Exit For
        End If
    Next lCtrl
    
    IsErrRow = blnReturn
End Function

Private Function SetRowFocus(ByVal lngRow As Long, ByVal lngNewRow As Long, Optional ByVal blnClickEvent As Boolean = False) As Long
    With fpSKetQua
        .Col = -1
        .Row = lngRow
        .BackColor = vbWhite
        
        .Row = lngNewRow
        .BackColor = RGB(212, 343, 423)
        
        If blnClickEvent Then
            .SetActiveCell 2, lngNewRow
        Else
            .SetActiveCell 2, lngRow
        End If
        
    End With
    SetRowFocus = lngNewRow
End Function


Private Function GetDataFileNames(arrStrDataFiles() As String, strPeriod As String) As String
    Dim intCtrl As Integer
    Dim strReturn As String
    
    For intCtrl = 0 To UBound(arrStrDataFiles)
        If intCtrl = 0 Then
            strReturn = arrStrDataFiles(intCtrl) & "_" & strPeriod
        Else
            strReturn = strReturn & "," & arrStrDataFiles(intCtrl) & "_" & strPeriod
        End If
    Next intCtrl
    
    GetDataFileNames = strReturn
End Function


Private Function GetDataFileNamesBS(ByVal strFileNames As String) As String
    Dim intCtrl As Integer
    Dim strReturn As String
    Dim arrTempValue() As String, arrTempValue1() As String
    arrTempValue = Split(strFileNames, "~")
    For intCtrl = 0 To UBound(arrTempValue) - 1
        strReturn = strReturn & arrTempValue(intCtrl) & "~"
    Next intCtrl
    arrTempValue1 = Split(arrTempValue(UBound(arrTempValue)), ",")
    
    strReturn = strReturn & arrTempValue1(0) & "," & arrTempValue1(UBound(arrTempValue1))
    
    
    GetDataFileNamesBS = strReturn
End Function


Private Sub DeleteDataFiles(ByVal strFileNames As String)
    Dim arrStrDataFiles() As String
    Dim intCtrl As Integer
    Dim fso As New FileSystemObject
    
    arrStrDataFiles = Split(strFileNames, ",")
    
    For intCtrl = 0 To UBound(arrStrDataFiles)
        If fso.FileExists(GetAbsolutePath(TAX_Utilities_New.DataFolder & arrStrDataFiles(intCtrl) & ".xml")) Then
            fso.DeleteFile GetAbsolutePath(TAX_Utilities_New.DataFolder & arrStrDataFiles(intCtrl) & ".xml"), True
            '.DeleteRows i + 1, 1
        End If
    Next intCtrl
End Sub

Private Sub TraCuu()
    Dim lRow As Long
    Dim arrStrPeriods() As String
    Dim strFromDay As String, strToDay As String
    Dim arrStrTemp() As String
    Dim lCtrl As Long
    
        
    fpsDkNgay_LeaveCell 1, 1, 2, 1, False
    If blnDKienTraCuu = False Then
        fpsDkNgay.SetFocus
        Exit Sub
    End If
    
    'Kiem tra dk tu ky phai nho hon den ky
    If Not KiemTraDKngay Then Exit Sub
    
    'Kiem tra, khoi tao cac gia tri dieu kien tra cuu
    With fpsDkNgay
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        strFromDay = .Text
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        strToDay = .Text
    End With
    
    If Trim(Replace(strFromDay, "/", "")) = "" Then
        strFromDay = "2000"
    End If
    
    If Trim(Replace(strToDay, "/", "")) = "" Then
        strToDay = "9990"
    End If
    
    'Danh sach file trong thu muc Data
    LoadXMLFileNames
    
    ReDim Preserve arrStrPeriods(0)
    If fpsLoaiTK.TypeComboBoxCurSel <> 0 Then
        GetTaxReportsById arrStrId(fpsLoaiTK.TypeComboBoxCurSel), strFromDay, _
                    strToDay, arrStrPeriods
    Else
        For lCtrl = 1 To UBound(arrStrId)
            GetTaxReportsById arrStrId(lCtrl), strFromDay, _
                    strToDay, arrStrPeriods
        Next
    End If
    
    lSoBG = 0
    ReDim arrLngErrRows(0)
    With fpSKetQua
        .Visible = False
        
        'Xoa bo ket qua cu tren Grid
        .DeleteRows 2, fpSKetQua.MaxRows - 1
        .MaxRows = 2
        
        'Xoa trang thai nut Check
        .Col = 2
        .Row = 1
        .value = "0"
        
        lRow = 2
        For lCtrl = 1 To UBound(arrStrPeriods)
            If lRow > .MaxRows Then
                .MaxRows = .MaxRows + 1
            End If
            
            arrStrTemp = Split(arrStrPeriods(lCtrl), "~")
            
            ' STT
            lSoBG = lSoBG + 1
            .InsertRows lRow, 1
            '.SetCellBorder 1, lRow, .MaxCols, lRow, CellBorderIndexTop, 0, CellBorderStyleFineDot
            .SetText 1, lRow, lSoBG
                        
            'Check
            .Col = 2
            .Row = lRow
            .CellType = CellTypeCheckBox
            .TypeHAlign = TypeHAlignCenter
            .Lock = False
            
             'Check
            .Col = 6
            .Row = lRow
            .TypeHAlign = TypeHAlignCenter
            .Lock = True
            
            'Thue khau tru
            .Col = 9
            .CellType = CellTypeNumber
            .TypeNumberDecimal = ","
            .TypeNumberSeparator = "."
            .TypeNumberDecPlaces = 0
            .TypeNumberShowSep = True
                        
            'Thue phai nop
            .Col = 10
            .CellType = CellTypeNumber
            .TypeNumberDecimal = ","
            .TypeNumberSeparator = "."
            .TypeNumberDecPlaces = 0
            .TypeNumberShowSep = True
                        
            .SetText 3, lRow, arrStrTemp(0)     'Ma to khai
            .SetText 4, lRow, arrStrTemp(1)     'Ten to khai
            .SetText 5, lRow, arrStrTemp(2)     'Ngay hieu luc
            .SetText 6, lRow, arrStrTemp(3)     'Ky ke khai
            .SetText 7, lRow, arrStrTemp(4)     'Tu ngay
            .SetText 8, lRow, arrStrTemp(5)     'Den ngay
            .SetText 9, lRow, arrStrTemp(9)     'Thue khau tru
            .SetText 10, lRow, arrStrTemp(10)    'Thue phai nop
            .SetText 11, lRow, arrStrTemp(8)    'Danh sach ten File
            .SetText 12, lRow, arrStrTemp(7)    'Trang thai loi
            fpSKetQua.RowHeight(lRow) = fpSKetQua.MaxTextRowHeight(lRow)
            
            If arrStrTemp(6) = "False" Then
                .Col = -1
                .Row = lRow
                .ForeColor = vbRed
                ReDim Preserve arrLngErrRows(UBound(arrLngErrRows) + 1)
                arrLngErrRows(UBound(arrLngErrRows)) = lRow
            End If
            'CheckErrTaxReport arrStrTemp(0), arrStrTemp(2), _
                              arrStrTemp(3), lRow - 1
            lRow = lRow + 1
        Next
        
        .Visible = True
        
        'Hien thi dong duoc chon de mo to khai.
        .Row = 2
        .Col = 3 ' Id
        If .MaxRows > 2 Or .value <> vbNullString Then
            If lngRowFocus = 0 Then
                lngRowFocus = SetRowFocus(2, 2, True)
            ElseIf lngRowFocus > .MaxRows Then
                lngRowFocus = SetRowFocus(2, .MaxRows, True)
            Else
                lngRowFocus = SetRowFocus(2, lngRowFocus, True)
            End If
            btnMo.Enabled = True
        Else
            btnMo.Enabled = False
        End If
        
        'Setfocus to fpsKetQua
        .SetFocus
    End With
    
    'Hien thi status
    If UBound(arrLngErrRows) > 0 Then
        lblStatus.Visible = True
    Else
        lblStatus.Visible = False
    End If
End Sub


Private Function SearchKHBS(StrFromdate As String, StrToDate As String, ByRef strPeriods() As String) As Boolean
    Dim xmlNodeMenu As MSXML.IXMLDOMNode
    Dim xmlNodeValid As MSXML.IXMLDOMNode
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Dim strReturn() As String, strPeriodReturn As String
    Dim lngIndex As Long, lngIndex2 As Long
    Dim blnReturn As Boolean, blnValidFinanceYear As Boolean
    Dim dPeriodFrom As Date, dPeriodTo As Date, dKHBS As Date
    Dim sdateKHBS As String
    Dim strDataFile As String
    Dim strThuePhaiNopId As String
    xmlDocument.Load TAX_Utilities_New.GetAbsolutePath("map.xml")
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
    
    dPeriodFrom = DateSerial(CInt(Mid$(StrFromdate, 7, 4)), CInt(Mid$(StrFromdate, 4, 2)), CInt(Mid$(StrFromdate, 1, 2)))
    dPeriodTo = DateAdd("M", 1, DateSerial(CInt(Mid$(StrToDate, 7, 4)), CInt(Mid$(StrToDate, 4, 2)), CInt(Mid$(StrToDate, 1, 2))))
    'Khoi  tao gia tri cho bien lngIndex2
    lngIndex2 = UBound(strPeriods()) + 1
    For lngIndex = 0 To UBound(arrStrXMLFileNames)
            'If InStr(1, arrStrXMLFileNames(lngIndex), "KHBS_") = 1 Then
             If Left$(arrStrXMLFileNames(lngIndex), 5) = "KHBS_" Then
                strReturn = Split(arrStrXMLFileNames(lngIndex), "_")
                If Len(strReturn(3)) = 6 Then strReturn(3) = Left(strReturn(3), 2) & "/" & Right(strReturn(3), 4)
                 For Each xmlNodeMenu In xmlNodeListMap
                        If Split(GetAttribute(xmlNodeMenu.childNodes(0), "DataFile"), ",")(0) = strReturn(1) & "_" & strReturn(2) Then
                            Exit For
                        End If
                  Next
                sdateKHBS = Right(arrStrXMLFileNames(lngIndex), 8)
                dKHBS = DateSerial(CInt(Mid$(sdateKHBS, 5, 4)), CInt(Mid$(sdateKHBS, 3, 2)), CInt(Mid$(sdateKHBS, 1, 2)))
                 If dKHBS >= dPeriodFrom And _
                            dKHBS < dPeriodTo Then
                        
                        ReDim Preserve strPeriods(lngIndex2)
                        strPeriods(lngIndex2) = "KHBS_" & GetAttribute(xmlNodeMenu, "ID") & _
                            "~" & "KHBS " & GetAttribute(xmlNodeMenu, "Caption") & _
                            "~" & sdateKHBS & _
                            "~" & strReturn(3)
                            If strReturn(1) = "03" And strReturn(2) = "TNDN" Then
                                strPeriods(lngIndex2) = strPeriods(lngIndex2) & "~" & Mid(strReturn(4), 1, 2) & "/" & Mid(strReturn(4), 3, 2) & "/" & Mid(strReturn(4), 5, 4)
                                strPeriods(lngIndex2) = strPeriods(lngIndex2) & "~" & Mid(strReturn(5), 1, 2) & "/" & Mid(strReturn(5), 3, 2) & "/" & Mid(strReturn(5), 5, 4)
                                strPeriods(lngIndex2) = strPeriods(lngIndex2) & _
                                "~" & _
                                "~" & _
                                "~" & arrStrXMLFileNames(lngIndex) & "," & Replace(arrStrXMLFileNames(lngIndex), "KHBS", "KHBS1") & _
                                "~" & _
                                "~" & GetTaxValue(arrStrXMLFileNames(lngIndex), strThuePhaiNopId, True)
                            Else
                                strPeriods(lngIndex2) = strPeriods(lngIndex2) & _
                                "~" & _
                                "~" & _
                                "~" & _
                                "~" & _
                                "~" & arrStrXMLFileNames(lngIndex) & "," & Replace(arrStrXMLFileNames(lngIndex), "KHBS", "KHBS1") & _
                                "~" & _
                                "~" & GetTaxValue(arrStrXMLFileNames(lngIndex), strThuePhaiNopId, True)
                            End If
                            
                        lngIndex2 = lngIndex2 + 1
                  End If
'            ElseIf Left$(arrStrXMLFileNames(lngIndex), 2) = "BS" Then
'                strReturn = Split(arrStrXMLFileNames(lngIndex), "_")
'                If Len(strReturn(3)) = 6 Then strReturn(3) = Left(strReturn(3), 2) & "/" & Right(strReturn(3), 4)
'                 For Each xmlNodeMenu In xmlNodeListMap
'                        If Split(GetAttribute(xmlNodeMenu.childNodes(0), "DataFile"), ",")(0) = strReturn(1) & "_" & strReturn(2) Then
'                            Exit For
'                        End If
'                  Next
                
            End If
    Next lngIndex
End Function

