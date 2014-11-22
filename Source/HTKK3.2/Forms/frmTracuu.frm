VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmTraCuu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   11565
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
      Left            =   5430
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
      Left            =   7950
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
      Left            =   9210
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
      Left            =   6690
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
      Width           =   5655
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   495
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   5415
         _Version        =   458752
         _ExtentX        =   9551
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
      Caption         =   "Chän kú tÝnh thuÕ"
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
      Left            =   5730
      TabIndex        =   8
      Top             =   420
      Width           =   5745
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5445
         _Version        =   458752
         _ExtentX        =   9604
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
         MaxCols         =   16
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmTracuu.frx":04A1
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
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   11445
      Begin FPUSpreadADO.fpSpread fpSKetQua 
         Height          =   3225
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11145
         _Version        =   458752
         _ExtentX        =   19659
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
         SpreadDesigner  =   "frmTracuu.frx":0C07
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
      Width           =   10335
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
Private Const fpsDkNgayRow = 2
Private Const fpsDkNgayColF = "G"
Private Const fpsDkNgayColT = "M"
Private Const fpsDkNgayColXB = "O"
Private Const fpsLoaiTkRow = 2
Private Const fpsLoaiTkCol = "C"
Private Const mFormColor = -2147483633
Private Const mHeaderColor = 16709097
Private Const LOI_KY_HIEU_LUC = "1"
Private Const LOI_NGAY_BAT_DAU_NAM_TAI_CHINH = "2"
Private Const LOI_TU_NGAY_DEN_NGAY = "3"
Private lstryear As String
Private lstrMonth As String
Private lstrThreemonths As String

Private lstrDay As String

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

Private strTkGTGT As String
Private strTkDK As String
Private strGtgtIdTmp As String

Private tkNode As MSXML.IXMLDOMNode

Private Sub btnMo_Click()
    Dim frmTK       As frmInterfaces
    Dim strxoa      As Variant, varErrDesc As Variant
    Dim varId       As Variant, varPeriod As Variant
    Dim varFirstDay As Variant, varLastDay As Variant, vCheckStatus As Variant
    Dim varFileName As Variant
    Dim i, j As Integer
    Dim varDateKHBS As Variant, strFileName As Variant
    Dim LoaiTk      As Variant
    Dim tkThangQuy  As Variant
    Dim LanXB       As Variant

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
        .GetText 13, lngRowFocus, LoaiTk   ' get loai tk
        .GetText 14, lngRowFocus, tkThangQuy ' get loai tk thang/quy
        .GetText 15, lngRowFocus, LanXB ' get lan xuat ban

        ResetAllProperty

        If Left(varId, 4) = "KHBS" Then
            varId = Right(varId, 2)
            strKHBS = "frmKHBS_BS"
            TAX_Utilities_v1.DateKHBS = varDateKHBS
        End If

        If Left(varFileName, 2) = "bs" Then
            strKHBS = "TKBS"
            strSolanBS = Right(Split(varFileName, "_")(0), Len(Split(varFileName, "_")(0)) - 2)
            ngayLapTkBs = getKHBSDate(CStr(varFileName))
        Else

            ' Neu la loai to khai TNCN thi dat trang thai cua strKHBS ="TKCT"
            If Trim(varId) = "46" Or Trim(varId) = "47" Or Trim(varId) = "48" Or Trim(varId) = "49" Or Trim(varId) = "15" Or Trim(varId) = "16" Or Trim(varId) = "53" Or Trim(varId) = "37" Or Trim(varId) = "50" Or Trim(varId) = "51" Or Trim(varId) = "54" Or Trim(varId) = "38" Or Trim(varId) = "39" Or Trim(varId) = "40" Or Trim(varId) = "36" Or Trim(varId) = "70" Or Trim(varId) = "17" Or Trim(varId) = "41" Or Trim(varId) = "42" Or Trim(varId) = "43" Then
                strKHBS = "TKCT"
            End If
            
        End If
        
        varId = Left$(varId, 2)
        
        TAX_Utilities_v1.NodeMenu = getNode(CStr(varId))

        ' 12110211 xu ly to khai BS
        If strKHBS = "TKBS" Then

            For i = 1 To TAX_Utilities_v1.NodeMenu.childNodes(0).childNodes.length - 1

                If i = TAX_Utilities_v1.NodeMenu.childNodes(0).childNodes.length - 1 Then
                    SetAttribute TAX_Utilities_v1.NodeMenu.childNodes(0).childNodes(i), "Active", "1"
                Else
                    SetAttribute TAX_Utilities_v1.NodeMenu.childNodes(0).childNodes(i), "Active", "0"
                End If

            Next
            
        End If
        
        If LoaiTk = KIEU_KY_THANG Then
            TAX_Utilities_v1.month = Mid$(CStr(varPeriod), 1, 2)
            TAX_Utilities_v1.Year = Mid$(CStr(varPeriod), 4, 4)

            If tkThangQuy = "1" Then
                strQuy = "TK_THANG"
                TAX_Utilities_v1.ThreeMonths = CInt(Mid$(CStr(varPeriod), 1, 2))
            End If

            strLoaiTKThang_PS = "TK_THANG"
        ElseIf LoaiTk = "KTN" Then
            TAX_Utilities_v1.month = Mid$(CStr(varPeriod), 1, 2)
            TAX_Utilities_v1.Year = Mid$(CStr(varPeriod), 4, 4)

            strQuy = "TK_THANG"

            strLoaiTkDk = LoaiTk
        ElseIf LoaiTk = KIEU_KY_QUY Then
            
            TAX_Utilities_v1.ThreeMonths = CInt(Mid$(CStr(varPeriod), 1, 2))
            TAX_Utilities_v1.Year = Mid$(CStr(varPeriod), 4, 4)
            strQuy = "TK_QUY"
            If tkThangQuy = "1" Then
                TAX_Utilities_v1.month = CInt(Mid$(CStr(varPeriod), 1, 2))
            End If
        ElseIf LoaiTk = KIEU_KY_NGAY_PS Then
            TAX_Utilities_v1.Day = Mid$(CStr(varPeriod), 1, 2)
            TAX_Utilities_v1.month = Mid$(CStr(varPeriod), 4, 2)
            TAX_Utilities_v1.Year = Right(CStr(varPeriod), 4)
            TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
            TAX_Utilities_v1.LastDay = CStr(varLastDay)
            strQuy = "TK_LANPS"
            strLoaiTKThang_PS = "TK_LANPS"
        ElseIf LoaiTk = "CD" Or LoaiTk = "DT" Then
            TAX_Utilities_v1.Day = Mid$(CStr(varPeriod), 1, 2)
            TAX_Utilities_v1.month = Mid$(CStr(varPeriod), 4, 2)
            TAX_Utilities_v1.Year = Right(CStr(varPeriod), 4)
            TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
            TAX_Utilities_v1.LastDay = CStr(varLastDay)
            strQuy = "TK_LANXB"
            strLoaiTkDk = LoaiTk
            strSoLanXuatBan = LanXB
        ElseIf LoaiTk = KIEU_KY_NGAY_NAM Then
            TAX_Utilities_v1.Year = Right(CStr(varPeriod), 4)
            TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
            TAX_Utilities_v1.LastDay = CStr(varLastDay)
        ElseIf LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then
            TAX_Utilities_v1.Year = CStr(varPeriod)
            TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
            TAX_Utilities_v1.LastDay = CStr(varLastDay)
        ElseIf LoaiTk = "K" Then
            TAX_Utilities_v1.ThreeMonths = CInt(Mid$(CStr(varPeriod), 1, 2))
            TAX_Utilities_v1.Year = Mid$(CStr(varPeriod), 4, 4)
            TAX_Utilities_v1.FirstDay = getCellValue(Left$(varFileName, InStr(varFileName, ",") - 1), "D_17")
            TAX_Utilities_v1.LastDay = getCellValue(Left$(varFileName, InStr(varFileName, ",") - 1), "E_17")
        ElseIf LoaiTk = "N" Then
            TAX_Utilities_v1.ThreeMonths = CInt(Mid$(CStr(varPeriod), 1, 2))
            TAX_Utilities_v1.Year = Mid$(CStr(varPeriod), 4, 4)

        ElseIf LoaiTk = KIEU_KY_THANG_NAM Then
            TAX_Utilities_v1.month = Mid$(CStr(varPeriod), 1, 2)
            TAX_Utilities_v1.Year = Mid$(CStr(varPeriod), 4, 4)
            TAX_Utilities_v1.FirstDay = varFirstDay
            TAX_Utilities_v1.LastDay = varLastDay
            strQuy = "TK_TU_THANG"
            strLoaiTKThang_PS = "TK_TU_THANG"
            TAX_Utilities_v1.ThreeMonths = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh).q
            
        ElseIf LoaiTk = "KTN_Y" Then
            TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
            TAX_Utilities_v1.LastDay = CStr(varLastDay)
            TAX_Utilities_v1.Year = CStr(varPeriod)
            strLoaiTkDk = Left$(LoaiTk, 3)
            
        ElseIf LoaiTk = "CD_Y" Or LoaiTk = "DT_Y" Then
            TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
            TAX_Utilities_v1.LastDay = CStr(varLastDay)
            TAX_Utilities_v1.Year = CStr(varPeriod)
            strLoaiTkDk = Left$(LoaiTk, 2)
        Else
            'If varId = "87" Then
                TAX_Utilities_v1.FirstDay = CStr(varFirstDay)
                TAX_Utilities_v1.LastDay = CStr(varLastDay)
                TAX_Utilities_v1.Year = CStr(varPeriod)
            'Else
'                TAX_Utilities_v1.Year = CStr(varPeriod)
'            End If
        End If

        TAX_Utilities_v1.NodeValidity = GetValidityNode()
        hanNopTk = GetHanNopTk()

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

        ' to khai BC26
        If varId = "68" Then
            strQuy = "TK_QUY"
        End If
        
        Set frmTK = New frmInterfaces
        frmTK.Show
        
         If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
            frmTK.cmdLoadToKhai.Visible = True
            frmTK.cmdInsert.Visible = False
            'frmTK.cmdDelete.Visible = False
            frmTK.cmdDelete.Left = frmTK.Frame1.Width - 12100
            frmTK.cmdKiemTra.Visible = True
        End If
        
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
'                If fso.FileExists(GetAbsolutePath(TAX_Utilities_v1.DataFolder & strDataFile & ".xml")) Then
'                    fso.DeleteFile GetAbsolutePath(TAX_Utilities_v1.DataFolder & strDataFile & ".xml"), True
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
        .SetActiveCell .ColLetterToNumber(fpsLoaiTkCol), fpsLoaiTkRow
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
        .Col = .ColLetterToNumber(fpsDkNgayColF)
        .Row = fpsDkNgayRow
        .BackColor = vbWhite
        .CellType = CellTypePic
        .TypePicMask = "9999"
        
        .Col = .ColLetterToNumber(fpsDkNgayColT)
        .Row = fpsDkNgayRow
        .BackColor = vbWhite
        .CellType = CellTypePic
        .TypePicMask = "9999"
        
        .Col = .ColLetterToNumber(fpsDkNgayColXB)
        .Row = fpsDkNgayRow
        .BackColor = vbWhite
        
        .Col = .ColLetterToNumber(fpsDkNgayColXB) + 1
        .Row = fpsDkNgayRow
        .BackColor = mFormColor
        .Lock = True
        
        .Col = 2
        .ColHidden = True
        .Col = 3
        .ColHidden = True
        .Col = 4
        .ColHidden = True
        .Col = 5
        .ColHidden = False
        .Col = 6
        .ColHidden = True
        
        .Col = 8
        .ColHidden = True
        .Col = 9
        .ColHidden = True
        .Col = 10
        .ColHidden = True
        .Col = 11
        .ColHidden = False
        .Col = 12
        .ColHidden = True
        .Col = 14
        .ColHidden = True
        .Col = 15
        .ColHidden = True

    End With

    With fpsLoaiTK
        .BackColor = mFormColor
        'fpSpread4.BorderStyle = BorderStyleNone
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        .Col = .ColLetterToNumber(fpsLoaiTkCol)
        .Row = fpsLoaiTkRow
        .BackColor = vbWhite
    End With

    With fpSKetQua
        .MaxCols = 15
        .EditModePermanent = True
        .EditModeReplace = True
        .CursorType = CursorTypeLockedCell
        .CursorStyle = CursorStyleArrow
        .TypeNumberNegStyle = TypeNumberNegStyle1
        .ColWidth(11) = 0
        .Row = 1
        .RowHeight(1) = 25
        
        .Col = 11
        .ColHidden = True
        .Col = 12
        .ColHidden = True
        
        .Col = 13
        .ColHidden = True
        .Col = 14
        .ColHidden = True
        .Col = 15
        .ColHidden = True

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
            
        .Col = .ColLetterToNumber(fpsDkNgayColF)
        .Row = fpsDkNgayRow
        .Text = strarrdate(2)
        '.Text = strarrdate(0) & "/" & strarrdate(2)
            
        .Col = .ColLetterToNumber(fpsDkNgayColT)
        .Row = fpsDkNgayRow
        .Text = strarrdate(2)
        '.Text = strarrdate(0) & "/" & strarrdate(2)
    End With

    'Lay du lieu cho cbo
    With fpsLoaiTK
        .Col = .ColLetterToNumber(fpsLoaiTkCol)
        .Row = fpsLoaiTkRow
        Dim xmlDocument     As New MSXML.DOMDocument
        Dim xmlNode         As MSXML.IXMLDOMNode
        Dim strDataFileName As String
        Dim i               As Integer

        xmlDocument.Load TAX_Utilities_v1.GetAbsolutePath("Map.xml")
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        ReDim Preserve arrStrId(0)
        arrStrId(0) = "00"

        For Each xmlNode In xmlNodeListMap
            Dim id       As String
            Dim LoaiTk   As String
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
        If fpsDkNgay.ActiveRow = fpsDkNgayRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(fpsDkNgayColT) Then
            fpSKetQua.SetFocus
        End If
    End With
End If
If KeyCode = vbKeyTab And Shift = 1 Then
    If fpsDkNgay.ActiveRow = fpsDkNgayRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(fpsDkNgayColF) Then
        fpsLoaiTK.SetFocus
        With fpsLoaiTK
            .SetActiveCell .ColLetterToNumber(fpsLoaiTkCol), fpsLoaiTkRow
        End With
    End If
    If fpsDkNgay.ActiveRow = fpsDkNgayRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(fpsDkNgayColT) Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(fpsDkNgayColF), fpsDkNgayRow
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

Private Sub fpsDkNgay_LeaveCell(ByVal Col As Long, _
                                ByVal Row As Long, _
                                ByVal NewCol As Long, _
                                ByVal NewRow As Long, _
                                Cancel As Boolean)

    With fpsDkNgay
        Dim strarrdate()                  As String
        Dim strPrefix                     As String
        Dim vdtehientai                   As String
        Dim LoaiTk                        As String
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

        If fpsLoaiTK.TypeComboBoxCurSel <> 0 Then
            LoaiTk = GetAttribute(tkNode, "LoaiTK")
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow

            If Trim(Replace(.Text, "/", "")) <> "" Then

                formatPrefix .Text, strarrdate

                'Bat dk thang
                If (Val(strarrdate(0)) > 12 Or Val(strarrdate(0)) <= 0) And (LoaiTk = KIEU_KY_THANG Or LoaiTk = KIEU_KY_THANG_NAM Or LoaiTk = "KTN") Then
                    .Text = ""
                    DisplayMessage "0090", msOKOnly, miInformation
                    blnOnDKienTraCuu_LeaveCell = False
                    Exit Sub
                End If

                'Bat dk quy
                If (Val(strarrdate(0)) > 4 Or Val(strarrdate(0)) <= 0) And LoaiTk = KIEU_KY_QUY Then
                    .Text = ""
                    DisplayMessage "0091", msOKOnly, miInformation
                    blnOnDKienTraCuu_LeaveCell = False
                    Exit Sub
                End If

                'Bat dk ngay
                If LoaiTk = KIEU_KY_NGAY_PS Or LoaiTk = KIEU_KY_NGAY_NAM Or LoaiTk = "DT" Or LoaiTk = "CD" Or LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then
                    If (Val(strarrdate(1)) > 12 Or Val(strarrdate(1)) <= 0) Then
                        .Text = ""
                        DisplayMessage "0091", msOKOnly, miInformation
                        blnOnDKienTraCuu_LeaveCell = False
                        Exit Sub
                    ElseIf (Val(strarrdate(0)) > Day(GetNgayCuoiThang(Val(strarrdate(2)), Val(strarrdate(1)))) Or Val(strarrdate(0)) <= 0) Then
                        .Text = ""
                        DisplayMessage "0071", msOKOnly, miInformation
                        blnOnDKienTraCuu_LeaveCell = False
                        Exit Sub
                    End If

                End If
                
                'Bat dk ky
                If (Val(strarrdate(0)) > 2 Or Val(strarrdate(0)) <= 0) And LoaiTk = "K" Then
                    .Text = ""
                    DisplayMessage "0311", msOKOnly, miInformation
                    blnOnDKienTraCuu_LeaveCell = False
                    Exit Sub
                End If

                'bat dk nam
                If LoaiTk = KIEU_KY_NAM Then

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

                'Hien thi lai kq
                .Col = .ColLetterToNumber(fpsDkNgayColF)
                .Row = fpsDkNgayRow

                If LoaiTk = KIEU_KY_NGAY_PS Or LoaiTk = KIEU_KY_NGAY_NAM Or LoaiTk = "DT" Or LoaiTk = "CD" Or LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then
                    .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
                ElseIf LoaiTk = KIEU_KY_THANG Or LoaiTk = KIEU_KY_THANG_NAM Or LoaiTk = "KTN" Or LoaiTk = KIEU_KY_QUY Or LoaiTk = "K" Then
                    .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1)
                Else
                    .SetText .Col, .Row, strarrdate(0)
                End If
            End If

            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow

            If Trim(Replace(.Text, "/", "")) <> "" Then

                formatPrefix .Text, strarrdate

                'Bat dk thang
                If (Val(strarrdate(0)) > 12 Or Val(strarrdate(0)) <= 0) And (LoaiTk = KIEU_KY_THANG Or LoaiTk = KIEU_KY_THANG_NAM Or LoaiTk = "KTN") Then
                    .Text = ""
                    DisplayMessage "0090", msOKOnly, miInformation
                    blnOnDKienTraCuu_LeaveCell = False
                    Exit Sub
                End If

                'Bat dk quy
                If (Val(strarrdate(0)) > 4 Or Val(strarrdate(0)) <= 0) And LoaiTk = KIEU_KY_QUY Then
                    .Text = ""
                    DisplayMessage "0091", msOKOnly, miInformation
                    blnOnDKienTraCuu_LeaveCell = False
                    Exit Sub
                End If
                
                'Bat dk ky
                If (Val(strarrdate(0)) > 2 Or Val(strarrdate(0)) <= 0) And LoaiTk = "K" Then
                    .Text = ""
                    DisplayMessage "0311", msOKOnly, miInformation
                    blnOnDKienTraCuu_LeaveCell = False
                    Exit Sub
                End If

                'Bat dk ngay
                If LoaiTk = KIEU_KY_NGAY_PS Or LoaiTk = KIEU_KY_NGAY_NAM Or LoaiTk = "DT" Or LoaiTk = "CD" Or LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then
                    If (Val(strarrdate(1)) > 12 Or Val(strarrdate(1)) <= 0) Then
                        .Text = ""
                        DisplayMessage "0091", msOKOnly, miInformation
                        blnOnDKienTraCuu_LeaveCell = False
                        Exit Sub
                    ElseIf (Val(strarrdate(0)) > Day(GetNgayCuoiThang(Val(strarrdate(2)), Val(strarrdate(1)))) Or Val(strarrdate(0)) <= 0) Then
                        .Text = ""
                        DisplayMessage "0071", msOKOnly, miInformation
                        blnOnDKienTraCuu_LeaveCell = False
                        Exit Sub
                    End If

                End If

                'bat dk nam
                If LoaiTk = KIEU_KY_NAM Then

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

                'Hien thi lai kq
                .Col = .ColLetterToNumber(fpsDkNgayColT)
                .Row = fpsDkNgayRow

                If LoaiTk = KIEU_KY_NGAY_PS Or LoaiTk = KIEU_KY_NGAY_NAM Or LoaiTk = "DT" Or LoaiTk = "CD" Or LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then
                    .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
                ElseIf LoaiTk = KIEU_KY_THANG Or LoaiTk = KIEU_KY_THANG_NAM Or LoaiTk = "KTN" Or LoaiTk = KIEU_KY_QUY Or LoaiTk = "K" Then
                    .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1)
                Else
                    .SetText .Col, .Row, strarrdate(0)
                End If
            End If

        Else
            .Row = fpsDkNgayRow

            .Col = .ColLetterToNumber(fpsDkNgayColF)

            Select Case Len(Trim(.Text))

                Case 1
                    .Text = Year(vdtehientai)

                Case 2
                    .Text = "20" & Trim(.Text)

                Case 3
                    .Text = "2" & Trim(.Text)

                Case 4

                    If Val(.Text) < 2000 Then
                        .Text = "2000"
                    End If

            End Select
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)

            Select Case Len(Trim(.Text))

                Case 1
                    .Text = Year(vdtehientai)

                Case 2
                    .Text = "20" & Trim(.Text)

                Case 3
                    .Text = "2" & Trim(.Text)

                Case 4

                    If Val(.Text) < 2000 Then
                        .Text = "2000"
                    End If

            End Select

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
            .SetActiveCell .ColLetterToNumber(fpsDkNgayColT), fpsDkNgayRow
        End With
    End If
    If KeyCode = vbKeyDown And lngRowFocus < fpSKetQua.MaxRows Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus + 1)
    ElseIf KeyCode = vbKeyUp And lngRowFocus > 2 Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus - 1)
    End If
    
End Sub

Private Sub fpsLoaiTK_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim xmlDocument     As New MSXML.DOMDocument
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim tkID            As String
    strTkGTGT = ""

    With fpsLoaiTK
        .Col = .ColLetterToNumber(fpsLoaiTkCol)
        .Row = fpsLoaiTkRow

        If .TypeComboBoxCurSel = 0 Then

            With fpsDkNgay
                Dim vdtehientai As String
                vdtehientai = format(Date, "dd/mm/yyyy")
                Dim strarrdate() As String

                formatPrefix vdtehientai, strarrdate
            
                .Col = .ColLetterToNumber(fpsDkNgayColF)
                .Row = fpsDkNgayRow
                .CellType = CellTypePic
                .TypePicMask = "9999"
                '.Text = strarrdate(0) & "/" & strarrdate(2)
                .Text = strarrdate(2)
                
                .Col = .ColLetterToNumber(fpsDkNgayColT)
                .Row = fpsDkNgayRow
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
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 4
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 5
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 6
                fpsDkNgay.ColHidden = True
                
                fpsDkNgay.Col = 8
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 9
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 10
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 11
                fpsDkNgay.ColHidden = False
                fpsDkNgay.Col = 12
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 14
                fpsDkNgay.ColHidden = True
                fpsDkNgay.Col = 15
                fpsDkNgay.ColHidden = True
            End With

            Exit Sub
        End If

        xmlDocument.Load TAX_Utilities_v1.GetAbsolutePath("map.xml")
        tkID = arrStrId(.TypeComboBoxCurSel)

        For Each xmlNode In xmlDocument.getElementsByTagName("Map")

            If GetAttribute(xmlNode, "ID") = tkID Then
                Set tkNode = xmlNode.CloneNode(True)
                Exit For
            End If
            
        Next
        
    End With
   
    CreateDkKy
    
End Sub

Private Sub fpsLoaiTK_GotFocus()
    btnTracuu.Default = True
End Sub

Private Sub fpsLoaiTK_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 0 Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(fpsDkNgayColF), fpsDkNgayRow
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
        
        .Col = 2
        .ColHidden = True
        .Col = 3
        .ColHidden = True
        .Col = 4
        .ColHidden = True
        .Col = 5
        .ColHidden = True
        .Col = 6
        .ColHidden = True
                .Col = 7
        .ColHidden = False
        .Col = 8
        .ColHidden = True
        .Col = 9
        .ColHidden = True
        .Col = 10
        .ColHidden = True
        .Col = 11
        .ColHidden = True
        .Col = 12
        .ColHidden = True
                .Col = 13
        .ColHidden = False
        .Col = 14
        .ColHidden = True
        .Col = 15
        .ColHidden = True
        
        Dim LoaiTk As String
        LoaiTk = GetAttribute(tkNode, "LoaiTK")
        .Row = fpsDkNgayRow

        If LoaiTk = KIEU_KY_THANG Or LoaiTk = KIEU_KY_THANG_NAM Then
        
            lstryear = ""
            lstrMonth = "1"
            lstrThreemonths = ""
            lstrDay = ""
            .Col = 3
            .ColHidden = False
            .Col = 9
            .ColHidden = False
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strarrdate(1) & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strarrdate(1) & "/" & strarrdate(2)
        ElseIf LoaiTk = KIEU_KY_QUY Then
            lstryear = ""
            lstrMonth = ""
            lstrThreemonths = "1"
            lstrDay = ""
            .Col = 4
            .ColHidden = False
            .Col = 10
            .ColHidden = False
            
            Dim strQuy As String

            If Val(strarrdate(1)) < 4 Then
                strQuy = "01"
            ElseIf Val(strarrdate(1)) >= 4 And Val(strarrdate(1)) < 7 Then
                strQuy = "02"
            ElseIf Val(strarrdate(1)) >= 7 And Val(strarrdate(1)) < 10 Then
                strQuy = "03"
            Else
                strQuy = "04"
            End If

            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strQuy & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strQuy & "/" & strarrdate(2)
        ElseIf LoaiTk = KIEU_KY_NAM Then
            lstryear = "1"
            lstrMonth = ""
            lstrThreemonths = ""
            lstrDay = ""
            .Col = 5
            .ColHidden = False
            .Col = 11
            .ColHidden = False
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "9999"
            .Text = strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "9999"
            .Text = strarrdate(2)
        ElseIf LoaiTk = KIEU_KY_NGAY_PS Or LoaiTk = KIEU_KY_NGAY_NAM Or LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then
            lstryear = "1"
            lstrMonth = "1"
            lstrThreemonths = ""
            lstrDay = "1"
            .Col = 2
            .ColHidden = False
            .Col = 8
            .ColHidden = False
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//99//9999"
            .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//99//9999"
            .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
        ElseIf LoaiTk = "DT" Or LoaiTk = "CD" Then
            lstryear = "1"
            lstrMonth = "1"
            lstrThreemonths = ""
            lstrDay = "1"
            .Col = 2
            .ColHidden = False
            .Col = 8
            .ColHidden = False
            .Col = 14
            .ColHidden = False
            .Col = 15
            .ColHidden = False
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//99//9999"
            .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//99//9999"
            .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
        ElseIf LoaiTk = "KTN" Then
            .Col = 3
            .ColHidden = False
            .Col = 9
            .ColHidden = False
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strarrdate(1) & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strarrdate(1) & "/" & strarrdate(2)
        ElseIf LoaiTk = "K" Then
            .Col = 6
            .ColHidden = False
            .Col = 12
            .ColHidden = False
            
            If Val(strarrdate(1)) <= 6 Then
                strQuy = "01"
            Else
                strQuy = "02"
            End If
            
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strQuy & "/" & strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strQuy & "/" & strarrdate(2)
        ElseIf LoaiTk = "N" Then
            .Col = 7
            .ColHidden = True
            .Col = 13
            .ColHidden = True
       ElseIf LoaiTk = "KTN_Y" Or LoaiTk = "CD_Y" Or LoaiTk = "DT_Y" Then
            lstryear = "1"
            lstrMonth = ""
            lstrThreemonths = ""
            lstrDay = ""
            .Col = 5
            .ColHidden = False
            .Col = 11
            .ColHidden = False
            .Col = .ColLetterToNumber(fpsDkNgayColF)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "9999"
            .Text = strarrdate(2)
            
            .Col = .ColLetterToNumber(fpsDkNgayColT)
            .Row = fpsDkNgayRow
            .CellType = CellTypePic
            .TypePicMask = "9999"
            .Text = strarrdate(2)
        End If
    
    End With

End Sub

Private Function KiemTraDKngay() As Boolean
    Dim strarrdate() As String
    Dim LoaiTk       As Variant
    Dim strF         As String
    Dim strT         As String
    
    KiemTraDKngay = True

    With fpsDkNgay
        .Row = fpsDkNgayRow
        .Col = .ColLetterToNumber(fpsDkNgayColF)
        strF = .Text
        .Col = .ColLetterToNumber(fpsDkNgayColT)
        strT = .Text

        If fpsLoaiTK.TypeComboBoxCurSel <> 0 Then
            LoaiTk = GetAttribute(tkNode, "LoaiTK")
            
            If LoaiTk = KIEU_KY_NGAY_PS Or LoaiTk = KIEU_KY_NGAY_NAM Or LoaiTk = "DT" Or LoaiTk = "CD" Then

                If Trim$(Replace$(strF, "/", "")) = "" Then
                    strF = "01/01/1900"
                End If

                If Trim$(Replace$(strT, "/", "")) = "" Then
                    strT = "01/01/9900"
                End If
                    
                If ToDate(Replace$(strF, "/", "")) > ToDate(Replace$(strT, "/", "")) Then
                    DisplayMessage "0309", msOKOnly, miCriticalError
                    .SetFocus
                    .SetActiveCell .Col, .Row
                    KiemTraDKngay = False
                End If

            ElseIf LoaiTk = KIEU_KY_THANG Or LoaiTk = KIEU_KY_THANG_NAM Or LoaiTk = "KTN" Then

                If Trim$(Replace$(strF, "/", "")) = "" Then
                    strF = "01/1900"
                End If

                If Trim$(Replace$(strT, "/", "")) = "" Then
                    strT = "01/9900"
                End If
                    
                If ToDate("01" & Replace$(strF, "/", "")) > ToDate("01" & Replace$(strT, "/", "")) Then
                    DisplayMessage "0097", msOKOnly, miCriticalError
                    .SetFocus
                    .SetActiveCell .Col, .Row
                    KiemTraDKngay = False
                End If

            ElseIf LoaiTk = KIEU_KY_QUY Then

                If Trim$(Replace$(strF, "/", "")) = "" Then
                    strF = "01/1900"
                End If

                If Trim$(Replace$(strT, "/", "")) = "" Then
                    strT = "01/9900"
                End If
                    
                If ToDate("01" & Replace$(strF, "/", "")) > ToDate("01" & Replace$(strT, "/", "")) Then
                    DisplayMessage "0098", msOKOnly, miCriticalError
                    .SetFocus
                    .SetActiveCell .Col, .Row
                    KiemTraDKngay = False
                End If

            ElseIf LoaiTk = "K" Then

                If Trim$(Replace$(strF, "/", "")) = "" Then
                    strF = "01/1900"
                End If

                If Trim$(Replace$(strT, "/", "")) = "" Then
                    strT = "01/9900"
                End If
                    
                If ToDate("01" & Replace$(strF, "/", "")) > ToDate("01" & Replace$(strT, "/", "")) Then
                    DisplayMessage "0312", msOKOnly, miCriticalError
                    .SetFocus
                    .SetActiveCell .Col, .Row
                    KiemTraDKngay = False
                End If

            Else

                If Trim$(Replace$(strF, "/", "")) = "" Then
                    strF = "1900"
                End If

                If Trim$(Replace$(strT, "/", "")) = "" Then
                    strT = "9900"
                End If

                If Val(strF) > Val(strT) Then
                    DisplayMessage "0093", msOKOnly, miCriticalError
                    .SetFocus
                    .SetActiveCell .Col, .Row
                    KiemTraDKngay = False
                End If

            End If

        Else

            If Trim$(Replace$(strF, "/", "")) = "" Then
                strF = "1900"
            End If

            If Trim$(Replace$(strT, "/", "")) = "" Then
                strT = "9900"
            End If

            If Val(strF) > Val(strT) Then
                DisplayMessage "0093", msOKOnly, miCriticalError
                .SetFocus
                .SetActiveCell .Col, .Row
                KiemTraDKngay = False
            End If
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
    
    xmlDocument.Load TAX_Utilities_v1.GetAbsolutePath("map.xml")
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
    
    'Khoi tao gia tri khoang tra cuu
    If Len(strPeriodFrom) = 4 Then
        strPeriodFrom = "01/01/" & strPeriodFrom
    Else
        If strId = "80" Or strId = "82" Or strId = "98" Then
        Else
            strPeriodFrom = "01/" & strPeriodFrom
        End If
    End If
    
    If Len(strPeriodTo) = 4 Then
        strPeriodTo = "01/12/" & strPeriodTo
    Else
        If strId = "80" Or strId = "82" Or strId = "98" Then
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
        ' to khai quy GTGT
        If strId = "01" Or strId = "02" Or strId = "04" Or strId = "71" Or strId = "36" Or strId = "94" Or strId = "96" Then
            If strTkGTGT = "TK_QUY" Then
                strKieu_Ky = KIEU_KY_QUY
            Else
                strKieu_Ky = KIEU_KY_THANG
            End If
        ElseIf strId = "98" Then
            If strTkGTGT = "TK_LANXB" Then
                strKieu_Ky = KIEU_KY_NGAY_NAM
            End If
        Else
            strKieu_Ky = KIEU_KY_THANG
        End If
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
                        Dim tmpFileNam As String
                        If strId = "01" Or strId = "02" Or strId = "04" Or strId = "71" Or strId = "36" Or strId = "94" Or strId = "96" Then
                            If strTkGTGT = "TK_QUY" Then
                                tmpFileNam = GetAttribute(xmlNodeMenu, "Caption") & " quý"
                            Else
                                tmpFileNam = GetAttribute(xmlNodeMenu, "Caption")
                            End If
                        ElseIf strId = "98" Then
                            If strTkGTGT = "TK_LANXB" Then
                                tmpFileNam = GetAttribute(xmlNodeMenu, "Caption") & " Lan XB"
                            End If
                        Else
                            tmpFileNam = GetAttribute(xmlNodeMenu, "Caption")
                        End If
                        ReDim Preserve strPeriods(lngIndex2)
                        strPeriods(lngIndex2) = strId & _
                            "~" & tmpFileNam & _
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

Private Sub GetDsToKhai(strFroms As String, _
                        strTos As String, _
                        LanXB As String, _
                        ByRef strPeriods() As String)
    On Error GoTo ErrHandle
    Dim LoaiTk           As String
    Dim DataFileName     As String
    Dim ListDataFile()   As String
    Dim tkThangQuy       As String
    Dim kyKeKhai         As String
    Dim strId            As String
    Dim strThueKhauTruId As String
    Dim strThuePhaiNopId As String
    Dim strPeriodReturn  As String
    Dim tenTk            As String
    Dim tkIndex          As Integer
    Dim fileStr          As Variant
    Dim strFrom          As String
    Dim strTo            As String
    Dim tkSplit()        As String
    Dim tkBoSung         As String
    
    LoaiTk = GetAttribute(tkNode, "LoaiTK")
    DataFileName = GetAttribute(tkNode.firstChild, "DataFile")
    ListDataFile = Split(DataFileName, ",")
    DataFileName = ListDataFile(0)
    tkThangQuy = GetAttribute(tkNode, "TkThangQuy")
    strId = GetAttribute(tkNode, "ID")
    strThueKhauTruId = GetAttribute(tkNode.firstChild, "ThueKhauTru")
    strThuePhaiNopId = GetAttribute(tkNode.firstChild, "ThuePhaiNop")
    
    tenTk = ""
    tkBoSung = ""
    tkIndex = UBound(strPeriods) + 1
    strFrom = strFroms
    strTo = strTos

    'Kiem tra ngay bat dau nam tai chinh
    If GetAttribute(tkNode, "FinanceYear") = "1" Then
        strNgayTaiChinh = GetNgayBatDauNamTaiChinh
        If KiemTraNgayTaiChinh(strNgayTaiChinh, False) Then
            iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
            iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
        End If
    Else
        strNgayTaiChinh = "01/01"
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
    End If

    If LoaiTk = KIEU_KY_THANG Then
        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/" & strFroms
            strTo = "12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If
                
            strPeriodReturn = Mid$(kyKeKhai, 1, 2) & "/" & Mid$(kyKeKhai, 3, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

            If Len(kyKeKhai) = 6 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 Then

                If ToDate("01" & kyKeKhai) >= ToDate("01" & strFrom) And ToDate("01" & kyKeKhai) <= ToDate("01" & strTo) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
               
            End If
                
        Next

    ElseIf LoaiTk = KIEU_KY_QUY Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/" & strFroms
            strTo = "04/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            If tkThangQuy = "1" Then
                kyKeKhai = Replace$(fileStr, DataFileName & "_Q", "")
            Else
                kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            End If

            tenTk = GetAttribute(tkNode, "Caption")
     
            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If

            strPeriodReturn = Mid$(kyKeKhai, 1, 2) & "/" & Mid$(kyKeKhai, 3, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

            If Len(kyKeKhai) = 6 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 Then

                If ToDate("01" & kyKeKhai) >= ToDate("01" & strFrom) And ToDate("01" & kyKeKhai) <= ToDate("01" & strTo) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
               
            End If

        Next

    ElseIf LoaiTk = KIEU_KY_NAM Then
        strFrom = strFroms
        strTo = strTos

        For Each fileStr In arrStrXMLFileNames
            ' xy ly cac to khai QT bo sung them tu thang den thang
            'If arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "87" or   Then
                kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
                'If Len(kyKeKhai) = 18 And Val(Left$(kyKeKhai, 4)) > 100 Then
                If InStr(fileStr, DataFileName) > 0 And InStr(fileStr, "KHBS") <= 0 Then
                    tkSplit = Split(kyKeKhai, "_")
                    tenTk = ""
                    If UBound(tkSplit) = 2 Then
                        tenTk = GetAttribute(tkNode, "Caption") & "(" & Left$(tkSplit(1), 2) & "/" & Right$(tkSplit(1), 4) & "-" & Left$(tkSplit(2), 2) & "/" & Right$(tkSplit(2), 4) & ")"
                    ElseIf UBound(tkSplit) = 3 Then
                        tenTk = GetAttribute(tkNode, "Caption") & "(" & Left$(tkSplit(2), 2) & "/" & Right$(tkSplit(2), 4) & "-" & Left$(tkSplit(3), 2) & "/" & Right$(tkSplit(3), 4) & ")"
                    End If
                    
                    If InStr(kyKeKhai, "bs") > 0 Then
                        tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                        tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                        kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
                        tkSplit = Split(kyKeKhai, "_")
                    End If
                
                    strPeriodReturn = Left$(kyKeKhai, 4) & "~" & Left$(tkSplit(1), 2) & "/" & Right$(tkSplit(1), 4) & "~" & Left$(tkSplit(2), 2) & "/" & Right$(tkSplit(2), 4) & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)
                    If (Val(Left$(kyKeKhai, 4)) >= Val(Right$(strFrom, 4)) And Val(Left$(kyKeKhai, 4)) <= Val(Right$(strTo, 4))) Then
                        ReDim Preserve strPeriods(tkIndex)
                        strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                        tkIndex = tkIndex + 1
                    End If
                   
                End If
'            Else
'                kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
'                tenTk = GetAttribute(tkNode, "Caption")
'
'                If InStr(kyKeKhai, "bs") > 0 Then
'                    tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
'                    tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
'                    kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
'                End If
'
'                strPeriodReturn = kyKeKhai & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)
'
'                If Len(kyKeKhai) = 4 And Val(Right$(kyKeKhai, 4)) > 0 Then
'
'                    If (Val(Right$(kyKeKhai, 4)) >= Val(Right$(strFrom, 4)) And Val(Right$(kyKeKhai, 4)) <= Val(Right$(strTo, 4))) Then
'                        ReDim Preserve strPeriods(tkIndex)
'                        strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
'                        tkIndex = tkIndex + 1
'                    End If
'
'                End If
'             End If
        Next

    ElseIf LoaiTk = KIEU_KY_NGAY_PS Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/01/" & strFroms
            strTo = "31/12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If
    
            strPeriodReturn = Left$(kyKeKhai, 2) & "/" & Mid$(kyKeKhai, 3, 2) & "/" & Right$(kyKeKhai, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

            If Len(kyKeKhai) = 8 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 And Val(Mid$(kyKeKhai, 2, 2)) > 0 Then

                If ToDate(kyKeKhai) >= ToDate(strFrom) And ToDate(kyKeKhai) <= ToDate(strTo) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
               
            End If

        Next

    ElseIf LoaiTk = KIEU_KY_NGAY_NAM Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/01/" & strFroms
            strTo = "31/12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If
    
            tkSplit = Split(kyKeKhai, "_")

            If UBound(tkSplit) = 2 Then
                If Len(tkSplit(0)) = 4 And Len(tkSplit(1)) = 8 And Len(tkSplit(2)) = 8 Then
                    If ToDate(tkSplit(1)) >= ToDate(strFrom) And ToDate(tkSplit(2)) <= ToDate(strTo) Then
                        strPeriodReturn = tkSplit(0) & "~" & Left$(tkSplit(1), 2) & "/" & Mid$(tkSplit(1), 3, 2) & "/" & Right$(tkSplit(1), 4) & "~" & Left$(tkSplit(2), 2) & "/" & Mid$(tkSplit(2), 3, 2) & "/" & Right$(tkSplit(2), 4) & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

                        ReDim Preserve strPeriods(tkIndex)
                        strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                        tkIndex = tkIndex + 1
                    End If
                End If
            End If
            
        Next
    ElseIf LoaiTk = KIEU_KY_TU_NGAY_DEN_NGAY Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/01/" & strFroms
            strTo = "31/12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If
    
            tkSplit = Split(kyKeKhai, "_")
            
            tenTk = tenTk & "(" & Left$(tkSplit(0), 2) & "/" & Mid$(tkSplit(0), 3, 2) & "/" & Right$(tkSplit(0), 4) & "-" & Left$(tkSplit(1), 2) & "/" & Mid$(tkSplit(1), 3, 2) & "/" & Right$(tkSplit(1), 4) & ")"
            
            If UBound(tkSplit) = 1 Then
                If Len(tkSplit(0)) = 8 And Len(tkSplit(1)) = 8 Then
                    If ToDate(tkSplit(0)) >= ToDate(strFrom) And ToDate(tkSplit(1)) <= ToDate(strTo) Then
                        strPeriodReturn = Right$(tkSplit(0), 4) & "~" & Left$(tkSplit(0), 2) & "/" & Mid$(tkSplit(0), 3, 2) & "/" & Right$(tkSplit(0), 4) & "~" & Left$(tkSplit(1), 2) & "/" & Mid$(tkSplit(1), 3, 2) & "/" & Right$(tkSplit(1), 4) & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

                        ReDim Preserve strPeriods(tkIndex)
                        strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                        tkIndex = tkIndex + 1
                    End If
                End If
            End If
            
        Next
    ElseIf LoaiTk = KIEU_KY_THANG_NAM Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/01/" & strFroms
            strTo = "31/12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If

            tkSplit = Split(kyKeKhai, "_")
            
            If UBound(tkSplit) = 1 Then
                If Len(tkSplit(0)) = 6 And Len(tkSplit(1)) = 6 Then

                    If ToDate("01" & tkSplit(0)) >= ToDate("01" & strFrom) And ToDate("01" & tkSplit(1)) <= ToDate("01" & strTo) Then
                        strPeriodReturn = Left$(kyKeKhai, 2) & "/" & Mid$(kyKeKhai, 3, 4) & "~" & Left$(tkSplit(0), 2) & "/" & Right$(tkSplit(0), 4) & "~" & Left$(tkSplit(1), 2) & "/" & Right$(tkSplit(1), 4) & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)
                        ReDim Preserve strPeriods(tkIndex)
                        strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                        tkIndex = tkIndex + 1
                    End If
                End If
            End If
        Next

    ElseIf LoaiTk = "DT" Or LoaiTk = "CD" Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/01/" & strFroms
            strTo = "31/12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames
            tenTk = ""
            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If

            tkSplit = Split(kyKeKhai, "_")

            If UBound(tkSplit) = 1 Then
                If tkSplit(0) = "L" & LanXB Or fpsLoaiTK.TypeComboBoxCurSel = 0 Then
                    kyKeKhai = tkSplit(1)

                    If tkSplit(0) <> "L" Then
                        strPeriodReturn = Left$(kyKeKhai, 2) & "/" & Mid$(kyKeKhai, 3, 2) & "/" & Right$(kyKeKhai, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "_" & tkSplit(0), tkBoSung)
                        tenTk = GetAttribute(tkNode, "Caption") & " Lan XB " & Right$(tkSplit(0), Len(tkSplit(0)) - 1) & tenTk
                    Else
                        tenTk = GetAttribute(tkNode, "Caption") & tenTk
                    End If
                       
                    If Len(kyKeKhai) = 8 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 And Val(Mid$(kyKeKhai, 2, 2)) > 0 Then

                        If ToDate(kyKeKhai) >= ToDate(Replace$(strFrom, "/", "")) And ToDate(kyKeKhai) <= ToDate(strTo) Then
                            ReDim Preserve strPeriods(tkIndex)
                            strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~" & Right$(tkSplit(0), Len(tkSplit(0)) - 1)
                            tkIndex = tkIndex + 1
                        End If
               
                    End If

                End If
                    
            End If
                
        Next

    ElseIf LoaiTk = "KTN" Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/" & strFroms
            strTo = "12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If
                
            strPeriodReturn = Mid$(kyKeKhai, 1, 2) & "/" & Mid$(kyKeKhai, 3, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

            If Len(kyKeKhai) = 6 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 Then

                If ToDate("01" & kyKeKhai) >= ToDate("01" & strFrom) And ToDate("01" & kyKeKhai) <= ToDate("01" & strTo) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
               
            End If
                
        Next
    ElseIf LoaiTk = "DT_Y" Or LoaiTk = "CD_Y" Then
        strFrom = strFroms
        strTo = strTos

        For Each fileStr In arrStrXMLFileNames
            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            If InStr(fileStr, DataFileName) > 0 And InStr(fileStr, "KHBS") <= 0 Then
                tkSplit = Split(kyKeKhai, "_")
                tenTk = ""
                If UBound(tkSplit) = 2 Then
                    tenTk = GetAttribute(tkNode, "Caption") & "(" & Left$(tkSplit(1), 2) & "/" & Right$(tkSplit(1), 4) & "-" & Left$(tkSplit(2), 2) & "/" & Right$(tkSplit(2), 4) & ")"
                ElseIf UBound(tkSplit) = 3 Then
                    tenTk = GetAttribute(tkNode, "Caption") & "(" & Left$(tkSplit(2), 2) & "/" & Right$(tkSplit(2), 4) & "-" & Left$(tkSplit(3), 2) & "/" & Right$(tkSplit(3), 4) & ")"
                End If
                
                If InStr(kyKeKhai, "bs") > 0 Then
                    tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                    tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                    kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
                    tkSplit = Split(kyKeKhai, "_")
                End If
                
                 strPeriodReturn = Left$(kyKeKhai, 4) & "~" & Left$(tkSplit(1), 2) & "/" & Right$(tkSplit(1), 4) & "~" & Left$(tkSplit(2), 2) & "/" & Right$(tkSplit(2), 4) & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)
                If (Val(Left$(kyKeKhai, 4)) >= Val(Right$(strFrom, 4)) And Val(Left$(kyKeKhai, 4)) <= Val(Right$(strTo, 4))) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
            End If
        Next
    ElseIf LoaiTk = "KTN_Y" Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/" & strFroms
            strTo = "12/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")
            tenTk = GetAttribute(tkNode, "Caption")

            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If
                
            strPeriodReturn = Mid$(kyKeKhai, 1, 2) & "/" & Mid$(kyKeKhai, 3, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

            If Len(kyKeKhai) = 6 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 Then

                If ToDate("01" & kyKeKhai) >= ToDate("01" & strFrom) And ToDate("01" & kyKeKhai) <= ToDate("01" & strTo) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
               
            End If
                
        Next
    ElseIf LoaiTk = "K" Then

        If Len(strFroms) = 4 And Len(strTos) = 4 Then
            strFrom = "01/" & strFroms
            strTo = "02/" & strTos
        End If

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")

            tenTk = GetAttribute(tkNode, "Caption")
     
            If InStr(kyKeKhai, "bs") > 0 Then
                tenTk = tenTk & " BS lan " & Mid$(kyKeKhai, 3, InStr(kyKeKhai, "_") - 3)
                tkBoSung = Left$(kyKeKhai, InStr(kyKeKhai, "_"))
                kyKeKhai = Right$(kyKeKhai, Len(kyKeKhai) - InStr(kyKeKhai, "_"))
            End If

            strPeriodReturn = Mid$(kyKeKhai, 1, 2) & "/" & Mid$(kyKeKhai, 3, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, kyKeKhai, "", tkBoSung)

            If Len(kyKeKhai) = 6 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 Then

                If ToDate("01" & kyKeKhai) >= ToDate("01" & strFrom) And ToDate("01" & kyKeKhai) <= ToDate("01" & strTo) Then
                    ReDim Preserve strPeriods(tkIndex)
                    strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                    tkIndex = tkIndex + 1
                End If
               
            End If

        Next

    ElseIf LoaiTk = "N" Then
        Dim CurrentKyKeKhai As String
        CurrentKyKeKhai = "011900"

        For Each fileStr In arrStrXMLFileNames

            kyKeKhai = Replace$(fileStr, DataFileName & "_", "")

            If Len(kyKeKhai) = 6 And Val(Left$(kyKeKhai, 2)) > 0 And Val(Right$(kyKeKhai, 4)) > 0 Then
                If ToDate("01" & kyKeKhai) > ToDate("01" & CurrentKyKeKhai) Then
                    CurrentKyKeKhai = kyKeKhai
                End If
            End If

        Next

        If CurrentKyKeKhai <> "011900" Then
            tenTk = GetAttribute(tkNode, "Caption")
     
            strPeriodReturn = Mid$(CurrentKyKeKhai, 1, 2) & "/" & Mid$(CurrentKyKeKhai, 3, 4) & "~" & "~" & "~True~~" & GetDataFileNames(ListDataFile, CurrentKyKeKhai, "", "")

                ReDim Preserve strPeriods(tkIndex)
                strPeriods(tkIndex) = strId & "~" & tenTk & "~" & GetAttribute(tkNode.firstChild, "StartDate") & "~" & strPeriodReturn & "~" & GetTaxValue(fileStr, strThueKhauTruId, False) & "~" & GetTaxValue(fileStr, strThuePhaiNopId, False) & "~" & LoaiTk & "~" & tkThangQuy & "~"
                tkIndex = tkIndex + 1
        End If

    End If

    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "GetTaxReportsById", Err.Number, Err.Description
End Sub

Private Sub LoadXMLFileNames()
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As file
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v1.DataFolder)).Files
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
    ElseIf arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "981" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "982" Then
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
                        strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
                        IsValidPeriod = True
                        Exit Function
                    End If
                End If
            End If
        Case KIEU_KY_QUY
            If strTkGTGT = "TK_QUY" Then
                Dim tmp1 As String
                
                lStrPeriod = Right$(strFileName, 6)
                tmp1 = Right$(strFileName, 7)
                tmp1 = Left$(tmp1, 1)
                If tmp1 <> "Q" Then
                    IsValidPeriod = False
                    Exit Function
                End If
            
                If strDataFiles(0) <> Mid(strFileName, 1, Len(strFileName) - Len(lStrPeriod) - 2) Then
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
                            If strTkGTGT = "TK_QUY" And Len(strGtgtIdTmp) = 3 Then
                                strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, "Q" & lStrPeriod, "", "")
                            Else
                                strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
                            End If
                            IsValidPeriod = True
                            Exit Function
                        End If
                    End If
                End If
            Else
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
                            strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
                            IsValidPeriod = True
                            Exit Function
                        End If
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
                    strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
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
                        strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
                        IsValidPeriod = True
                        Exit Function
                    End If
                End If
                
            ElseIf arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "981" Or arrStrId(fpsLoaiTK.TypeComboBoxCurSel) = "982" Then
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
                        strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
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
                            strPeriod = strPeriod & "~" & GetDataFileNames(strDataFiles, lStrPeriod, "", "")
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
    
    xmlDom.Load TAX_Utilities_v1.DataFolder & strDataFileName & ".xml"
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


Private Function GetDataFileNames(arrStrDataFiles() As String, strPeriod As String, LanXB As String, lanBS As String) As String
    Dim intCtrl As Integer
    Dim strReturn As String
    
    For intCtrl = 0 To UBound(arrStrDataFiles)
        If intCtrl = 0 Then
            strReturn = lanBS & arrStrDataFiles(intCtrl) & LanXB & "_" & strPeriod
        Else
            strReturn = strReturn & "," & lanBS & arrStrDataFiles(intCtrl) & LanXB & "_" & strPeriod
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
        If fso.FileExists(GetAbsolutePath(TAX_Utilities_v1.DataFolder & arrStrDataFiles(intCtrl) & ".xml")) Then
            fso.DeleteFile GetAbsolutePath(TAX_Utilities_v1.DataFolder & arrStrDataFiles(intCtrl) & ".xml"), True
            '.DeleteRows i + 1, 1
        End If
    Next intCtrl
End Sub

Private Sub TraCuu()
    Dim lRow            As Long
    Dim arrStrPeriods() As String
    Dim strFromDay      As String, strToDay As String
    Dim arrStrTemp()    As String
    Dim lCtrl           As Long
    Dim LanXB           As String
    
    fpsDkNgay_LeaveCell 1, 1, 2, 1, False

    If blnDKienTraCuu = False Then
        fpsDkNgay.SetFocus
        Exit Sub
    End If
    
    'Kiem tra dk tu ky phai nho hon den ky
    If Not KiemTraDKngay Then Exit Sub
    
    'Kiem tra, khoi tao cac gia tri dieu kien tra cuu
    With fpsDkNgay
        .Col = .ColLetterToNumber(fpsDkNgayColF)
        .Row = fpsDkNgayRow
        strFromDay = .Text
        .Col = .ColLetterToNumber(fpsDkNgayColT)
        .Row = fpsDkNgayRow
        strToDay = .Text
        
        .Col = .ColLetterToNumber(fpsDkNgayColXB)
        .Row = fpsDkNgayRow
        LanXB = .Text
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
        GetDsToKhai strFromDay, strToDay, LanXB, arrStrPeriods
    Else
        Dim xmlDocument As New MSXML.DOMDocument
        xmlDocument.Load TAX_Utilities_v1.GetAbsolutePath("map.xml")
       
        For Each tkNode In xmlDocument.getElementsByTagName("Map")

            GetDsToKhai strFromDay, strToDay, LanXB, arrStrPeriods
            
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
            .SetText 13, lRow, arrStrTemp(11)   'LoaiTk
            .SetText 14, lRow, arrStrTemp(12)   'tk thang/quy
            .SetText 15, lRow, arrStrTemp(13)   'Lan xuat ban
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
    xmlDocument.Load TAX_Utilities_v1.GetAbsolutePath("map.xml")
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

Private Function getKHBSDate(DataFile As String) As String
    Dim xmlDoc      As New MSXML.DOMDocument
    Dim xmlNode     As MSXML.IXMLDOMNode
    Dim DataFiles() As String
    Dim fso         As New FileSystemObject
    Dim KHBSfile    As Variant
    DataFiles = Split(DataFile, ",")

    For Each KHBSfile In DataFiles

        If InStr(KHBSfile, "KHBS") > 0 Then
            If fso.FileExists(GetAbsolutePath(TAX_Utilities_v1.DataFolder & KHBSfile & ".xml")) Then
                xmlDoc.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & KHBSfile & ".xml")
                Set xmlNode = xmlDoc.nodeFromID("B_47")
                If Not xmlNode Is Nothing Then
                    getKHBSDate = GetAttribute(xmlNode, "Value")
                    Exit Function
                End If
            End If
        End If

    Next
    getKHBSDate = ""
End Function

Private Function getCellValue(DataFile As String, cellID As String) As String
    Dim xmlDoc As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim fso    As New FileSystemObject
    On Error GoTo ErrHandle

    If fso.FileExists(GetAbsolutePath(TAX_Utilities_v1.DataFolder & DataFile & ".xml")) Then
        xmlDoc.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & DataFile & ".xml")
        Set xmlNode = xmlDoc.nodeFromID(cellID)

        If Not xmlNode Is Nothing Then
            getCellValue = GetAttribute(xmlNode, "Value")
            Exit Function
        End If
    End If

    getCellValue = ""
    
ErrHandle:
    getCellValue = ""
    SaveErrorLog "frmTraCuu", "getCellValue", Err.Number, Err.Description
End Function

Sub ResetAllProperty()
    TAX_Utilities_v1.month = ""
    TAX_Utilities_v1.ThreeMonths = ""
    TAX_Utilities_v1.Year = ""
    TAX_Utilities_v1.DateKHBS = ""
    TAX_Utilities_v1.Day = ""
    TAX_Utilities_v1.FirstDay = ""
    TAX_Utilities_v1.LastDay = ""
    strQuy = ""
    strKHBS = ""
    strKieuKy = ""
    strLoaiNNKD = ""
    strLoaiTkDk = ""
    strLoaiTKQT = ""
    strLoaiTKThang_PS = ""
    strTkDK = ""
    strTkGTGT = ""
    ngayLapTkBs = ""
End Sub
