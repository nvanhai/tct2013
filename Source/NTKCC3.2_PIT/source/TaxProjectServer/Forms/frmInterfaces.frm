VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmInterfaces 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7905
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   11535
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmInterfaces"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6750
      Left            =   0
      TabIndex        =   4
      Top             =   270
      Width           =   11535
      Begin MSCommLib.MSComm MSComm1 
         Left            =   1050
         Top             =   1770
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   6600
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   11475
         _Version        =   458752
         _ExtentX        =   20241
         _ExtentY        =   11642
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterfaces.frx":0000
      End
      Begin MSForms.Label lblExit 
         Height          =   945
         Left            =   2760
         TabIndex        =   13
         Top             =   2820
         Visible         =   0   'False
         Width           =   7095
         Size            =   "12515;1667"
         FontName        =   "Tahoma"
         FontHeight      =   255
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblConnecting 
         Height          =   945
         Left            =   3480
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   7095
         Size            =   "12515;1667"
         FontName        =   "Tahoma"
         FontHeight      =   255
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblLoading 
         Height          =   945
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   7095
         Size            =   "12515;1667"
         FontName        =   "Tahoma"
         FontHeight      =   255
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   6990
      Width           =   11535
      Begin VB.CommandButton cmd_insert 
         Caption         =   "Ghi QHS"
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   390
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdViewNow 
         Height          =   375
         Left            =   6060
         TabIndex        =   18
         Top             =   360
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Xem TK"
         Size            =   "2302;661"
         Accelerator     =   86
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblWarning 
         Height          =   255
         Left            =   9000
         TabIndex        =   16
         Top             =   150
         Visible         =   0   'False
         Width           =   2415
         ForeColor       =   255
         VariousPropertyBits=   8388627
         Size            =   "4260;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVersion 
         Height          =   255
         Left            =   8670
         TabIndex        =   15
         Top             =   150
         Width           =   435
         VariousPropertyBits=   8388627
         Size            =   "767;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblLabelVersion 
         Height          =   255
         Left            =   3630
         TabIndex        =   14
         Top             =   150
         Width           =   4875
         VariousPropertyBits=   8388627
         Size            =   "8599;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblBarcode 
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
         VariousPropertyBits=   8388627
         Size            =   "2143;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFile 
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
         VariousPropertyBits=   8388627
         Size            =   "2143;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFilePath 
         Height          =   255
         Left            =   1530
         TabIndex        =   10
         Top             =   150
         Width           =   1785
         VariousPropertyBits=   8388627
         Size            =   "3149;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   375
         Left            =   7440
         TabIndex        =   2
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Clear"
         Size            =   "2293;661"
         Accelerator     =   88
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   8760
         TabIndex        =   0
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Save"
         Size            =   "2293;661"
         Accelerator     =   71
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdExit 
         Height          =   375
         Left            =   10140
         TabIndex        =   1
         Top             =   390
         Width           =   1305
         Caption         =   "Exit"
         Size            =   "2293;661"
         Accelerator     =   84
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSForms.Label lblCaption 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3975
      ForeColor       =   -2147483634
      Size            =   "7011;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mCommPort = 1
Const mBaudRate = 9600
Const mParity = "N"
Const mDataBits = 8
Const mStopBits = 1
Const mHandshaking = 1
Const tt156 = "01~02~04~71~11~12~15~16~50~51~36~74~75~06~90~23~25"
Const tt156_tkbs = "01~02~04~71~72~11~12~73~15~16~50~51~36~74~75~70~81~06~05~90~23~25"

Private xmlDocumentInit()       As MSXML.DOMDocument
Private arrStrElements()        As String               ' array of barcode string or file name string
Private mHeaderSheet            As Integer              ' Save value of Header sheet (last sheet)
Private blnReceiveByBarcode     As Boolean                    ' Check whether form is loaded
Private objTaxBusiness          As Object               ' private business object (cls001, cls002, cls003, ...)
Private strTaxReportInfo        As String                  ' Info about current tax report

Private mOnLoad                 As Boolean
Private blnOnLoadEvent          As Boolean
'Private strMaSoTep              As String
Private strNgayNhanToKhai       As String
Private strMaPhongXuLy          As String
Private blnSaveSuccess          As Boolean
Private rsPXL                   As ADODB.Recordset      ' Luu danh sach cac phong ban
Private strTaxReportVersion     As String
Private arrBCBuffer() As String
Private arrBCNew() As String
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Dim menuId As Integer

Public KYKK_TU_NGAY As String
Public KYKK_TU_NGAY_F As String  ' Dung cho ham check thanh tra kiem tra (dhdang)
Public KYKK_DEN_NGAY As String
' dhdang in BB nop cham
Public MST_PRINT As String
Public MATK_PRINT As String
Public NNT_PRINT As String
Public LOAihs_PRINT As String
Public DIACHI_PRINT As String
Public NGAYNOP_PRINT As Date
Public NGNOP_S As String
Public HAN_NOP As Date
Public KyKeKhai As String
Public CAN_CU1 As String
Public CAN_CU2 As String
' End
Public SO_TEP As Variant
Public DHS_MA As String
Dim USER As Variant

Private checkSoCT As Integer ' check so chi tieu tren to khai
Private isSheetTk As Boolean ' kiem tra sheet la to khai hay phu luc


Private strMaPhongQuanLy As String
Private strTenPhongQuanLy As String
Private isTonTaiAC As Boolean

'****************************
'Description: StartBarcodeReader procedure start barcode listener on com 1 port
'Author:TuyenDS
'Date:
'Input:
'OutPut:
'Return:
'****************************
Private Sub StartBarcodeReader()
    Dim strSetting As String
On Error GoTo ErrHandle
    strSetting = mBaudRate & "," & mParity & "," & mDataBits & "," & mStopBits
    With MSComm1
        If .PortOpen = False Then
            .Handshaking = mHandshaking
            .CommPort = mCommPort
            .Settings = strSetting                          ' 9600 baud, no parity, 8 data, and 1 stop bit.
            .InputLen = 0                                   ' Read entire buffer
            .RThreshold = 1                                 ' Call **_OnComm for each character
            .InputMode = comInputModeBinary
            On Error GoTo PortOpenedErr ' Port in use
            .PortOpen = True      ' Opens the port
        End If
    End With
    Exit Sub
PortOpenedErr:
    DisplayMessage "0061", msOKOnly, miCriticalError
    Unload Me
    frmTreeviewMenu.Show
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "StartBarcodeReader", Err.Number, Err.Description
    Err.Raise Err.Number
End Sub

'****************************
'Description: StopBarcodeReader procedure stop barcode listener on com 1 port
'Author:TuyenDS
'Date:
'Input:
'OutPut:
'Return:
'****************************
Private Sub StopBarcodeReader()
On Error GoTo ErrHandle
    With MSComm1
        If .PortOpen = True Then .PortOpen = False
    End With
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "StopBarcodeReader", Err.Number, Err.Description
End Sub
'dhdang insert vao CSDL QLT
Private Sub cmd_insert_Click()
On Error GoTo ErrHandle

    Dim strSQL As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String, strSQL_KHBS As String
    Dim HdrID As Variant, strDate() As String, dDate As Date
    Dim rs As ADODB.Recordset, i As Long
    Dim blHoiTonTai As Integer
    Dim blUpdateTHUETKY2 As Boolean
    Dim bln As Boolean
    Dim blnKTRB As Integer
    Dim sSaiCT11 As String
    Dim vKYLBO As Variant
    Dim vNGAYQUET As Variant
    Dim vNGAY_DAU_KYLBO As Variant
    Dim sSQL As String
    'Dim menuId As Integer
    Dim NGAY_HIENTAI As Date
    Dim s As String
    Dim TEP_ID As String
    NGAY_HIENTAI = GetNgayNhap
    'Set rs = New ADODB.Recordset
    sSaiCT11 = ""
    '***************************
    'ThanhDX added
    'Date:23/11/2005
    If TAX_Utilities_Svr_New.Data(0) Is Nothing Then Exit Sub
    '***************************
       
    blnSaveSuccess = False
    
    CallFinish
    ' Kiem tra xem da khoa so trong ky lap bo nay chua
    ' hlnam edit
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionStringSQL spathQHSCC
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
     End If
     
     With fpSpread1
        .Sheet = 1
        If menuId = 5 Then
            .GetText .ColLetterToNumber("I"), 10, vKYLBO
        ElseIf menuId = 6 Or menuId = 8 Or menuId = 9 Then
            .GetText .ColLetterToNumber("I"), 9, vKYLBO
        Else
            .GetText .ColLetterToNumber("E"), 10, vKYLBO
        End If
        
        vNGAY_DAU_KYLBO = "01/" & IIf(Len(Trim(vKYLBO)) = 6, "0" & vKYLBO, vKYLBO) ' Lay ngay dau cua ky lap bo de xem ngay quet co phu hop voi ky khoa so hay khong?
        
        If Trim(vKYLBO) = vbNullString Or Trim(vKYLBO) = "../...." Then
            DisplayMessage "0106", msOKOnly, miCriticalError
            Exit Sub
        Else
            If Len(Trim(vKYLBO)) = 6 Then
                vKYLBO = "'0" & vKYLBO & "'"
            Else
                vKYLBO = "'" & vKYLBO & "'"
            End If
        End If
        
        
        If clsDAO.Connected = False Then
            Me.MousePointer = vbHourglass
            frmSystem.MousePointer = vbHourglass
            clsDAO.CreateConnectionStringSQL spathQHSCC
            clsDAO.Connect
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
        End If
                
        strSQL_DTL = Prepare_QLT
        
        If Trim(strSQL_DTL) <> vbNullString Then
            bln = clsDAO.ExecuteDLL(strSQL_DTL)
            
            
            ' Dong tep
            'If SO_TEP = "50" Then
            
            'Sinh so hieu tep
            
             's = format(NGAY_HIENTAI, "YYMM")
             's = s + DHS_MA
                
             
             'strSQL = "Select top 1 SO_HIEU, NGAY_TAO from QHSCC.dbo.QHS_TEP_HOSO where SO_HIEU like '" & s & "%' order by ID DESC "
             'Set rs = clsDAO.Execute(strSQL)
                
             '   If rs Is Nothing Then
             '       s = s + "-1"
             '   Else
             '       If Left$(rs(0), 4) <> format(NGAY_HIENTAI, "YYMM") Then
             '           s = s + "-1"
             '       Else
             '           I = Right$(rs(0), Len(rs(0)) - InStr(1, rs(0), "-"))
             '           I = I + 1
             '           s = s & "-" & I
             '       End If
             '   End If
                
             '   TEP_ID = s
            
            'Update QHS_SO_HOSO
            'strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set SO_HIEU_TEP = '" & s & "' where SO_HIEU_TEP = '' and DHS_MA = '" + DHS_MA + "' and HTHUC_NOP = '02' and NGUOI_NHAP = '" + USER + "'"
            'bln = clsDAO.ExecuteDLL(strSQL)
            ' insert QHS_TEP_HOSO
            'strSQL = "insert into QHSCC.dbo.QHS_TEP_HOSO (SO_HIEU, DHS_MA, KYKK_TU_NGAY, KYKK_DEN_NGAY, NGAY_TAO, SO_HOSO, TTHAI, NGUOI_TAO) values ('" & s & "', '" & DHS_MA & "', " & format(KYKK_TU_NGAY, "mm/dd/yyyy") & ", " & format(KYKK_DEN_NGAY, "mm/dd/yyyy") & ", '" & format(NGAY_HIENTAI, "mm/dd/yyyy") & "', '" & SO_TEP & "', '', '" & USER & "')"
            'bln = clsDAO.ExecuteDLL(strSQL)
           'End If
          
           Debug.Print strSQL_DTL
        End If
        
        clsDAO.Disconnect
    If Not objTaxBusiness Is Nothing Then
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    StartReceiveForm
    '***************************
    Set rs = Nothing
    blnSaveSuccess = True
    Exit Sub
ErrHandle:

    If Err.Number = -2147217865 Then
        MessageBox "0094", msOKOnly, miCriticalError
    ElseIf Err.Number = 53 Then
        'MessageBox "0096", msOKOnly, miCriticalError
        ' "0109" Thong bao Truoc khi chay ban hay khoi tao ky ke khai ben UD VATCC truoc roi moi nhan bang NTKCC
        MessageBox "0109", msOKOnly, miCriticalError
    Else
        MessageBox "0049", msOKOnly, miCriticalError
        SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    End If
    On Error GoTo ExitErr
    'Rollback
    'clsDAO.RollbackTrans
    clsDAO.Disconnect
    Set rs = Nothing
    blnSaveSuccess = True
    Exit Sub
ExitErr:
    Set rs = Nothing
    SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    MessageBox "0049", msOKOnly, miCriticalError
    blnSaveSuccess = True
End With
End Sub

'****************************
'Description: cmdClear_Click procedure clear current data on
'             the screen and go to next tax report.
'Author:ThanhDX
'Date:23/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub cmdClear_Click()

On Error GoTo ErrHandle
    If Not TAX_Utilities_Svr_New.Data(0) Is Nothing Then
        If MessageBox("0050", msYesNo, miQuestion) = mrYes Then
            If Not objTaxBusiness Is Nothing Then
                objTaxBusiness.Prepared4 dNgayDauKy
                'Get Params
                objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
            End If
            StartReceiveForm
        End If
    End If
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdClear_Click", Err.Number, Err.Description
End Sub

'****************************
'Description: cmdExit_Click procedure.
'Author:ThanhDX
'Date:23/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub cmdExit_Click()
    Dim intReturn As Integer
    Dim blnExit As Boolean
    
On Error GoTo ErrHandle
    
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepared4 dNgayDauKy
        ' Get Params
         objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    
    If Not blnReceiveByBarcode Then 'Receive data from file
        If UBound(arrStrElements) > 0 Then '
            If MessageBox("0055", msYesNo, miQuestion) = mrNo Then _
                Exit Sub
            blnExit = True
        End If
    End If

    If Not TAX_Utilities_Svr_New.Data(0) Is Nothing Then
        Select Case MessageBox("0052", msYesNoCancel, miQuestion)
            Case 1 ' Cancel
                Exit Sub
            Case 3 'No
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            Case 6 'Yes
                cmdSave_Click
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
        End Select
    End If
    
    If blnExit Then
        Unload Me
        frmTreeviewMenu.Show
        Exit Sub
    End If
    
    If blnReceiveByBarcode Then
        If MessageBox("0051", msYesNo, miQuestion) = mrYes Then
            Unload Me
            frmTreeviewMenu.Show
            Exit Sub
        End If
    Else
        Unload Me
        frmTreeviewMenu.Show
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdExit_Click", Err.Number, Err.Description
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrHandle

    Dim strSQL            As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String, strSQL_KHBS As String
    Dim HdrID             As Variant, strDate() As String, dDate As Date
    Dim rs                As ADODB.Recordset, i As Long
    Dim blHoiTonTai       As Integer
    Dim blUpdateTHUETKY2  As Boolean
    Dim bln               As Boolean
    Dim blnKTRB           As Integer
    Dim sSaiCT11          As String
    Dim vKYLBO            As Variant
    Dim vNGAYQUET         As Variant
    Dim vNGAY_DAU_KYLBO   As Variant
    Dim vTHANG_CUOI_KYKK  As Variant
    'dhdang sua loi so sanh ngay
    'ngay 21/10
    Dim vNGAY_DAU_KYLBO1  As Variant
    Dim vTHANG_CUOI_KYKK1 As Variant
    Dim sSQL              As String
    'Dim menuId As Integer
    Dim CHKGIAHAN         As Variant
    Dim vNgayNop          As Variant
    Dim NgayPS            As Variant
    Dim varTemp           As Variant
        
    sSaiCT11 = ""
    
    '***************************
    'ThanhDX added
    'Date:23/11/2005
    
    If TAX_Utilities_Svr_New.Data(0) Is Nothing Then Exit Sub
    '***************************
       
    blnSaveSuccess = False
    
    CallFinish
    ' kiem tra neu trien khai PIT thi se khong nhan cac to khai TNCN theo mau cu
    
    ' Cac to khai PIT se khong nhan to khai co ky ke khai < thang 7 hoac quy 3
        
    If TAX_Utilities_Svr_New.isCheckPIT = True Then
        menuId = Val(GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID"))

        If menuId = 46 Or menuId = 48 Or menuId = 15 Or menuId = 50 Or menuId = 36 Then
            If TAX_Utilities_Svr_New.Year <= 2011 And TAX_Utilities_Svr_New.Month < 7 Then
                MessageBox "0150", msOKOnly, miWarning
                Exit Sub
            End If
        End If

        If menuId = 47 Or menuId = 49 Or menuId = 16 Or menuId = 51 Then
            If TAX_Utilities_Svr_New.Year <= 2011 And TAX_Utilities_Svr_New.ThreeMonths < 3 Then
                MessageBox "0147", msOKOnly, miWarning
                Exit Sub
            End If
        End If
                
        '16/12/2011 dntai: check 2 truong hop ke khai cua to 08_TNCN va 08A_TNCN
        If menuId = 74 Or menuId = 75 Then
            If UCase(objTaxBusiness.kieuKeKhai) = "T" Then
                If Val(Right(objTaxBusiness.vkKhaiTuThang, 4)) <= 2011 And Val(Left(objTaxBusiness.vkKhaiTuThang, 2)) < 7 Then
                    MessageBox "0150", msOKOnly, miWarning
                    Exit Sub
                End If

            Else

                If TAX_Utilities_Svr_New.Year <= 2011 And TAX_Utilities_Svr_New.ThreeMonths < 3 Then
                    MessageBox "0147", msOKOnly, miWarning
                    Exit Sub
                End If
            End If
        End If
                
    End If
              
    ' end

    ' end if
    
    ' Kiem tra xem da khoa so trong ky lap bo nay chua
    ' hlnam edit
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & "\DB_HT\"
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
    End If

    menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")

    ' chi kiem tra validate cho cac mau an chi
    ' Kiem tra ngay nop khong duoc lon hon ngay quet
    If (Val(menuId) >= 64 And Val(menuId) <= 68) Or Val(menuId) = 91 Then
        If CheckValidData = False Then
            MessageBox "0136", msOKOnly, miWarning
            Exit Sub
        End If
    End If

    'nkhoan: kiem tra ngay nop khong dc > ngay hien tai goi PIT
    'co 5 goi voi menuId: 15, 16, 50, 51, 36
    'dntai them cac to : 08_TNCN, 08A_TNCN
    If Val(menuId) = 15 Or Val(menuId) = 16 Or Val(menuId) = 50 Or Val(menuId) = 51 Or Val(menuId) = 30 Or Val(menuId) = 74 Or Val(menuId) = 75 Then
        If CheckValidData = False Then
            MessageBox "0159", msOKOnly, miWarning
            Exit Sub
        End If
    End If
    
    'dntai 21/05/2012 kiem tra check gia han
    ' Kiem tra gia han to khai 01/GTGT
    If Val(menuId) = 1 Then

        ' Lay thong tin ve gia han nop thue GTGT
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 38
            varTemp = .Value
        End With

        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2012 thi thong bao khong duoc gia han nop thue
        '        If Val(TAX_Utilities_Svr_New.Year) = "2012" And (Val(TAX_Utilities_Svr_New.Month) = "4" Or Val(TAX_Utilities_Svr_New.Month) = "5" Or Val(TAX_Utilities_Svr_New.Month) = "6") Then
        '        Else
        '            If Val(varTemp) = 1 Then
        '                MessageBox "0160", msOKOnly, miInformation
        '                Exit Sub
        '            End If
        '        End If
    End If
    
    'End
     
    With fpSpread1
        .Sheet = 1

        If menuId = 8 Or menuId = 9 Then
            .GetText .ColLetterToNumber("I"), 9, vKYLBO
            ' vttoan: lay KyLapBo
        ElseIf menuId = 15 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        ElseIf menuId = 16 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        ElseIf menuId = 50 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        ElseIf menuId = 51 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        ElseIf menuId = 72 Or menuId = 86 Or menuId = 87 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        
        ElseIf menuId = 5 Then
            .GetText .ColLetterToNumber("E"), 23, vKYLBO
        ElseIf menuId = 36 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        ElseIf menuId = 6 Then
            .GetText .ColLetterToNumber("E"), 23, vKYLBO
        ElseIf menuId = 70 Then
            .GetText .ColLetterToNumber("E"), 23, vKYLBO
        ElseIf menuId = 81 Or menuId = 80 Or menuId = 82 Or menuId = 89 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 73 Then
            .GetText .ColLetterToNumber("E"), 42, vKYLBO
        ElseIf menuId = 1 Or menuId = 74 Or menuId = 75 Or menuId = 3 Or menuId = 71 Or menuId = 95 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 2 Or menuId = 59 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 4 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 11 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 12 Or menuId = 77 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
            'dntai 12/05/2011
            'lay VKYLBO cho truong to an chi 01/AC
        ElseIf (menuId >= 64 And menuId <= 68) Or menuId = 91 Then
            vKYLBO = Month(Date) & "/" & Year(Date)
        ElseIf menuId = 23 Then
            .GetText .ColLetterToNumber("D"), 27, vKYLBO
        ElseIf menuId = 25 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 90 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        Else
            .GetText .ColLetterToNumber("E"), 10, vKYLBO
        End If

        'vttoan:
        ' lay ngay nop
        If menuId = 15 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        ElseIf menuId = 16 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        ElseIf menuId = 50 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        ElseIf menuId = 51 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        ElseIf menuId = 36 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        ElseIf menuId = 72 Or menuId = 86 Or menuId = 87 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        
        ElseIf menuId = 5 Then
            .GetText .ColLetterToNumber("E"), 25, vNgayNop
        ElseIf menuId = 6 Then
            .GetText .ColLetterToNumber("E"), 25, vNgayNop
        ElseIf menuId = 70 Then
            .GetText .ColLetterToNumber("E"), 25, vNgayNop
                
        ElseIf menuId = 1 Or menuId = 74 Or menuId = 75 Or menuId = 3 Or menuId = 71 Or menuId = 95 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 2 Or menuId = 59 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 4 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 11 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 12 Or menuId = 77 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 8 Then
            .GetText .ColLetterToNumber("I"), 11, vNgayNop
        ElseIf menuId = 64 Or menuId = 65 Or menuId = 68 Or menuId = 91 Then
            .GetText .ColLetterToNumber("E"), 10, vNgayNop
        ElseIf menuId = 81 Or menuId = 80 Or menuId = 82 Or menuId = 89 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 73 Then
            .GetText .ColLetterToNumber("E"), 44, vNgayNop
        ElseIf menuId = 66 Then
            .GetText .ColLetterToNumber("E"), 13, vNgayNop
        ElseIf menuId = 67 Then
            .GetText .ColLetterToNumber("D"), 12, vNgayNop
        ElseIf menuId = 23 Then
            .GetText .ColLetterToNumber("D"), 29, vNgayNop
        ElseIf menuId = 25 Then
            .GetText .ColLetterToNumber("E"), 32, vNgayNop
        ElseIf menuId = 90 Then
            .GetText .ColLetterToNumber("E"), 24, vNgayNop
        Else
            .GetText .ColLetterToNumber("E"), 12, vNgayNop
        End If
        
        'nkhoan: kiem tra ngay nop khong dc lon hon ngay hien tai
        If Val(menuId) = 80 Or Val(menuId) = 81 Or Val(menuId) = 82 Or Val(menuId) = 73 Or Val(menuId) = 86 Or Val(menuId) = 87 Or Val(menuId) = 59 Or Val(menuId) = 74 Or Val(menuId) = 75 Or Val(menuId) = 3 Or Val(menuId) = 77 Or Val(menuId) = 15 Or Val(menuId) = 16 Or Val(menuId) = 50 Or Val(menuId) = 51 Or Val(menuId) = 36 Or Val(menuId) = 71 Or Val(menuId) = 72 Or Val(menuId) = 89 Then

            If objTaxBusiness.CheckValidData = False Then
                MessageBox "0159", msOKOnly, miWarning
                Exit Sub
            End If
        End If
        
        'ngay nop khong duoc de trong
        If vNgayNop = "" Or vNgayNop = "../../...." Then
            DisplayMessage "0146", msOKOnly, miCriticalError
            clsDAO.Disconnect
            Exit Sub
        End If

        'vttoan:
        'lay ngay phat sinh
        If menuId = 5 Then
            .GetText .ColLetterToNumber("AA"), 44, NgayPS
        ElseIf menuId = 70 Then
            .GetText .ColLetterToNumber("AG"), 42, NgayPS
        ElseIf menuId = 81 Then
            .GetText .ColLetterToNumber("S"), 37, NgayPS
        End If
        
        vNGAY_DAU_KYLBO = "01/" & IIf(Len(Trim(vKYLBO)) = 6, "0" & vKYLBO, vKYLBO)
        
        ' Lay ngay dau cua ky lap bo de xem ngay quet co phu hop voi ky khoa so hay khong?
             
        If Trim(vKYLBO) = vbNullString Or Trim(vKYLBO) = "../...." Then
            DisplayMessage "0106", msOKOnly, miCriticalError
            Exit Sub
        Else

            If Len(Trim(vKYLBO)) = 6 Then
                vKYLBO = "'0" & vKYLBO & "'"
            Else
                vKYLBO = "'" & vKYLBO & "'"
            End If
        End If
        
        ' Ngay dau ky lap bo chua khoa so
        If Trim(vNGAY_DAU_KYLBO) = vbNullString Or Trim(vNGAY_DAU_KYLBO) = "01/../...." Then
            vNGAY_DAU_KYLBO = "CTOD('')"
        Else
            vNGAY_DAU_KYLBO = DateSerial(Int(Mid$(vNGAY_DAU_KYLBO, 7, 4)), Int(Mid$(vNGAY_DAU_KYLBO, 4, 2)), Int(Mid$(vNGAY_DAU_KYLBO, 1, 2)))
            vNGAY_DAU_KYLBO1 = vNGAY_DAU_KYLBO

            'nkhoan: ky lap bo khong dc lon hon thang hien tai
            'dntai 07/01/2012 : sua lai
            If Year(vNGAY_DAU_KYLBO) > Year(Now) Then
                DisplayMessage "0151", msOKOnly, miCriticalError
                clsDAO.Disconnect
                Exit Sub
            ElseIf Year(vNGAY_DAU_KYLBO) = Year(Now) Then

                If Month(vNGAY_DAU_KYLBO) > Month(Now) Then
                    DisplayMessage "0151", msOKOnly, miCriticalError
                    clsDAO.Disconnect
                    Exit Sub
                End If
            End If

            'dntai 2/8/2011 them dkien ky lap bo > 08/2011
            If menuId = 1 Or menuId = 2 Or menuId = 4 Or menuId = 11 Or menuId = 12 Or menuId = 5 Or menuId = 6 Or menuId = 15 Or menuId = 16 Or menuId = 50 Or menuId = 51 Or menuId = 36 Or menuId = 70 Or menuId = 71 Or menuId = 72 Or menuId = 73 Or menuId = 3 Or menuId = 59 Or menuId = 74 Or menuId = 75 Or menuId = 77 Or menuId = 80 Or menuId = 81 Or menuId = 82 Or menuId = 86 Or menuId = 87 Or menuId = 89 Then

                If (TAX_Utilities_Svr_New.Month <> vbNullString) And (TAX_Utilities_Svr_New.Month <> "") Then
                    If (Val(Month(vNGAY_DAU_KYLBO)) < 8) And (Val(Year(vNGAY_DAU_KYLBO)) < 2011) Then
                        DisplayMessage "0143", msOKOnly, miInformation
                        clsDAO.Disconnect
                        Exit Sub
                    End If
                End If
            End If

            'vttoan:
            'to 01_NTNN va 01_TTDB lan phat sinh thi ky lap bo bang ky ke khai van ghi binh thuong
            If (menuId = 5 Or menuId = 70 Or menuId = 81) And NgayPS <> "" Then
                clsDAO.Disconnect

                If (Month(vNGAY_DAU_KYLBO) = Int(TAX_Utilities_Svr_New.Month)) And (Year(vNGAY_DAU_KYLBO) = TAX_Utilities_Svr_New.Year) Then
                    clsDAO.Connect
                End If

            Else

                If TAX_Utilities_Svr_New.isCheckPIT = False And menuId <> 91 Then
                    If (TAX_Utilities_Svr_New.Month <> vbNullString) And (TAX_Utilities_Svr_New.Month <> "") Then
                        If (Month(vNGAY_DAU_KYLBO) = Int(TAX_Utilities_Svr_New.Month)) And (Year(vNGAY_DAU_KYLBO) = TAX_Utilities_Svr_New.Year) Then
                            DisplayMessage "0120", msOKOnly, miCriticalError
                            clsDAO.Disconnect
                            Exit Sub
                        End If
                    End If
                End If
            End If

            'vttoan: ky lap bo phai lon hon ky ke khai
            '29/07/2011
            If (TAX_Utilities_Svr_New.Month <> vbNullString) And (TAX_Utilities_Svr_New.Month <> "") Then
                If (Month(vNGAY_DAU_KYLBO) < Int(TAX_Utilities_Svr_New.Month)) And (Year(vNGAY_DAU_KYLBO) <= TAX_Utilities_Svr_New.Year) Then
                    DisplayMessage "0142", msOKOnly, miCriticalError
                    clsDAO.Disconnect
                    Exit Sub
                End If
            End If
            
            'Ky lap bo phai lon hon ky ke khai doi voi to khai quy
            If (TAX_Utilities_Svr_New.ThreeMonths <> vbNullString) And (TAX_Utilities_Svr_New.ThreeMonths <> "") Then
                Dim NgayDauQuy As Date
                NgayDauQuy = GetNgayDauQuy(CInt(TAX_Utilities_Svr_New.ThreeMonths), CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)

                If (Month(vNGAY_DAU_KYLBO) < Month(NgayDauQuy)) And (Year(vNGAY_DAU_KYLBO) <= TAX_Utilities_Svr_New.Year) Then
                    DisplayMessage "0142", msOKOnly, miCriticalError
                    clsDAO.Disconnect
                    Exit Sub
                End If
            End If

            'end
            
            vNGAY_DAU_KYLBO = "CTOD('" & format(vNGAY_DAU_KYLBO, "mm/dd/yyyy") & "')"
        End If
        
        ' Lay thang cuoi cung cua ky ke khai
        'dhdang edit
        
        '        If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") <> "11" And GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") <> "12" And GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") <> "03" Then
        If (TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "") And LoaiKyKK = False Then
            vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
            vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
        Else
            vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
        End If

        '        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") = "11" Then
        '            .GetText .ColLetterToNumber("E"), 17, CHKGIAHAN
        '            If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
        '                    If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
        '                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
        '                         vTHANG_CUOI_KYKK = "01/02/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
        '                         vTHANG_CUOI_KYKK = "04/05/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
        '                         vTHANG_CUOI_KYKK = "30/07/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
        '                         vTHANG_CUOI_KYKK = "01/11/2010"
        '                        End If
        '                    ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
        '                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
        '                         vTHANG_CUOI_KYKK = "30/07/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
        '                         vTHANG_CUOI_KYKK = "30/11/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
        '                         vTHANG_CUOI_KYKK = "31/01/2011"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
        '                         vTHANG_CUOI_KYKK = "03/05/2011"
        '                        End If
        '                    Else
        '                        If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
        '                            vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        '                        ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
        '                            vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
        '                        Else
        '                            vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
        '                        End If
        '                    End If
        '              Else
        '                    If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
        '                        vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        '                    ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
        '                        vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
        '                    Else
        '                        vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
        '                    End If
        '              End If
        '        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") = "12" Then
        '            .GetText .ColLetterToNumber("E"), 17, CHKGIAHAN
        '            If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
        '                    If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
        '                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
        '                         vTHANG_CUOI_KYKK = "01/02/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
        '                         vTHANG_CUOI_KYKK = "04/05/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
        '                         vTHANG_CUOI_KYKK = "30/07/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
        '                         vTHANG_CUOI_KYKK = "01/11/2010"
        '                        End If
        '                    ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
        '                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
        '                         vTHANG_CUOI_KYKK = "30/07/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
        '                         vTHANG_CUOI_KYKK = "30/11/2010"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
        '                         vTHANG_CUOI_KYKK = "31/01/2011"
        '                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
        '                         vTHANG_CUOI_KYKK = "03/05/2011"
        '                        End If
        '                    Else
        '                        If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
        '                            vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        '                        ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
        '                            vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
        '                        Else
        '                            vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
        '                        End If
        '                    End If
        '            Else
        '                If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
        '                    vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        '                ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
        '                    vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
        '                Else
        '                    vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
        '                End If
        '            End If
        '        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") = "03" Then
        '               .GetText .ColLetterToNumber("E"), 15, CHKGIAHAN
        '               If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
        '                   If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
        '                       vTHANG_CUOI_KYKK = "02/11/2010"
        '                   ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
        '                       vTHANG_CUOI_KYKK = "30/06/2011"
        '                   End If
        '               Else
        '                    If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
        '                        vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        '                    ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
        '                        vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
        '                    Else
        '                        vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
        '                    End If
        '               End If
        '        End If
        '
        'vTHANG_CUOI_KYKK = format(vTHANG_CUOI_KYKK, "dd/mm/yyyy")
        
        vTHANG_CUOI_KYKK = DateSerial(Int(Mid$(vTHANG_CUOI_KYKK, 7, 4)), Int(Mid$(vTHANG_CUOI_KYKK, 4, 2)), Int(Mid$(vTHANG_CUOI_KYKK, 1, 2)))
        
        'CDate(vTHANG_CUOI_KYKK)
        vTHANG_CUOI_KYKK = DateAdd("M", 1, vTHANG_CUOI_KYKK)
        vTHANG_CUOI_KYKK1 = vTHANG_CUOI_KYKK
        vTHANG_CUOI_KYKK = "CTOD('" & format(vTHANG_CUOI_KYKK, "mm/dd/yyyy") & "')"

        ' Ngay quet
        .Sheet = 1

        If menuId = 8 Then  '01_TAIN, 03_TAIN
            .GetText .ColLetterToNumber("T"), 11, vNGAYQUET
        ElseIf menuId = 9 Then ' 02_TAIN
            .GetText .ColLetterToNumber("Q"), 11, vNGAYQUET
        ElseIf menuId = 17 Then ' 04_TNCN
            .GetText .ColLetterToNumber("L"), 12, vNGAYQUET
            'vttoan
        ElseIf menuId = 15 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        ElseIf menuId = 16 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        ElseIf menuId = 50 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        ElseIf menuId = 51 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        ElseIf menuId = 72 Or menuId = 86 Or menuId = 87 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        
        ElseIf menuId = 5 Then
            .GetText .ColLetterToNumber("R"), 25, vNGAYQUET
        ElseIf menuId = 36 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        ElseIf menuId = 6 Then
            .GetText .ColLetterToNumber("R"), 25, vNGAYQUET
        ElseIf menuId = 70 Then
            .GetText .ColLetterToNumber("R"), 25, vNGAYQUET
            
        ElseIf menuId = 11 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
        ElseIf menuId = 1 Or menuId = 74 Or menuId = 75 Or menuId = 3 Or menuId = 71 Or menuId = 95 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
        ElseIf menuId = 2 Or menuId = 59 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
        ElseIf menuId = 4 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
        ElseIf menuId = 12 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
            ' dntai 12/05/2011
            'them truong hop cho to 01_AC
        ElseIf menuId = 65 Then ' 01_AC
            .GetText .ColLetterToNumber("K"), 12, vNGAYQUET
        ElseIf menuId = 64 Then
            .GetText .ColLetterToNumber("K"), 12, vNGAYQUET
        ElseIf menuId = 91 Then
            .GetText .ColLetterToNumber("K"), 12, vNGAYQUET
        ElseIf menuId = 67 Then
            .GetText .ColLetterToNumber("N"), 14, vNGAYQUET
        ElseIf menuId = 68 Then
            .GetText .ColLetterToNumber("K"), 12, vNGAYQUET
        ElseIf menuId = 66 Then
            .GetText .ColLetterToNumber("S"), 15, vNGAYQUET
        ElseIf menuId = 77 Then
            .GetText .ColLetterToNumber("R"), 32, vNGAYQUET
        ElseIf menuId = 81 Or menuId = 80 Or menuId = 82 Or menuId = 89 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
        ElseIf menuId = 73 Then
            .GetText .ColLetterToNumber("M"), 44, vNGAYQUET
        ElseIf menuId = 23 Then
            .GetText .ColLetterToNumber("O"), 29, vNGAYQUET
        ElseIf menuId = 25 Then
            .GetText .ColLetterToNumber("M"), 32, vNGAYQUET
        ElseIf menuId = 90 Then
            .GetText .ColLetterToNumber("M"), 24, vNGAYQUET
        Else
            .GetText .ColLetterToNumber("M"), 12, vNGAYQUET
        End If
        
        vNGAYQUET = DateSerial(Int(Mid$(vNGAYQUET, 7, 4)), Int(Mid$(vNGAYQUET, 4, 2)), Int(Mid$(vNGAYQUET, 1, 2)))

        If Trim(vNGAYQUET) = vbNullString Or Trim(vNGAYQUET) = "../../...." Then
            vNGAYQUET = "CTOD('')"
        Else
            vNGAYQUET = "CTOD('" & format(vNGAYQUET, "mm/dd/yyyy") & "')"
        End If

    End With
     
    sSQL = "SELECT KYLBO, NGAYKHOA FROM KHOASO WHERE KYLBO = " & vKYLBO
    Dim vNGAYKHOASO As Variant
    
    'kiem tra ton tai tep khoaso.dbf chua? Neu chua thong bao de cap nhat VATCC
    Dim strFileName As String
    Dim fso         As New FileSystemObject
    strFileName = spathVat & "\DB_HT\" & "KHOASO.DBF"

    If fso.FileExists(strFileName) = False Then
        DisplayMessage "0111", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    Set rs = clsDAO.Execute(sSQL)

    If Not rs Is Nothing Then
        DisplayMessage "0107", msOKOnly, miInformation
        'dntai 29/07/2011  dong ket noi de phuc vu cho viec load lai to khai
        clsDAO.Disconnect
        Exit Sub
    Else

        If vNGAYQUET < vNGAY_DAU_KYLBO Then
            DisplayMessage "0108", msOKOnly, miInformation
            'dntai 29/07/2011  dong ket noi de phuc vu cho viec load lai to khai
            clsDAO.Disconnect
            Exit Sub
        End If
    End If

    clsDAO.Disconnect
    ' Ket thuc kiem tra khoa so
    
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepared4 dNgayDauKy
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    
    '    If Not objTaxBusiness.TKTTNGNOP Then
    '        MessageBox "0093", msOKOnly, miCriticalError
    '        Exit Sub
    '    End If
    
    If Not objTaxBusiness.TKRB Then
        blnKTRB = MessageBox("0085", msYesNo, miCriticalError)

        If blnKTRB = 6 Then
            sSaiCT11 = "'S'"
            objTaxBusiness.sSaiCT11 = sSaiCT11
        Else
            objTaxBusiness.sSaiCT11 = ""
            Exit Sub
        End If
    End If

    'Kiem tra to khai da ton tai
    If Not objTaxBusiness.TKTT Then

        ' Truong hop chua tao ra file du lieu tu VATCC
        If objTaxBusiness.isExistFile = False Then Exit Sub
        
        If LoaiTk = "" Then

            'dhdang sua loi so sanh ngay
            'ngay 21/10
            If vNGAY_DAU_KYLBO1 > vTHANG_CUOI_KYKK1 Then
                objTaxBusiness.TTHTK = "'4'"
            Else
                objTaxBusiness.TTHTK = "'1'"
            End If

        Else

            ' Neu kiem tra da ton tai to khai thuoc ky ke khai thi dat lai trang thai TTHTK = 2
            If objTaxBusiness.isToKhaiChinhThuc = True Then
                objTaxBusiness.TTHTK = "'2'"
                strSQL_KHBS = objTaxBusiness.InsertDTL_KHBS
            Else ' Neu chua ton tai to khai chinh thuc thi thong bao, to khai bo sung nay ko hop le.
                DisplayMessage "0110", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        If clsDAO.Connected = False Then
            Me.MousePointer = vbHourglass
            frmSystem.MousePointer = vbHourglass
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
        End If
        
        strSQL_DTL = objTaxBusiness.InsertDTL
        
        If Trim(strSQL_DTL) <> vbNullString Then
            bln = clsDAO.ExecuteDLL(strSQL_DTL)
        End If
            
        strSQL_HDR = objTaxBusiness.InsertHDR

        If Trim(strSQL_HDR) <> vbNullString Then
            bln = clsDAO.ExecuteDLL(strSQL_HDR)
        End If
        
        clsDAO.Disconnect
        
    Else

        If LoaiTk = "" Then ' Truong hop to khai da ton tai thi se thay the
            blHoiTonTai = MessageBox("0086", msYesNo, miQuestion)

            If blHoiTonTai = 6 Then
                'Update THUETKY2=1
                blUpdateTHUETKY2 = objTaxBusiness.UpdateTHUETKY2

                If clsDAO.Connected = False Then
                    Me.MousePointer = vbHourglass
                    frmSystem.MousePointer = vbHourglass
                    clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
                    clsDAO.Connect
                    frmSystem.MousePointer = vbDefault
                    Me.MousePointer = vbDefault
                End If
                
                objTaxBusiness.TTHTK = "'3'"
                strSQL_DTL = objTaxBusiness.InsertDTL
                
                If Trim(strSQL_DTL) <> vbNullString Then
                    bln = clsDAO.ExecuteDLL(strSQL_DTL)
                End If
                
                strSQL_HDR = objTaxBusiness.InsertHDR

                If Trim(strSQL_HDR) <> vbNullString Then
                    bln = clsDAO.ExecuteDLL(strSQL_HDR)
                End If

                clsDAO.Disconnect
            End If

        Else ' Day la to khai bo sung
            ' Neu kiem tra da ton tai to khai thuoc ky ke khai thi dat lai trang thai TTHTK = 2
            objTaxBusiness.TTHTK = "'2'"
            strSQL_KHBS = objTaxBusiness.InsertDTL_KHBS
            
            If clsDAO.Connected = False Then
                Me.MousePointer = vbHourglass
                frmSystem.MousePointer = vbHourglass
                clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
                clsDAO.Connect
                frmSystem.MousePointer = vbDefault
                Me.MousePointer = vbDefault
            End If
            
            strSQL_DTL = objTaxBusiness.InsertDTL
            
            If Trim(strSQL_DTL) <> vbNullString Then
                bln = clsDAO.ExecuteDLL(strSQL_DTL)
            End If
                
            strSQL_HDR = objTaxBusiness.InsertHDR

            If Trim(strSQL_HDR) <> vbNullString Then
                bln = clsDAO.ExecuteDLL(strSQL_HDR)
            End If

            'dhdang sua loi thong bao to khai chinh thuc chua ton tai
            'DisplayMessage "0110", msOKOnly, miInformation
            clsDAO.Disconnect
        End If
    End If
        
    If frmSystem.chkSaveQHS = True Then

        'dntai 13/01/2012 khong ghi cac to an chi vao QHS
        If (menuId < 64 Or menuId > 68) And menuId <> 91 Then
            Insert_QHS
        End If
        
    End If

    '
    If Not objTaxBusiness Is Nothing Then
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If

    StartReceiveForm
    '***************************
    Set rs = Nothing
    blnSaveSuccess = True
        
    Exit Sub
ErrHandle:

    If Err.Number = -2147217865 Then
        MessageBox "0094", msOKOnly, miCriticalError
    ElseIf Err.Number = 53 Then
        'MessageBox "0096", msOKOnly, miCriticalError
        ' "0109" Thong bao Truoc khi chay ban hay khoi tao ky ke khai ben UD VATCC truoc roi moi nhan bang NTKCC
        MessageBox "0109", msOKOnly, miCriticalError
    Else
        MessageBox "0049", msOKOnly, miCriticalError
        SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    End If

    On Error GoTo ExitErr
    'Rollback
    'clsDAO.RollbackTrans
    clsDAO.Disconnect
    Set rs = Nothing
    blnSaveSuccess = True
    Exit Sub
ExitErr:
    Set rs = Nothing
    SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    MessageBox "0049", msOKOnly, miCriticalError
    blnSaveSuccess = True
End Sub

Private Function GetLastMonthOfThreeMonth(strPeriod As String) As String
    Select Case strPeriod
        Case "01"
            GetLastMonthOfThreeMonth = "03"
        Case "02"
            GetLastMonthOfThreeMonth = "06"
        Case "03"
            GetLastMonthOfThreeMonth = "09"
        Case "04"
            GetLastMonthOfThreeMonth = "12"
    End Select
End Function

Sub CallFinish(Optional blFinish As Boolean)

'On Error GoTo ErrHandle
'    Dim iSheet As Integer, iActiveSheet As Integer
'    Dim lActiveCol As Long, lActiveRow As Long
'    Dim lCol As Long, lRow As Long
'    Dim i As Integer
'    With fpSpread1
'        .Visible = False
'        .ReDraw = False
'        iActiveSheet = .ActiveSheet
'        lActiveCol = .ActiveCol
'        lActiveRow = .ActiveRow
'
'        For i = 1 To .SheetCount - 1
'            .ActiveSheet = i
'            .SetActiveCell 1, 1
'        Next
'        .ActiveSheet = iActiveSheet
'        .Sheet = iActiveSheet
'        .Col = lActiveCol
'        .Row = lActiveRow
'        .SetActiveCell lActiveCol, lActiveRow
'        .ReDraw = True
'        .Visible = True
'    End With
'    Exit Sub
'ErrHandle:
'    SaveErrorLog Me.Name, "CallFinish", Err.Number, Err.Description
    
    On Error GoTo ErrorHandle
        
    Dim iSheet As Integer, iActiveSheet As Integer
    Dim lActiveCol As Long, lActiveRow As Long
    Dim lCol As Long, lRow As Long
    Dim i As Integer
    With fpSpread1
        .Visible = False
        .ReDraw = False
        .EditMode = False
        iActiveSheet = .ActiveSheet
        lActiveCol = .ActiveCol
        lActiveRow = .ActiveRow
        
        
        For i = 1 To .SheetCount
            .ActiveSheet = i
            .Sheet = .ActiveSheet
            .Row = 1
            .Col = 1
            .Lock = False
            .SetActiveCell 1, 1
            .EditMode = True
        Next

        For i = 1 To .SheetCount
            .ActiveSheet = i
            .Sheet = .ActiveSheet
            .Row = 1
            .Col = 1
            .Lock = True
            .EditMode = False
        Next
        .ActiveSheet = iActiveSheet
        .Sheet = iActiveSheet
        .Col = lActiveCol
        .Row = lActiveRow
        .EditMode = True
        .SetActiveCell lActiveCol, lActiveRow
        .ReDraw = True
        .Visible = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "CallFinish", Err.Number, Err.Description
End Sub

Private Sub cmdViewNow_Click()
    On Error GoTo ErrHandle
    Dim intBarcodeCount As Integer, intBarcodeNo As Integer, intBarcodeIncre As Integer
    Dim strPrefix As String, strBarcodeCount As String, strBarcode As String
    Dim i, j, t, counter As Integer
    Dim chkToKhai As Boolean
    
    For i = 1 To UBound(arrBCBuffer)
        If arrBCBuffer(i) <> vbNullString Then
            intBarcodeCount = intBarcodeCount + 1
        End If
    Next
    ' Khai bao lai mang BCNew luu tat ca cac phan tu khac null trong BCBuffer
    ReDim Preserve arrBCNew(intBarcodeCount)
    ' Clear mang bawcode arrStrElements hien dang co tren Ram de bat dau quet lai
    ReDim arrStrElements(0)
    ' Bat dau se la so thu tu cua chuoi ma vach = 1 sau do se tang dan len
    intBarcodeIncre = 1
    ' Ban dau dat dieu kien la khong co to khai, sau do kiem tra neu quet ma co to khai thi chkToKhai = true
    chkToKhai = False
    For j = 1 To UBound(arrBCBuffer)
        If arrBCBuffer(j) <> vbNullString Then
            counter = counter + 1
            strPrefix = Left$(arrBCBuffer(j), 36)
            strBarcodeCount = Right$(strPrefix, 6)
            strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 6)
            
            strBarcode = Mid$(arrBCBuffer(j), 37)
            intBarcodeNo = CInt(Val(Left$(strBarcodeCount, 3)))
            If intBarcodeNo = 1 Then
                strBarcodeCount = vbNullString
                strBarcodeCount = "001" & Right("000" & intBarcodeCount, 3)
                chkToKhai = True
            Else
                strBarcodeCount = vbNullString
                intBarcodeIncre = intBarcodeIncre + 1
                strBarcodeCount = Right("000" & intBarcodeIncre, 3) & Right("000" & intBarcodeCount, 3)
            End If
            arrBCNew(counter) = strPrefix & strBarcodeCount & strBarcode
        End If
    Next
    ' Neu chua quet to khai ma co yeu cau hien thi thi thong bao phai quet to khai
    If chkToKhai = False Then
        DisplayMessage "0095", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    For t = 1 To UBound(arrBCNew)
        Barcode_Scaned arrBCNew(t)
    Next
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "arrBCBuffer Error!", Err.Number, Err.Description

End Sub

Private Sub Command1_Click()
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String
    
'02_TNCN
'thang
'aa312150100100079   03201200100100100101/0101/01/2010<S01><S></S><S>523453~52345~41234~53453455~0~5234623~412342~0~0~41234~0</S><S>iiiii~26/04/2012~~~1~~</S></S01>

'str1 = "aa312150100100079   03201200100100100101/0101/01/2010<S01><S></S><S>523453~52345~41234~53453455~0~5234623~412342~0~0~41234~0</S><S>iiiii~26/04/2012~~~1~~</S></S01>"
'Barcode_Scaned str1

'02_TNCN quy
'aa312160100100079   01201200400400100101/0101/01/2010<S01><S></S><S>523452346~5234626~523456~0~0~234566~0~0~523456~0~0</S><S>iiiii~02/05/2012~~~1~~</S></S01>

'01_GTGT
'aa313013600247325   04201200700800100101/0114/06/2006<S01><S></S><S>~523452345~2354325~41234234~5134234~534534534~4697812~1423412451~4134234~512344~1423412451~51234~0~539232346~1423412451~1418278217~4123645645~0~0~5018471517~0~5018471517~0~0~0</S><S>~~rtgtggtfhgh~21/05/2012~1~~~1701~x~01</S></S01>

'str1 = "aa313010100100079   04201200200200100101/0114/06/2006<S01><S></S><S>~13573464363~1241234234~4135234534~5234553562~0~8833222~523452345324~2345235~6234534~523452345324~253453~0~8833222~523452345324~518217791762~0~0~0~504644327399~0~504644327399~0~0~0</S><S>~~iiiii~21/05/2012~1~~~1701~x~02</S></S01>"
'Barcode_Scaned str1

'aa313010100100079   04201200600600100101/0114/06/2006<S01><S></S><S>~13573464363~1241234234~52345345~2345345~0~8833222~5234534562~2345235~6234534~5234534562~253453~0~8833222~5234534562~5232189217~23542345~52345345~0~0~0~0~8370078146~0~8370078146</S><S>~~iiiii~21/05/2012~1~~~1701~~03</S></S01>
'str1 = "aa313010100100079   04201100600600100101/0114/06/2006<S01><S></S><S>~13573464363~1241234234~52345345~2345345~0~8833222~5234534562~2345235~6234534~5234534562~253453~0~8833222~5234534562~5232189217~23542345~52345345~0~0~0~0~8370078146~0~8370078146</S><S>~~iiiii~21/05/2012~1~~~1701~~03</S></S01>"
'Barcode_Scaned str1

'aa313013600247325   02201200100100100101/0114/06/2006<S01><S></S><S>~523452345~14234124~14234234~0~0~0~41234124~0~0~41234124~0~0~0~41234124~41234124~0~0~0~0~0~0~482218221~0~482218221</S><S>~~rtgtggtfhgh~21/05/2012~1~~~1701~~</S></S01>
'str1 = "aa313010100100079   02201200100100100101/0114/06/2006<S01><S></S><S>~523452345~14234124~14234234~0~0~0~41234124~0~0~41234124~0~0~0~41234124~41234124~0~0~0~0~0~0~482218221~0~482218221</S><S>~~rtgtggtfhgh~21/05/2012~1~~~1701~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa313512801179169   03201200100100100101/0101/01/2010<S01><S></S><S>100000000~5000000~0~0~0~0~0~0~0~0</S><S>~01/08/2012~~~1~~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa315702801179169   12201200100100100101/0101/01/1900<S01><S></S><S>dfsdf~2222222222~~0~~100000000~10~10~1000000~10000000~5~0~500000~1500000</S><S>100000000~1000000~10000000~0~500000~1500000</S><S>~X</S><S>~~Nguy�n Phi Phong~18/12/2012~1~~~18/12/2012</S></S01>"
'Barcode_Scaned str1

'str1 = "aa316010104183189   01201300100100100101/0114/06/2006<S01><S></S><S>~35329039~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~35329039~0~35329039</S><S>~~~26/02/2013~1~~~1701~x~02</S></S01>"
'Barcode_Scaned str1

'str1 = "aa316110100100079   04201200900900100101/0114/06/2006<S01><S></S><S>100000000~0~100000000~0~0~100000000~0~0~100000000~25~0~25000000~x</S><S></S><S>~~~27/02/2013~1~0~~1052~01</S></S01>"
'Barcode_Scaned str1
'str1 = "aa316120100100079   04201201001000100101/0114/06/2006<S01><S></S><S>~10190000~10000000~190000~10~2~~5~20950~20000~950~0~20950~x</S><S>~27/02/2013~~~1~~1052~03</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   09201300100200101501/0114/06/2006<S01><S>6868686868</S><S>~0~43543543~2177177~2177177~4353423~0~0~0~0~0~0~0~435342"
'Barcode_Scaned str1
'str1 = "aa320010100100079   0920130010020020153~0~-2177177~0~0~243234~0~0~0~0~2177177~0~0~2177177</S><S>fsdfds~asdsad~xcvcxvwref~04/10/2013~1~~~1701~~~0</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002003015<S01_1><S>213~12321~01/09/2013~ewrfwe~2222222222~dfsd~4353423~0~</S><S>~~~~~~0~0~<"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002004015/S><S>~~~~~~0~0~</S><S>~~~~~~0~0~</S><S>~~~~~~0~0~</S><S>4353423~0~0</S></S01_1>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002005015<S01_2><S>322~234~01/09/2013~fds~2222222222~sdfdsfsd~43543543~5~2177177~</S><S>~~~~~~0~0~0~<"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002006015/S><S>~~~~~~0~0~0~</S><S>~~~~~~0~0~0~</S><S>~~~~~~0~0~0~</S><S>43543543~2177177</S></S01_2>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002007015<S01_3><S>dsf~Chi�c~234~234~324~</S><S>~Chi�c~0~0~0~</S></S01_3>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002008015<S01_4A><S>2177177~2177177~0~0~4353423~0~0~0~0</S></S01_4A>"
'Barcode_Scaned str1
'
'str1 = "aa320010100100079   092013001002009015<S01_4B><S>2009~43823122~234234~54345~43534543~435435~43543~10~5435~544~435435~-434891</S></S01_4B>"
'
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002010015<S01_5><S>01/01/2013~243234~10307</S></S01_5>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002011015<S01_6><S>sdf~2222222222~0~0~0~0~0~10"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002012015304~10300</S><S>0~0~0~0</S></S01_6>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002013015<S01_6B><S>dsfsd~2222222222~0~0~0~10304</S><S>0~0~0~0</S></S01_6B>"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002014015<S01_7><S>test~~0~0~0~0~0<"
'Barcode_Scaned str1
'str1 = "aa320010100100079   092013001002015015/S><S>0~0~0~0</S></S01_7>"
'Barcode_Scaned str1

'str1 = "aa320020100100079   03201300100100101001/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0</S><S>sdfsdf~~sdfsdfds~12/11/2013~1~~~1</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001002010<S01_2_1><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_1>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001003010<S01_2_2><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_2>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001004010<S01_2_3><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_3>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001005010<S01_2_4><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_4>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001006010<S01_2_5><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_5>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001007010<S01_2_6><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_6>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001008010<S01_2_7><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_7>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001009010<S01_2_8><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_8>"
'Barcode_Scaned str1
'str1 = "aa320020100100079   032013001001010010<S01_2_9><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_9>"
'Barcode_Scaned str1

'str1 = "aa320040100100079   10201300000000100101/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0</S><S>sdfsdf~~sdfsdfds~12/11/2013~1~~~0</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320950100100079   10201300100300100501/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0</S><S>sdfsdf~~sdfsdfds~12/11/2013~1~~~0</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320950100100079   102013001003002005<S01_1><S>UQE~ueq~01/01/2013~fsdf~6868686868~dsfsd~43223423~</S><S>~~~~~~"
'Barcode_Scaned str1
'str1 = "aa320950100100079   1020130010030030050~</S><S>~~~~~~0~</S><S>01/01/2013~~~~~~0~</S><S>43223423~0</S></S01_1>"
'Barcode_Scaned str1
'str1 = "aa320950100100079   102013001003004005<S01_2><S>SDFDS~23423~01/01/2013~6868686868~6868686868~edrfwer~43534543~</"
'Barcode_Scaned str1
'str1 = "aa320950100100079   102013001003005005S><S>~~~~~~0~</S><S>~~~~~~0~</S><S>~~~~~~0~</S><S>43534543~0</S></S01_2>"
'Barcode_Scaned str1

'str1 = "aa320710100100079   10201300100100100101/0101/01/1900<S01><S>6868686868</S><S>0~0~0</S><S>0~0~0~</S><S>0~0~0~</S><S>0~0~0</S><S>0~0~0~</S><S>0~0~0</S><S>0~0~0~</S><S>0~0</S><S>sdfsdf~~sdfsdfds~12/11/2013~1~~~0</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320720100100079   10201300000000100101/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0</S><S>sdfsdf~~sdfsdfds~13/11/2013~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320120100100079   03201300000000100201/0114/06/2006<S01><S>6868686868</S><S>x~x~0~0~0~0~0~~0~0~0~0~0~0~</S><S>~13/11/2013~sdfsdf~sdfsdfds~1~~1052~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320120100100079   032013000000002002<S01-1><S>0</S><S>~~0~0~</S></S01-1>"
'Barcode_Scaned str1

'str1 = "aa320730100100079   03201300000000100201/0114/06/2006<S02><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~23~0~0~0~0~0~23"
'Barcode_Scaned str1
'str1 = "aa320730100100079   032013000000002002~1~0</S><S>1~~~~~~~</S><S>sdfsdf~sdfsdfds~~13/11/2013~1~~~1053~x~x</S></S02>"
'Barcode_Scaned str1

'str1 = "aa320500100100079   10201300000100100101/0101/01/2010<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0</S><S>ewr~11/11/2013~sdfsdf~sdfsdfds~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320360100100079   03201300000000100101/0101/01/2010<S07><S>6868686868</S><S>~0~0~27000000~27000000~0~0~0~0~0~0~0~0~0~0~20~0</S><S>~13/11/2013~sdfsdf~sdfsdfds~1~~~1</S></S07>"
'Barcode_Scaned str1

'str1 = "aa320740100100079   03201300000000100101/0101/01/1900<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0</S><S>~13/11/2013~sdfsdf~sdfsdfds~1~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320750100100079   03201300100300100201/0101/01/1900<S01><S>6868686868</S><S>65~3~34</S><S>~~0~0~"
'Barcode_Scaned str1
'str1 = "aa320750100100079   0320130010030020020~0~0~0~0~0</S><S>~13/11/2013~sdfsdf~sdfsdfds~1~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320700100100079   10201300000000100101/0101/01/1900<S01><S>6868686868</S><S>~~~0~~0~0~0~0~0~0~0~0~0</S><S>0~0~0~0~0~0</S><S>X~</S><S>sdfsdf~sdfsdfds~~13/11/2013~1~~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320810100100079   10201300000000100101/0101/01/1900<S01><S>6868686868</S><S>~~~0~~0~0~0~0</S><S>0~0~0~0</S><S>1~</S><S>sdfsdf~~sdfsdfds~13/11/2013~1~1~~</S></S01>"
'Barcode_Scaned str1


'str1 = "aa320060100100079   10201300000000100201/0114/06/2006<S01><S>6868686868</S><S>~~0.000~0.00~0.000~0~0.00</S><S>~~"
'Barcode_Scaned str1
'str1 = "aa320060100100079   1020130000000020020.000~0.00~0.000~0~0.00</S><S>sdfsdf~sdfsdfds~~13/11/2013~1~~0</S></S01>"
'Barcode_Scaned str1

'str1 = "aa320684300346630   03201300500500100101/0101/01/2009<S01><S>~~01/07/2013~30/09/2013</S><S>~~~0~~~~~~~0~0~0~~0~~0~~~~0~0</S><S>~sdadsad~14/11/2013~1</S></S01>"
'Barcode_Scaned str1
'str1 = "aa320684300346630   03201300400400100201/0101/01/2009<S01><S>~~01/07/2013~30/09/2013</S><S>Ha ��n gi� tr gia t�ng~01GTKT3/001~SN/11P~111~0000140~0000250~~~0000140~0000250~111~26~10~143;145;148;151;153;160;162;163;165;174~0~~75~176;177;178;179;180;181;182;183;184;185;186;187;188;189;190;191;192;193;194;195;1"
'Barcode_Scaned str1
'str1 = "aa320684300346630   03201300400400200296;197;198;199;200;201;202;203;204;205;206;207;208;209;210;211;212;213;214;215;216;217;218;219;220;221;222;223;224;225;226;227;228;229;230;231;232;233;234;235;236;237;238;239;240;241;242;243;244;245;246;247;248;249;250~~~0~0</S><S>~V��ng H�ng Ngh�a~30/10/2013~1</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999041700260044   03201300000000100101/0114/06/2006<S01><S>6868686868</S><S>1940000~0~0~0~0~-1940000~0</S><S>sdfsdf~~sdfsdfds~05/12/2013~1~~~1</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999151700260044   11201300000000100101/0101/01/2010<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0</S><S>~05/12/2013~sdfsdf~sdfsdfds~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999501700260044   11201300000000100101/0101/01/2010<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0</S><S>~05/12/2013~sdfsdf~sdfsdfds~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999361700260044   05201300000100100101/0101/01/2010<S07><S>6868686868</S><S>~0~0~4000000~4000000~0~0~0~0~0~0~0~0~0~0~20~0</S><S>~06/12/2013~sdfsdf~sdfsdfds~1~~~0</S></S07>"
'Barcode_Scaned str1

'str1 = "aa999741700260044   03201300000000100101/0101/01/1900<S01><S>6868686868</S><S>12~1000000~1900000~12625146~9000000~3600000~123~23~25000~0~0~10000~0~0</S><S>duan~26/11/2013~du�n~kh123~1~</S></S01>"
'Barcode_Scaned str1
'
'str1 = "aa999751700260044   03201300100400100201/0101/01/1900<S01><S>6868686868</S><S>1800000~30000~3000</S><S>adad~1700260044~3~53100~90~0~53100~2655~2~2653~�da~2222222222~4~70800~120~0~70800~3540~3~3537~asd~2300790401~3~5310"
'Barcode_Scaned str1
'str1 = "aa999751700260044   0320130010040020020~90~0~53100~2655~2~2653~qwwww~2300790419~35~619500~1050~0~619500~30975~26~30949~d~2300790426~55~973500~1650~0~973500~48675~41~48634</S><S>duan~06/12/2013~du�n~kh123~1~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999461700260044   11201300000000100101/0101/01/2010<S01><S>6868686868</S><S>0~0~0</S><S>~sdfsdf~sdfsdfds~06/12/2013~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999111700260044   03201300000000100401/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~"
'Barcode_Scaned str1
'str1 = "aa999111700260044   0320130000000020040~0~0~~~~0~0</S><S>~</S><S>sdfsdf~sdfsdfds~~06/12/2013~1~0~~1052</S></S01>"
'Barcode_Scaned str1
'str1 = "aa999111700260044   032013000000003004<S01-1><S>0</S><S>~"
'Barcode_Scaned str1
'str1 = "aa999111700260044   032013000000004004~0~0~</S></S01-1>"
'Barcode_Scaned str1

'str1 = "aa999021700260044   11201300300300101101/0114/06/2006<S01><S>6868686868</S><S>0~0~43909099008~2195454950~0~0~0~0~2195454950~0~2195454950~0~0~0~2195454950</S><S>sdfsdf~~sdfsdfds~06/12/2013~1~~~0</S></S01>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003002011<S01_2><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003003011<S01_2_1><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_1>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003004011<S01_2_2><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_2>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003005011<S01_2_3><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_3>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003006011<S01_2_4><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_4>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003007011<S01_2_5><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_5>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003008011<S01_2_6><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_6>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003009011<S01_2_7><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_7>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003010011<S01_2_8><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_8>"
'Barcode_Scaned str1
'str1 = "aa999021700260044   112013003003011011<S01_2_9><S>02GTTT~02GTTT~001~034~15/10/2013~lan h-�ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2_9>"
'Barcode_Scaned str1

'str1 = "aa999111700260044   03201300100100100401/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~1~Doanh nghi�p c� quy m�"
'Barcode_Scaned str1
'str1 = "aa999111700260044   032013001001002004 nh� v� v�a~01~01/01/2014~342432~32432423</S><S>~</S><S>sdfsdf~sdfsdfds~~06/12/2013~1~0~~1052</S></S01>"
'Barcode_Scaned str1
'str1 = "aa999111700260044   032013001001003004<S01-1><S>0</S><S>~"
'Barcode_Scaned str1
'str1 = "aa999111700260044   032013001001004004~0~0~</S></S01-1>"
'Barcode_Scaned str1

'str1 = "aa999041700260044   11201300100100100101/0114/06/2006<S01><S>2100462770</S><S>435435~43534543~43543534~543543~5435~-982534~0</S><S>test~23432~test~03/12/2013~1~~~0</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999721700260044   12201300000000100101/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0</S><S>sdfsdf~~sdfsdfds~06/12/2013~1~~~1~06/12/2013</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999111700260044   03201300000300100201/0114/06/2006<S01><S>6868686868</S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0"
'Barcode_Scaned str1
'str1 = "aa999111700260044   032013000003002002~0~0~~~~0~0</S><S>~</S><S>sdfsdf~sdfsdfds~~10/12/2013~1~0~~1052</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999231700260044   03201300000100100201/0101/01/1900<S01><S>6868686868</S><S>~~~~~0~0~0~0~0~0~0~0~0~0</S><S>~"
'Barcode_Scaned str1
'str1 = "aa999231700260044   032013000001002002~0~0~0~0~0~0~0~0~0</S><S>sdfsdf~~sdfsdfds~20/12/2013~1~1~~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa999500100231226 11201300100100100101/0101/01/2010<S01><S></S><S>efesf~6868686868~fdvsdsd~3rwfgregtfre~fgdvfd~vgdsfv~wefwe~3432</S><S>4234234~234234~23423~4234~423423~21171~324234~32423~23432423~432432~32423~423432</S><S>dsfsd~24/12"
'Barcode_Scaned str1

'str2 = "bs999021700260044   11201300600900100301/0114/06/2006<S01><S>6868686868</S><S>32423432~343423~405439509435~20271975472~2343242~23423423~3242342~42323432~20253075463~23423423~20262418895~23432432~23434234~234234234~19981317995</S><S>sdfsdf~~sdfsdfds~16/12/2013~~1~1~0</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999021700260044   112013006009002003<SKHBS><S>~~0~0~0</S><S>Thu� GTGT ch�a ���c ho�n k� tr��c chuy�n sang~21~423233~32423432~32000199</S><S>23/12/2"
'Barcode_Scaned str2
'str2 = "bs999021700260044   112013006009003003013~3~12321321~432423432~sdfdsfdsf~01/01/2013~10700~10705~34324234~1234123213~sdfsdfd~0~0~8567767</S></SKHBS>"
'Barcode_Scaned str2

'str2 = "bs999041700260044   11201300200300100101/0114/06/2006<S01><S>0010011000</S><S>1940000~0~0~0~0~-1940000~0</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~20/12/2013~0~1~1~0</S></S01><SKHBS><S>~~0~0~0</S><S>Thu� GTGT ch�a ���c ho�n k� tr��c chuy�n sang~21~423233~32423432~32000199</S><S>23/12/2013~3~12321321~432423432~sdfdsfdsf~01/01/2013~10700~10705~34324234~1234123213~sdfsdfd~0~0~8567767</S></SKHBS>"
'Barcode_Scaned str2

'str2 = "aa999711700260044   10201300200300100101/0101/01/1900<S01><S>0010011000</S><S>20000000~3000000~30000~4000000~200000~5000000~150000~6000000~120000~18000000~500000~38000000~500000</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~19/12/2013~1~~~0</S></S01><S01_1><S>01GTKT~~003~05442~12/10/2013~nga~~~40000000~</S><S>07KPTQ~~004~03243~13/10/2013~h�nh~~~5000000~</S><S>02GTTT~~001~024234~14/10/2013~linh~~~6000000~</S><S>04HGDL~~002~023423~15/10/2013~�nh~~~7000000~~06HDXK~~005~023434~16/10/2013~ph-�ng~~~8000000~</S><S>03XKNB~~006~2423~18/10/2013~kh�i~~~2000000~</S><S>68000000~40000000~28000000</S></S01_1>"
'Barcode_Scaned str2

'str2 = "aa999111700260044   03201300100400100301/0114/06/2006<S01><S>6868686868</S><S>32423432~4324324~28099108~432423423~423432~460099099~432432432~4324423~23342244~0~0~233422"
'Barcode_Scaned str2
'str2 = "aa999111700260044   03201300100400200344~43~43243232432~43~18604627111~0~0~0~18604627111~~~~0~18604627111</S><S>~</S><S>sdfsdf~sdfsdfds~~18/12/2013~1~0~~1052</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999111700260044   032013001004003003<S01-1><S>18604627111</S><S>?wrwewe~2222222222~100~18604627111~10705</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999251700260044   11201300500500100101/0101/01/1900<S01><S>0010011000</S><S>4000000~200000~100000~300000~30000~100000~10000~130000~30000~21000~161000</S><S>Ho�ng Lan~20/12/2013~Huy�n Tr�m~KTV~1~~~0</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999731700260044   03201300400400100101/0114/06/2006<S02><S></S><S>324~1340~432~324~234~3~342~5~-1016~0~-1016~22~0~324~0~45~54~22~1~0</S><S>~1~sda~2222222222~sdfgdfs~5~01/01/2013~01/02/2013</S><S>sdfs~11~22~25/12/2013~1~~25/12/2013~1053~~</S></S02>"
'Barcode_Scaned str2

'str2 = "aa999721700260044   10201300200200100101/0114/06/2006<S01><S>0010011000</S><S>4000000~7000000~40000~140000~180000</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~19/12/2013~1~~~1~24/12/2013</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999161700260044   03201300200200100101/0101/01/2010<S01><S>0010011000</S><S>5000~2000~120000~300000~400000~100000~340000~130000~100000~200000~60000~10000~40000</S><S>Ho�ng Lan~20/12/2013~Huy�n Tr�m~KTV~1~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999511700260044   03201300200200100101/0101/01/2010<S01><S>0010011000</S><S>t� ch�c khai thay ABC~0102030405~�inh Li�t - Ho�n Ki�m - HN~Ho�n Ki�m~HN~0987689965~015678889~abc@yahoo.com</S><S>3000000~150000~1200000~1200~6000000~300000~5000000~500000~2000000~20000~1000000~1000</S><S>Ho�ng Lan~20/12/2013~Huy�n Tr�m~KTV~1~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999361700260044   03201300200200100101/0101/01/2010<S07><S>0010011000</S><S>~90000000~24000000~31190000~27000000~3600000~230000~120000~240000~34810000~2971500~260000~4292~2967208~0~20~0</S><S>Ho�ng Lan~20/12/2013~Huy�n Tr�m~KTV~1~~~0</S></S07>"
'Barcode_Scaned str2

'str2 = "aa999741700260044   03201300200200100101/0101/01/1900<S01><S>0010011000</S><S>40000000~12000000~28000000~13650000~9000000~3600000~500000~300000~250000~14350000~717500~100000~1281~716219</S><S>Ho�ng Lan~20/12/2013~Huy�n Tr�m~KTV~1~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999061700260044   11201300000000100101/0114/06/2006<S01><S>0010011000</S><S>~~0.000~0.00~0.000~0~0.00</S><S>~~0.000~0.00~0.000~0~0.00</S><S>~~0.000~0.00~0.000~0~0.00</S><S>Huy�n Tr�m~KTV~Ho�ng Lan~23/12/2013~1~~0~0~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999051700260044   11201300000000100101/0101/01/1900<S01><S>0010011000</S><S>~0~0~0~0~0</S><S>~~0.00~0~0.00~0.0~0~0~0</S><S>0~0.00~0~0~0</S><S>~~0.00~0~0.00~0.0~0~0~0</S><S>0</S><S>~~0.00~0</S><S>~~0.00~0</S><S>~~0.00~0</S><S>0~0~0~0~0</S><S>Ho�ng Lan~Huy�n Tr�m~KTV~23/12/2013~1~~0~~~0</S></S01><S01-1><S>~~~~~~0~0~0</S><S>H�ng ho�, d�ch v� th�...~0~0</S></S01-1><S01-2><S>~~~~0~0~0~0</S><S>~0~~0~0</S></S01-2>"
'Barcode_Scaned str2

'str2 = "aa999901700260044   11201300000000100101/0101/01/1900<S01><S>0010011000</S><S>~0.000~0~0~</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~23/12/2013~1~~~0~</S></S01><S01-1><S>~~~~0.00~0.00~0.00~0.00~0~0</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999231700260044   03201300000000100101/0101/01/1900<S01><S>0010011000</S><S>~~~~~0~0~0~0~0~0~0~0~0~0</S><S>~~0~0~0~0~0~0~0~0~0</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~22/12/2013~1~1~~1~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999251700260044   04201300500500100101/0101/01/1900<S01><S>0010011000</S><S>4000000~200000~100000~300000~30000~100000~10000~130000~30000~21000~161000</S><S>Ho�ng Lan~20/12/2013~Huy�n Tr�m~KTV~1~~~1</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999011700260044   03201301702000101301/0114/06/2006<S01><S>6868686868</S><S>~0~23421321~324324234~324324234~324234234~0~0~0~0~0~0~0~324234234~0~-32"
'Barcode_Scaned str2
'str2 = "aa999011700260044   0320130170200020134324234~0~0~4492071891~0~0~0~4816396125~0~4816396125</S><S>sdfsdf~sdfsdfds~~13/12/2013~1~~~1701~~~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020003013<S01_1><S>01GTKT~01GTKT~asdas~sad~01/01/2013~wadad~6868686868~sdfsdf~324234234~0~</S><S>~~~~~~~"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020004013~0~0~</S><S>~~~~~~~~0~0~</S><S>~~~~~~~~0~0~</S><S>~~~~~~~~0~0~</S><S>324234234~0~0</S></S01_1>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020005013<S01_2><S>06HDXK~06HDXK~342~rwer~01/01/2013~sfsf~6868686868~fdad~23421321~5~324324234~</S><S>~~~~~~~~0~0"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020006013~0~</S><S>~~~~~~~~0~0~0~</S><S>~~~~~~~~0~0~0~</S><S>~~~~~~~~0~0~0~</S><S>23421321~324324234</S></S01_2>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020007013<S01_3><S>001~12/10/2013~usd~2000000~CK~10 ng�y~S001~13/10/2013~usd~3000000~003~15/10/2013~usd~1500000~009~16/10/2013~usd~60"
'Barcode_Scaned str2
'str2 = "aa999011700260044   03201301702000801300000~008~12/10/2013~usd~4000000~004~11/10/2013~usd~2500000~002~17/10/2013~usd~2700000~ch�ng t�~ghi ch�</S><S></S></S01_3>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020009013<S01_4A><S>324324234~324324234~0~0~324234234~0~0~0~0</S></S01_4A>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020010013<S01_5><S>rew~01/01/2013~edfsd~10101~234324234~rewrwe~01/10/2013~dsfs~11107~23423423~rwerew~07/10/2000~34324~11103~4234324234</S></S01_5>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020011013<S01_6><S>erwe~6868686868~10705~231434~324234~555668~8799~0</S"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020012013><S>111~222~333~444</S></S01_6>"
'Barcode_Scaned str2
'str2 = "aa999011700260044   032013017020013013<S01_7><S>xe tai~Chi�c~21~231~</S><S>xe bus~Chi�c~55~55~</S></S01_7>"
'Barcode_Scaned str2

'str2 = "aa999711700260044   10201300200300100101/0101/01/1900<S01><S>0010011000</S><S>20000000~3000000~30000~4000000~200000~5000000~150000~6000000~120000~18000000~500000~38000000~500000</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~19/12/2013~1~~~0</S></S01><S01_1><S>01GTKT~~003~05442~12/10/2013~nga~~~40000000~</S><S>07KPTQ~~004~03243~13/10/2013~h�nh~~~5000000~</S><S>02GTTT~~001~024234~14/10/2013~linh~~~6000000~</S><S>04HGDL~~002~023423~15/10/2013~�nh~~~7000000~~06HDXK~~005~023434~16/10/2013~ph-�ng~~~8000000~</S><S>03XKNB~~006~2423~18/10/2013~kh�i~~~2000000~</S><S>68000000~40000000~28000000</S></S01_1>"
'Barcode_Scaned str2

'str2 = "bs999041700260044   11201300100100100101/0114/06/2006<S01><S>2100462770</S><S>1~2~3~4~5~6~7</S><S>sdfsdf~~sdfsdfds~05/12/2013~~1~1~0</S></S01><SKHBS><S>~~0~0~0</S><S>Thu� GTGT ch�a ���c ho�n k� tr��c chuy�n sang~21~423233~32423432~32000199</S><S>23/12/2013~3~12321321~432423432~sdfdsfdsf~01/01/2013~10700~10705~34324234~1234123213~sdfsdfd~0~0~8567767</S></SKHBS>"
'Barcode_Scaned str2

'str2 = "aa999121700260044   03201300400400100101/0114/06/2006<S01><S>0010011000</S><S>x~x~560000~520000~400000~100000~20000~40000~23.786~~20.567~22.345~17.678~~0~25724~25724~0~12000~1000~2000~13724~x~03~19/11/2013~40~13684</S><S>Ho�ng Lan~19/12/2013~Huy�n Tr�m~KTV~1~~1052</S></S01><S01-1><S>13724</S><S>doanh nghi�p mh~3600247325~45.78~6283~10703~doanh nghi�p ab~0102030405~54.22~7441~10307</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999741700260044   03201300100100100101/0101/01/1900<S01><S></S><S>32423423~324~32423099~12665170~9000000~3600000~324~32423~32423~19757929~1225793~23423~443~1225350</S><S>ewrwer~25/12/2013~fdsf~r~1~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999751700260044   03201300100100100101/0101/01/1900<S01><S></S><S>3243243~23432~324234</S><S>dgfsf~2600151082~100~3219811~324234~2413243~806568~40328~2031~38297</S><S>dsfdsf~01/01/2013~dsfsd~dfsdf~1~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999701700260044   12201300200200100101/0101/01/1900<S01><S></S><S>fdsfsf~2222222222~01/01~34324~01/01/2013~234324~0~21~0~324234~0~32423423~0~0~dgfdg~6868686868~01/01~234234~01/01/2013~324234~0~43~0~45235~0~23423423~0~0</S><S>0~0~0~0~0~0</S><S>~X</S><S>~~~25/12/2013~1~~~25/12/2013</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999061700260044   11201300100100100101/0114/06/2006<S01><S></S><S>010103~Kg~23.000~23213232.00~11.000~0~213213.00</S><S>010207~Kg~0.000~0.00~15.000~0~0.00~010208~Kg~0~0~10.000~0~0</S><S>010207~Kg~0.000~0.00~15.000~0~0.00~010203~Kg~0~0~15.000~0~0</S><S>~~~25/12/2013~1~~0~0~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999061700260044   12201300100100100101/0114/06/2006<S01><S></S><S>010102~Kg~0.000~0.00~10.000~0~0.00~010103~Kg~0~0~11.000~0~0~010104~Kg~0~0~11.000~0~0~010203~Kg~0~0~15.000~0~0~010208~Kg~0~0~10.000~0~0</S><S>010102~Kg~0.000~0.00~10.000~0~0.00~010103~Kg~0~0~11.000~0~0~010104~Kg~0~0~11.000~0~0</S><S>010104~Kg~0.000~0.00~11.000~0~0.00~010207~Kg~0~0~15.000~0~0~010203~Kg~0~0~15.000~0~0</S><S>dsfsdf~01/01/2013~dsfsdf~25/12/2013~1~~0~1~25/12/2013</S></S01>"
'Barcode_Scaned str2

'str2 = "bs999051700260044   12201300300300100101/0101/01/1900<S01><S></S><S>~23747766~14412243.00~0~0~9335524</S><S>10101~Bao~0.00~23423423~14196014.00~65.0~0~0~9227409~10300~L�t~0~324343~216229.00~50.0~0~0~108115</S><S>21321645555~18540528821.00~112742067~0~2668374668</S><S>20300~~0.00~324234~249411.00~30.0~0~0~74823~20500~~0~21321321321~18540279410.00~15.0~112742067~0~2668299845</S><S>0</S><S>10200~L�t~0.00~0~10103~~0~0</S><S>10101~Bao~0.00~0~104040102~C�i~0~0</S><S>10102~Bao~0.00~0</S><S>21345393321~18554941064.00~112742067~0~2677710192</S><S>11~22~33~26/12/2013~0~1~1~2677710192~26/12/2013~0</S></S01><S01-1><S>�DU�~01dfse~01/01/2013~dsfdsfsd~10102~fdsdf~34~23423~796382</S><S> Thu�c l� �i�u (65 %)~34~796382</S></S01-1><S01-2><S>FDSF~324~01/10/2013~dsfs~3432~321312312~234234~321078078</S><S>sdfsdf~213123~dsfsdfsd~23~23</S></S01-2><SKHBS><S>Thu� TT�B �-�c kh�u tr�~8~0~43243~43243</S><S>~~0~0~0</S><S>26/12/2013~4~423432~324234~werewr~01/01/2013~10300~10303~43240~234234~efrwerwerwe~0~0~167509</S></SKHBS>"
'Barcode_Scaned str2

'str2 = "aa999121700260044   03201300400400100101/0114/06/2006<S01><S>0010011000</S><S>x~x~560000~520000~400000~100000~20000~40000~23.786~~20.567~22.345~17.678~~0~25724~25724~0~12000~1000~2000~13724~x~~19/11/2013~40~13684</S><S>Ho�ng Lan~19/12/2013~Huy�n Tr�m~KTV~1~~1052</S></S01><S01-1><S>13724</S><S>doanh nghi�p mh~3600247325~45.78~6283~10703~doanh nghi�p ab~0102030405~54.22~7441~10307</S></S01-1>"
'Barcode_Scaned str2
'
'str2 = "bs999121700260044   03201300400400100101/0114/06/2006<S01><S>0010011000</S><S>x~x~560000~520000~400000~100000~20000~40000~23.786~~20.567~22.345~17.678~~0~25724~25724~0~12000~1000~2000~13724~x~~19/11/2013~40~13684</S><S>Ho�ng Lan~19/12/2013~Huy�n Tr�m~KTV~0~1~1052</S></S01><S01-1><S>13724</S><S>doanh nghi�p mh~3600247325~45.78~6283~10703~doanh nghi�p ab~0102030405~54.22~7441~10307</S></S01-1><SKHBS><S>Thu� TT�B �-�c kh�u tr�~8~0~43243~43243</S><S>~~0~0~0</S><S>26/12/2013~4~423432~324234~werewr~01/01/2013~10300~10303~43240~234234~efrwerwerwe~0~0~167509</S></SKHBS>"
'Barcode_Scaned str2

'str2 = "aa999861700260044   11201300200200100101/0101/01/1900<S01><S></S><S>Kg~342~23213~010103~Kg~234~23213~010104~Kg~43~213213~010207</S><S>Kg~213~2311321~010104~Kg~44~3432423~010102~Kg~545~45345~010103</S><S>dfdfd~fdgvdgvfd~123213~28/12/2013~1~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999901700260044   12201300100100100101/0101/01/1900<S01><S>6868686868</S><S>L�t~342.000~1000~342000~010101~L�t~324.000~500~162000~010103~L�t~432.000~300~129600~010104</S><S>sdfsdf~~sdfsdfds~14/01/2014~1~~~0~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999231700260044   03201300000000100101/0101/01/1900<S01><S>0010011000</S><S>05/2013~12/2013~~~~0~0~0~0~0~0~0~0~0~0</S><S>~~0~0~0~0~0~0~0~0~0</S><S>Huy�n Tr�m~Ho�ng Lan~KTV~22/12/2013~1~0~~0~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999111700260044   03201301101100100101/0114/06/2006<S01><S>0010011000</S><S>300000~200000~100000~23000~12000~111000~10000~12000~89000~30000~25000~34000~20.897~12000~22.678~21326~15326~7000~8000~6000~x~03~15/11/2013~4000~2000</S><S>x~x</S><S>Huy�n Tr�m~KTV~Ho�ng Lan~19/12/2013~1~0~~1052</S></S01><S01-1><S>6000</S><S>doanh nghi�p ab~0010011000~45.67~2740~10304~c� s� 1~0010011000-001~23.88~1433~10304~c� s�2~0010011000-002~30.45~1827~10304</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999111700260044   03201301101100100101/0114/06/2006<S01><S></S><S>300000~200000~100000~23000~12000~111000~10000~12000~89000~30000~25000~34000~20.897~12000~22.678~21326~15326~7000~8000~6000~x~03~15/11/2013~4000~2000</S><S>x~x</S><S>Huy?n Tr�m~KTV~Ho�ng Lan~19/12/2013~1~0~~1052</S></S01><S01-1><S>6000</S><S>doanh nghi?p ab~0010011000~45.67~2740~10304~co s? 1~0010011000-001~23.88~1433~10304~co s? 2~0010011000-002~30.45~1827~10304</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999011700260044   10201300600600100101/0114/06/2006<S01><S>0010011000</S><S>~32222000~47000000~3750000~267912"
'str2 = str2 & "0~8000000~90000000~6500000~10000000~30000000~1500000~50000000~5000000~98000000~6500000~3820880~500000~400000~7000000~0~0~0~35301120~40000~35261120</S><S>Huy�n Tr�m~KTV~Ho�ng Lan~19/"
'str2 = str2 & "12/2013~1~~~1701~~~0</S></S01><S01_1><S>01GTKT~01GTKT~001~012~11/10/2013~Nam~~tv~5000000~0~~03XKNB~03XKNB~002~012~11/10/2013~Nam~~~3000000~0~"
'str2 = str2 & "</S><S>02GTTT~02GTTT~002~0345~12/10/2013~Chi�n~~t� l�nh~6000000~0~~03XKNB~03XKNB~003~0345~12/10/2013~Chi�n~~~4000000~0~</S><"
'str2 = str2 & "S>06HDXK~06HDXK~003~0567~13/10/2013~Minh~~game~10000000~500000~~02GTTT~02GTTT~004~0567~13/10/2013~Minh~~~20000000~1000000~</S><S>07KPTQ~07KPTQ~004~0345~14/10/2013~Lan~~toy~20000000~2000000~~04HGDL~04HGD"
'str2 = str2 & "L~005~0345~14/10/2013~Lan~~~30000000~3000000~</S><S>03XKNB~03XKNB~005~0123~15/10/2013~H�nh~~b�n t�~4000000~20000~~01THDB~01T"
'str2 = str2 & "HDB~006~0123~15/10/2013~H�nh~~~1000000~10000~</S><S>98000000~6500000~90000000</S></S01_1><S01_2><S>02GTTT~02GTTT~006~567~10/10/2013~Giang~~game~6000000~5~300000~~04HGDL~04HGDL~001~567~10/10/2013~Giang"
'str2 = str2 & "~~t�~8000000~5~400000~</S><S>06HDXK~06HDXK~007~345~11/10/2013~Tr-�ng~~b�n~10000000~10~1000000~~02VEDB~02VEDB~002~345~11/10/2"
'str2 = str2 & "013~Tr-�ng~~tv~5000000~5~250000~</S><S>01GTKT~01GTKT~008~234~12/10/2013~Ph-�ng~~t� l�nh~8000000~10~800000~~01VEDB~01VEDB~003~234~12/10/2013~Ph-�ng~~�i�u h�a~10000000~10~1000000~</S"
'str2 = str2 & "><S>~~~~~~~~0~0~0~</S><S>04HGDL~04HGDL~009~123~19/10/2013~Kh�nh~~�� ��c~1000000~5~50000~~02THDB~02THDB~004~123~19/10/2013~Kh"
'str2 = str2 & "�nh~~th�c ph�m~4000000~5~200000~</S><S>47000000~3750000</S></S01_2><S01_3><S>001~12/10/2013~usd~2000000~CK~10 ng�y~S001~13/10/2013~usd~3000000~003~15/10/2013~usd~1500000~009~16/1"
'str2 = str2 & "0/2013~usd~6000000~008~12/10/2013~usd~4000000~004~11/10/2013~usd~2500000~002~17/10/2013~usd~2700000~ch�ng t�~ghi ch�</S><S><"
'str2 = str2 & "/S></S01_3><S01_4A><S>3750000~700000~1250000~1800000~98000000~90000000~91.84~1800000~1653120</S></S01_4A><S01_4B><S>2012~5600000~5000000~400000~200000~100000000~30000000~30~1200000~360000~34000~326000</S></S01_4B><S01_5><S>CT001~12/10/2013~C�c thu�~~3000000~CT002~12/10/2013~KBNN~~4000000</S></S01_5><S01_6><S>c� s� 1~0102030405~10303~3000000~12000000~15000000~270000~0~c� s� 2~0010011000~10301~5000000~20000000~25000000~450000~0</S><S>0~3000000~0~0</S></S01_6><S01_7><S>mez~Chi�c~200000~150000~c�n 50.000~raver~Chi�c~300000~280000~c�n 20.000</S><S>honda~Chi�c~500000~4800"
'str2 = str2 & "00~c�n 20~suzuki~Chi�c~200000~190000~c�n 10</S></S01_7>"
'Barcode_Scaned str2

'str2 = "aa999061700260044   12201300800800100201/0114/06/2006<S01><S>6868686868</S><S>010103~Kg~234,000~2342342.00~11.000~0~0.00~010104~Kg~234.000~234234.00~11.000~0~0.00~010207~Kg~4.000~324234.00~15.000~0~0.00</S><S>010103~Kg~435.000~32423423.00~11.000~0~0.00~010104~Kg~435.000~432423432.00~11.000~0~0.00~010208"
'Barcode_Scaned str2
'str2 = "aa999061700260044   122013008008002002~Kg~435.000~432432.00~10.000~0~0.00~010203~Kg~653.000~432432.00~15.000~0~0.00</S><S>010103~Kg~546.000~34543534.00~11.000~0~0.00~010104~Kg~765.000~435435.00~11.000~0~0.00~010208~Kg~78.000~43534543.00~10.000~0~0.00</S><S>sdfsdf~sdfsdfds~~13/01/2014~1~~0~0~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999711700260044   03201300100100100101/0101/01/1900<S01><S>6868686868</S><S>5000000~1000000~10000~2000000~100000~3000000~90000~4000000~80000~10000000~280000~15000000~280000</S><S>sdfsdf~~sdfsdfds~15/01/2014~1~~~1</S></S01>"
'Barcode_Scaned str2


'01/TNCN XS thang
'str2 = "aa999481700260044   12201300100100100101/0101/01/2010<S01><S></S><S>32423423~23423423~234234</S><S>dsfsdf~sdfsd~fsdfsd~19/01/2014~1~~</S></S01>"
'Barcode_Scaned str2
''01/TNCN XS quy
'str2 = "aa999491700260044   03201300200200100101/0101/01/2010<S01><S></S><S>23432423~3454345~43543</S><S>sdfsf~sdfsf~gdhgfh~19/01/2014~1~~</S></S01>"
'Barcode_Scaned str2
''01 BHDC thang
'str2 = "aa999251700260044   12201300200200100101/0101/01/1900<S01><S></S><S>35423423~32423423~435435~546467~54647~675775~67578~7567657~4353455~454354~1012014</S><S>dsfsdf~19/01/2014~werwtre~fdbgfdb~1~~~0</S></S01>"
'Barcode_Scaned str2
''01 BHDC quy
'str2 = "aa999251700260044   03201300300300100101/0101/01/1900<S01><S></S><S>234543654~54654645~7657657~564654654~56465465~5465465476~546546548~546546654~657567~78687~610748357</S><S>dfsfsd~19/01/2014~sdg~fdgfdg~1~~~1</S></S01>"
'Barcode_Scaned str2
'02 TNCN thang
'str2 = "aa999154100000018   12201300200200100101/0101/01/2010<S01><S></S><S>655546456~45654654~54654654~546546546~564456~4657657~87686780~89797~435435~76576~0~43544~15315</S><S>fddfhgf~19/01/2014~dfsdf~sfgfdg~1~~</S></S01>"
'Barcode_Scaned str2
''02 TNCN quy
'str2 = "aa999161700260044   03201300200200100101/0101/01/2010<S01><S></S><S>4354654647~54675875~5674575~546546~5645654~765765~54667~45365435~53435435~543543~4354353~5343544~108709</S><S>dfsfds~19/01/2014~3rwrwef~dsfvsdfs~1~~</S></S01>"
'Barcode_Scaned str2
''03 TNCN Thang
'str2 = "aa999501700260044   12201300300300100201/0101/01/2010<S01><S></S><S>cmc~1700260044~Pho duy tan~Cau Giay~Ha Noi~032492058~34243542~sdgfvdgf</S><S>2345324324~117266216~45"
'Barcode_Scaned str2
'str2 = "aa999501700260044   122013003003002002435435~45435~43543543~2177177~546546546~54654655~54654654~546456~65765765~65766</S><S>gfdg~19/01/2014~sdfsdf~234324~1~~</S></S01>"
'Barcode_Scaned str2
''03 TNCN Quy
'str2 = "aa999511700260044   03201300200200100201/0101/01/2010<S01><S></S><S>cmc~1700260044~Duy tan~cau giay~ha noi~342343426~5t23rt43~425rfsrgfedrg</S><S>34324524234~1716226212~324"
'Barcode_Scaned str2
'str2 = "aa999511700260044   03201300200200200223432423~32423432~23432423~1171621~5434543~543454~45443543543~76867867~87987987~87988</S><S>ew~19/01/2014~rewrwe~ewrwer~1~~</S></S01>"
'Barcode_Scaned str2
''07 TNCN thang
'str2 = "aa999361700260044   05201300100300100101/0101/01/2010<S07><S></S><S>~543455435~23423423~4000000~4000000~0~0~0~0~516032012~170761204~0~0~170761204~234234~20~46847</S><S>dsfdsf~19/01/2014~sdfdsf~sdfs~1~~~0</S></S07>"
'Barcode_Scaned str2
''07 TNCN quy
'str2 = "aa999361700260044   04201300200200100101/0101/01/2010<S07><S></S><S>~43535435~5654654~27000000~27000000~0~0~0~0~10880781~544039~435345~2720~541319~43543~20~8709</S><S>efwef~19/01/2014~fregv~ergverg~1~~~1</S></S07>"
'Barcode_Scaned str2
''08 TNCN
'str2 = "aa999741700260044   04201300100100100101/0101/01/1900<S01><S></S><S>4545435435~43543543~4501891892~448470870~9000000~3600000~435435~435435435~0~4053421022~1389147358~435435435~67181087~1321966271</S><S>3rwerwe~19/01/2014~fsds~dsdfsd~1~</S></S01>"
'Barcode_Scaned str2
''08A TNCN
'str2 = "aa999751700260044   04201300100100100201/0101/01/1900<S01><S></S><S>435435435~43435435~54354354</S><S>dfvdsfsdf~1700260044~30~117600000~16306306~0~117600000~19650000~1362325~18287675~fgfdgf~2222222222~30~117600000"
'Barcode_Scaned str2
'str2 = "aa999751700260044   042013001001002002~16306306~0~117600000~19650000~1362325~18287675~fdgfdgd~6868686868~40~156800000~21741742~0~156800000~29490000~2044528~27445472</S><S>dsfsdf~19/01/2014~dsfsd~fdsf~1~</S></S01>"
'Barcode_Scaned str2
'01/TTS Quy
'str2 = "aaa999234100000018   04201300100100100201/0101/01/1900<S01><S></S><S>~~1234~01/10/2013~~4344855855~4321423423~23432432~5~217242793~5~217242793~0~32432~34232432</S><S>dsfsf~6868686868~30~65172838"
'Barcode_Scaned str2
'str2 = "aaa999234100000018   042013001001002002~65172838~0~0~65172838~8084568~0~8084568~sdfsd~0800000016~70~152069955~152069955~0~0~152069955~28267489~0~28267489</S><S>~~~20/01/2014~1~1~~1~~</S></S01>"
'Barcode_Scaned str2
'01/TTS Thang
'str2 = "aaa999234100000018   04201300300300100201/0101/01/1900<S01><S></S><S>01/2014~05/2014~132~01/01/2014~~326588578~324234324~2354254~5~16329429~5~16329429~324324~234234~32423423</S><S>dsfsdf~6868686868~40~65"
'Barcode_Scaned str2
'str2 = "aaa999234100000018   04201300300300200231772~6531772~129730~0~6531772~326589~3243~323346~sdffsd~0800000016~60~9797657~9797657~194594~0~9797657~489883~4865~485018</S><S>~~~20/01/2014~1~1~~0~1~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999154100000018   01201400000000100101/0101/01/2010<S01><S></S><S>12000~200~5000~6000~6000~6000~70000~5000~5000~700~4000~500~140</S><S>~20/01/2014~~~1~~</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999154100000018   01201400000000100101/0101/01/2010<S01><S></S><S>13000~200~5000~6000~6000~5000~70000~5000~5000~700~4000~500~140</S><S>~20/01/2014~~~~1~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999164100000018   02201400000000100101/0101/01/2010<S01><S></S><S>600~100~300~200~210~310~300~200~300~200~100~30~40</S><S>~20/01/2014~~~1~~</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999164100000018   02201400000000100101/0101/01/2010<S01><S></S><S>600~200~300~100~210~310~300~200~300~200~100~300~40</S><S>~20/01/2014~~~~1~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999504100000025   01201400000000100201/0101/01/2010<S01><S></S><S>~~~~~~~</S><S>200~100~20000~20~200000~1"
'Barcode_Scaned str2
'str2 = "aa999504100000025   0120140000000020020000~10000~1000~2100~100~1000~2000</S><S>~20/01/2014~~~1~~</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999504100000025   01201400000000100201/0101/01/2010<S01><S></S><S>~~~~~~~</S><S>400~20~20000~20~200000~10"
'Barcode_Scaned str2
'str2 = "bs999504100000025   012014000000002002000~10000~1000~2100~100~1000~400</S><S>~20/01/2014~~~~1~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999514100000025   02201400000000100201/0101/01/2010<S01><S></S><S>~~~~~~~</S><S>399~20~400~200~3999~20"
'Barcode_Scaned str2
'str2 = "aa999514100000025   0220140000000020020~3488~349~200000~1000~300~299</S><S>~20/01/2014~~~1~~</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999514100000025   02201400000000100201/0101/01/2010<S01><S></S><S>~~~~~~~</S><S>399~20~400~200~3999~20"
'Barcode_Scaned str2
'str2 = "bs999514100000025   0220140000000020020~3488~349~200000~1000~3478~3</S><S>~20/01/2014~~~~1~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999744100000032   01201400000000100101/0101/01/1900<S01><S></S><S>34535345~200~34535145~12608040~9000000~3600000~40~4000~4000~21927105~1442711~340200~7106~1435605</S><S>~20/01/2014~~~1~</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999744100000032   01201400000000100101/0101/01/1900<S01><S></S><S>34535345~3434~34531911~12608040~9000000~3600000~40~4000~4000~21923871~1442387~22443~469~1441918</S><S>~20/01/2014~~~~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999754100000057   01201400100200100201/0101/01/1900<S01><S></S><S>10000~4534~3453</S><S>�dasdf~410000005"
'Barcode_Scaned str2
'str2 = "aa999754100000057   0120140010020020027~100~5466~3453~0~5466~273~86~187</S><S>~01/05/2014~~~1~</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999754100000057   01201400100200100201/0101/01/1900<S01><S></S><S>10000~4234~2234</S><S>�dasdf~410000005"
'Barcode_Scaned str2
'str2 = "bs999754100000057   0120140010020020027~100~5766~2234~0~5766~288~56~232</S><S>~01/05/2014~~~~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999254100000064   01201400000000100101/0101/01/1900<S01><S></S><S>234234~234234~3424~24234~2423~242~24~234234~23432~22009~27880</S><S>~20/01/2014~~~1~~~0</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999254100000064   01201400000000100101/0101/01/1900<S01><S></S><S>234234~204302~3424~24234~2423~2344~234~234234~23432~22009~28090</S><S>~20/01/2014~~~~1~1~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999254100000064   02201400000000100101/0101/01/1900<S01><S></S><S>435435353~353453~3545~242~24~686~69~5435345~6868~686~4324</S><S>~01/05/2014~~~1~~~1</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999254100000064   02201400000000100101/0101/01/1900<S01><S></S><S>53453453~353453~3545~242~24~686~69~5435345~6868~686~4324</S><S>~01/05/2014~~~~1~1~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999368118111076   01201300000000100101/0101/01/2010<S07><S></S><S>~20000~3201~5628267~4000000~1600000~3420~424~24423~0~0~3252~0~0~2230~20~123</S><S>~20/01/2014~~~1~~~0</S></S07>"
'Barcode_Scaned str2
'str2 = "bs999368118111076   01201300000000100101/0101/01/2010<S07><S></S><S>~10000~3201~5628267~4000000~1600000~3420~424~24423~0~0~3252~0~0~3900~20~780</S><S>~20/01/2014~~~~1~1~0</S></S07>"
'Barcode_Scaned str2
'str2 = "aa999368118111076   01201400000000100101/0101/01/2010<S07><S></S><S>~1000~500~30679500~27000000~3600000~34000~40000~5500~0~0~430~0~0~430~20~86</S><S>~20/01/2014~~~1~~~1</S></S07>"
'Barcode_Scaned str2
'str2 = "bs999368118111076   01201400000000100101/0101/01/2010<S07><S></S><S>~1000~600~30668723~27000000~3600000~23223~40000~5500~0~0~430~0~0~430~20~86</S><S>~20/01/2014~~~~1~1~1</S></S07>"
'Barcode_Scaned str2
'str2 = "aa999238118111083   01201400000000100201/0101/01/1900<S01><S></S><S>~~sdfsdf~01/01/2013~~32423432~32423432~0~5~1621172~5~1621172~4324~324234~32423</S><S>fdgsd~6868686868~30~486352"
'Barcode_Scaned str2
'str2 = "aa999238118111083   012014000000002002~486352~1297~0~486352~24318~32~24286~test~8118111083~70~1134820~1134820~3027~0~1134820~56741~76~56665</S><S>~~~01/05/2014~1~1~~1~~</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999238118111083   01201400000000100201/0101/01/1900<S01><S></S><S>01/2014~05/2014~~~~27788959~23423423~4365536~5~1389448~5~1389448~456464~546546~434344</S><S>fsdgdb~8118111083~30~41683"
'Barcode_Scaned str2
'str2 = "aa999238118111083   0120140000000020024~416834~136939~0~416834~20842~3424~17418~fdbb~2222222222~70~972614~972614~319525~0~972614~48631~7988~40643</S><S>~~~01/05/2014~1~1~~0~1~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa320068118111083   02201400100100100101/0114/06/2006<S01><S></S><S>010101~Kg~100.000~1000000.00~12.000~0~0.00~010103~Kg~200.000~300000.00~16.000~0~0</S><S>010102~Kg~150.000~2000000.00~11.000~0~0.00~010201~Kg~100.000~700000.00~15.000~0~0</S><S>~~0~0~0~0~0</S><S>~~~22/02/2014~1~~0~0~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa320868118111083   02201400100100100101/0101/01/1900<S01><S></S><S>Kg~100~1000000~010101~Kg~150~2000000~010103</S><S>Kg~50~1000000~010103~Kg~200~500000~010102~T�n~100~3000000~080101</S><S>Nguyen Van A~123Abc~123456~23/02/2014~1~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa320048118111083   01201400200200100101/0114/06/2006<S01><S></S><S>0~1000000~2000000~3000000~4000000~0~0</S><S>test~test1~test2~23/01/2014~1~~~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa320048118111083   02201400200200100101/0114/06/2006<S01><S></S><S>1000000~7000000~4000000~3000000~2000000~1000000~100000</S><S>test~test1~test2~01/05/2014~1~~~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa320718118111083   01201400100100100101/0101/01/1900<S01><S></S><S>5000000~3000000~30000~2000000~100000~1500000~45000~1000000~20000~7500000~195000~12500000~195000</S><S>test~test1~test2~01/05/2014~1~~~0~</S></S01>"
'Barcode_Scaned str2
'str2 = "aa320718118111083   02201400100100100101/0101/01/1900<S01><S></S><S>5000000~3500000~35000~2500000~125000~1500000~45000~5000000~100000~12500000~305000~17500000~305000</S><S>test~test1~test2~01/05/2014~1~~~1~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa320118118111083   04201300200200100201/0114/06/2006<S01><S></S><S>1000000~200000~800000~100000~50000~850000~400000~300000~150000"
'Barcode_Scaned str2
'str2 = "aa320118118111083   042013002002002002~150000~0~0~~0~0~0~30000~0~0~0~30000~~~~0~0</S><S>~</S><S>~~~17/01/2014~1~0~~1052</S></S01>"
'Barcode_Scaned str2

'str2 = "aa320128118111083   04201300100100100201/0114/06/2006<S01><S></S><S>~~14000000~12000000~5000000~4000000~3000000~2000000~5~"
'Barcode_Scaned str2
'str2 = "aa320128118111083   042013001001002002~20~22~0~~0~94000~94000~0~0~0~0~94000~~~~0~0</S><S>~24/01/2014~~~1~~1052</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320738118111083   04201300200200100201/0114/06/2006<S02><S></S><S>15000000~12000000~4000000~3000000~2000000~1000000~1000000~1000000~3000000~0~3000000~22~66"
'Barcode_Scaned str2
'str2 = "aa320738118111083   0420130020020020020000~0~660000~0~0~22~1~0</S><S>1~~~~~2344~01/01/2013~01/01/2013</S><S>sdfsdf~sdfsdfds~~15/01/2014~1~~~1052~~</S></S02>"
'Barcode_Scaned str2
'
'str2 = "aa320738118111083   02201400100100100201/0114/06/2006<S02><S></S><S>20000000~16500000~3000000~2000000~1500000~5000000~3000000~2000000~3500000~3000"
'Barcode_Scaned str2
'str2 = "aa320738118111083   022014001001002002000~500000~22~110000~0~110000~0~0~22~1~0</S><S>~1~~~~~~</S><S>~~~24/01/2014~1~~24/01/2014~1052~~</S></S02>"
'Barcode_Scaned str2

'str2 = "aa320018118111083   01201400200200100101/0114/06/2006<S01><S></S><S>0~10000000~500000~400000~300000~500000~4500000~3000000~300000~200000~1000000~4000000~2000000~5000000~3000000~2700000~4000000~3000000~1000000~0~0~0~7300000~0~7300000</S><S>~~~10/02/2014~1~~~1701~~~1</S></S01>"
'Barcode_Scaned str2

str2 = "aa320028118111083   01201400100100100101/0114/06/2006<S01><S></S><S>10000000~5000000~3000000~2000000~0~0~0~0~2000000~0~17000000~0~0~0~17000000</S><S>~~~10/02/2014~1~~~1</S></S01>"
Barcode_Scaned str2
'
'str2 = "aa320048118111083   01201400300300100101/0114/06/2006<S01><S></S><S>0~100000000~5000000~10000000~10000000~95000000~9500000</S><S>~~~10/02/2014~1~~~1</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320718118111083   01201400200200100101/0101/01/1900<S01><S></S><S>10000000~4000000~40000~3000000~150000~2000000~60000~1000000~20000~10000000~270000~20000000~270000</S><S>~~~10/02/2014~1~~~1~</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320118118111083   01201400100100100201/0114/06/2006<S01><S></S><S>100000000~5000000~95000000~10000000~10000000~95000000~10000000~10000000~75000000~5"
'Barcode_Scaned str2
'str2 = "aa320118118111083   0120140010010020020000000~5000000~20000000~~0~0~0~11100000~0~0~0~11100000~~~~0~0</S><S>~</S><S>~~~10/02/2014~1~0~~1052</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320128118111083   01201400100100100201/0114/06/2006<S01><S></S><S>~~200000000~180000000~100000000~50000000~30000000~20000000~5"
'Barcode_Scaned str2
'str2 = "aa320128118111083   012014001001002002~~20~22~0~~0~1550000~1550000~0~0~0~0~1550000~~~~0~0</S><S>~10/02/2014~~~1~~1052</S></S01>"
'Barcode_Scaned str2
'
'
'str2 = "aa320738118111083   01201400100100100201/0114/06/2006<S02><S></S><S>100000000~17000000~5000000~3000000~2000000~1000000~2000000~4000000~8300000"
'Barcode_Scaned str2
'str2 = "aa320738118111083   0120140010010020020~0~83000000~22~18260000~0~18260000~0~0~22~1~0</S><S>1~~~~~~~</S><S>~~~10/02/2014~1~~~1052~~</S></S02>"
'Barcode_Scaned str2
'
'
'str2 = "aa320168118111083   01201400200200100101/0101/01/2010<S01><S></S><S>50~20~10~10~10000000~3000000~4000000~5000000~5000000~3000000~2000000~500000~600000</S><S>~10/02/2014~~~1~~</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320518118111083   01201400100100100201/0101/01/2010<S01><S></S><S>~~~~~~~</S><S>100000000~5000000~20000000~20000~3000000~"
'Barcode_Scaned str2
'str2 = "aa320518118111083   012014001001002002150000~4000000~400000~50000000~100000~2000000~2000</S><S>~10/02/2014~~~1~~</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320368118111083   01201400100100100101/0101/01/2010<S07><S></S><S>~230000000~200000~27000000~27000000~0~0~0~0~202800000~43290000~0~0~43290000~0~20~0</S><S>~10/02/2014~~~1~~~1</S></S07>"
'Barcode_Scaned str2
'
'str2 = "aa320748118111083   01201400100100100101/0101/01/1900<S01><S></S><S>100000000~2000000~98000000~10000000~0~0~5000000~3000000~2000000~88000000~12650000~0~0~12650000</S><S>~10/02/2014~~~1~</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320758118111083   01201400100100100201/0101/01/1900<S01><S></S><S>3000000000~20000000~2000000</S><S>nguyen van a~6868686868~50~1490000000~1000000~0~1490000000~491950000~16"
'Barcode_Scaned str2
'str2 = "aa320758118111083   0120140010010020025084~491784916~nguyen van b~8118111083~50~1490000000~1000000~0~1490000000~491950000~165084~491784916</S><S>~10/02/2014~~~1~</S></S01>"
'Barcode_Scaned str2
'
'str2 = "aa320238118111083   01201400100100100201/0101/01/1900<S01><S></S><S>~~123~01/01/2014~101010~12000000~10000000~2000000~5~600000~5~600000~10000~200~2000000</S><S>nguyen van a~6868686868~5"
'Barcode_Scaned str2
'str2 = "aa320238118111083   0120140010010020020~300000~300000~5000~0~300000~5~0~5~nguyen van b~8118111083~50~300000~300000~5000~0~300000~15000~125~14875</S><S>~~~10/02/2014~1~1~~1~~</S></S01>"
'Barcode_Scaned str2


End Sub

Private Sub Form_Activate()
    'dhdang sua check ghi QHS
    '05/07/2010
'    If clsDAO.Connected = False Then
'        frmSystem.chkSaveQHS = False
'    Else
'        frmSystem.chkSaveQHS = True
'    End If
'On Error GoTo ErrHandle
'    If mOnLoad Then
'        mOnLoad = False
'        If UBound(arrStrElements) = 0 Then _
'            Unload Me
'    End If
'    If Not blnReceiveByBarcode And mOnLoad Then
'        mOnLoad = False
'        ShowFormReceiveFromFile
'    End If
'    Exit Sub
'ErrHandle:
'    SaveErrorLog Me.Name, "Form_Activate", Err.Number, Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error GoTo ErrorHandle
'    Dim strHelpContexID As String
'    Dim i As Integer
'    Dim lCol As Long, lRow As Long
'
'    If KeyCode = vbKeyF1 Then
'        fpSpread1.Sheet = mCurrentSheet
'        lCol = fpSpread1.ActiveCol
'        lRow = fpSpread1.ActiveRow
'        GetCellSpan fpSpread1, lCol, lRow
'        strHelpContexID = GetAttribute(xmlDocumentInit(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow)), "HelpContextID") 'Split(GetAttribute(xmlDocumentInit(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow)), "HelpContexID"), "_")
'        If strHelpContexID <> vbNullString Then
'            fpSpread1.HelpContextID = CLng(strHelpContexID) 'Val(strHelpContexID(0) & strHelpContexID(1) & CStr(fpSpread1.ColLetterToNumber(strHelpContexID(2))) & strHelpContexID(3))
'        Else
'            fpSpread1.HelpContextID = 0
'        End If
'    End If
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "Form_KeyDown", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    fpSpread1.EventEnabled(EventAllEvents) = False
    fpSpread1.ImportExcelBook GetAbsolutePath("..\InterfaceTemplates\Template.xls"), vbNullString
    fpSpread1.EventEnabled(EventAllEvents) = True
    
    blnOnLoadEvent = True
    SetControlCaption Me, Me.Name
    frmSystem.chkSaveQuestion.Visible = True
    frmSystem.chkSaveQHS.Visible = True
    
    If CheckConnection = True Then
        frmSystem.chkSaveQHS = True
    Else
        frmSystem.chkSaveQHS = False
        DisplayMessage "0117", msOKOnly, miInformation
    End If
    
    App.HelpFile = App.path & "\HTKK_CQT.chm"
    Me.Top = (frmSystem.ScaleHeight - frmInterfaces.Height) / 2
    If Me.Top <= 0 Then Me.Top = 50
    
    Me.Left = (frmSystem.ScaleWidth - Me.Width) \ 2
    
    If Me.Left <= 0 Then Me.Left = 0
    
    lblLoading.Top = Frame1.Top + (Frame1.Height - lblLoading.Height) / 2
    lblLoading.Left = Frame1.Left + (Frame1.Width - lblLoading.Width) / 2
    lblConnecting.Top = lblLoading.Top
    lblConnecting.Left = lblLoading.Left
    lblExit.Top = lblLoading.Top
    lblExit.Left = lblLoading.Left
    
    If blnReceiveByBarcode Then
        ShowFormReceiveFromBarcode
    Else
        lblLoading.Visible = False
        lblConnecting.Visible = True
    End If
    
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()

On Error GoTo ErrHandle

    'If UBound(arrStrElements) > 0 Then
        SetFormCaption Me, imgCaption, lblCaption
        Me.Refresh
    If Not blnReceiveByBarcode And blnOnLoadEvent Then
        ShowFormReceiveFromFile
        blnOnLoadEvent = False
    End If
        
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Resize", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrHandle
    If blnReceiveByBarcode Then
        StopBarcodeReader
    End If
    
    Set objTaxBusiness = Nothing
    ReDim arrStrElements(0)
    
    frmSystem.chkSaveQuestion.Visible = False
    frmSystem.chkSaveQHS.Visible = False

    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Unload", Err.Number, Err.Description
End Sub

''' fpSpread1_ButtonClicked description
''' Update value for cell (checkbox cell)
''' Parameter1 pCol         : active column
''' Parameter2 pRow         : active row
''' Parameter3 ButtonDown   : left, right or center mouse button
Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    On Error GoTo ErrHandle
    
    With fpSpread1
        .Sheet = .ActiveSheet 'mCurrentSheet
        GetCellSpan fpSpread1, Col, Row
        .Col = Col
        .Row = Row
        If .CellType = CellTypeCheckBox Then
            UpdateCell Col, Row, IIf(ButtonDown = 1, "x", vbNullString)
        End If
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "fpSpread1_ButtonClicked", Err.Number, Err.Description
End Sub

Private Sub fpSpread1_Change(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo ErrHandle
    Dim lValue As String
    Dim IsUpdate As Boolean

    If mOnLoad = True Then Exit Sub
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet 'mCurrentSheet

        .Col = Col
        .Row = Row
        If .Lock = False Then
            ' When user change value of cell, call UpdateCell function
            If .CellType = CellTypeNumber Then
                lValue = .Value
            Else
                lValue = .Text
            End If
            Select Case .CellType
                Case CellTypeCheckBox '10
                    ' Checkbox
                    IsUpdate = UpdateCell(Col, Row, IIf(Val(lValue) = 1, "x", vbNullString))
                Case Else
                    IsUpdate = UpdateCell(Col, Row, lValue)
            End Select
        End If
        .EventEnabled(EventAllEvents) = True
    End With

    Exit Sub

ErrHandle:
    SaveErrorLog Me.Name, "fpSpread1_Change", Err.Number, Err.Description
End Sub

Private Sub fpSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyDelete Then Exit Sub
'    With fpSpread1
'        .Sheet = .ActiveSheet
'        .Col = .ActiveCol
'        .Row = .ActiveRow
'        If .CellType = CellTypeNumber Then
'            .EventEnabled(EventAllEvents) = False
'            '.Text = vbNullString
'            fpSpread1_Change .Col, .Row
'            .EventEnabled(EventAllEvents) = True
'
'        End If
'
'    End With
End Sub







Private Sub MSComm1_OnComm()
    Static strTemp As String
    Dim i As Long
    Dim varBuff As Variant
    Dim lByte() As Byte
        
On Error GoTo ErrHandle
    Select Case MSComm1.CommEvent
        Case comEvReceive                                       ' Received RThreshold  of chars.
            varBuff = MSComm1.Input
            lByte = varBuff
            For i = 0 To UBound(lByte)
                If Chr$(lByte(i)) <> "#" Then
                    strTemp = strTemp & Chr$(lByte(i))
                Else
                    Barcode_Scaned TAX_Utilities_Svr_New.Convert(strTemp, TCVN, UNICODE)
                    strTemp = vbNullString
                End If
            Next
    End Select
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "MSComm1_OnComm", Err.Number, Err.Description
End Sub

'**********************************************
'Description: Barcode_Scaned procedure scan barcode image
'             and process barcode string to check the complete barcode
'Author:ThanhDX
'Date:19/11/2005
'Input:
'       strBarcode: the scanned barcode string
'Output:
'Return:
'**********************************************
Private Sub Barcode_Scaned(strBarcode As String)
    On Error GoTo ErrHandle
    
    Dim intBarcodeCount As Integer, intBarcodeNo As Integer
    Dim strPrefix       As String, strBarcodeCount As String, strData As String
    Dim idToKhai        As String
    
    If Left$(strBarcode, 2) = "bs" Then
        LoaiTk = "TKBS"
    Else
        LoaiTk = ""
    End If

    strBarcode = TrimString(strBarcode)
    'strBarcode = TAX_Utilities_Svr_New.Convert(strBarcode, TCVN, UNICODE)
    
    'dntai: 06/01/2012 chan khong nhan to 03_TAIN
    If Val(Mid$(strBarcode, 4, 2)) = 8 Then
        DisplayMessage "0140", msOKOnly, miCriticalError
        Exit Sub
    End If

    If Left$(strBarcode, 1) <> "0" Then

        'Version 1.2.0 and later
        'If Val(Left$(strBarcode, 3)) > Val(Replace$(APP_VERSION, ".", "")) Then
        If Val(Left$(strBarcode, 3)) > Val(Replace$(HTKK_LAST_VERSION, ".", "")) Then
            'Version tai doanh nghiep lon hon tai co quan thue APP_VERSION
            DisplayMessage "0074", msOKOnly, miCriticalError
            Exit Sub
        ElseIf Val(Left$(strBarcode, 3)) < 320 And InStr(tt156, Mid$(strBarcode, 4, 2)) > 0 Then
            DisplayMessage "0169", msOKOnly, miCriticalError
            Exit Sub
        ElseIf Val(Mid$(strBarcode, 21, 4)) < 2014 And InStr(tt156_tkbs, Mid$(strBarcode, 4, 2)) > 0 And LoaiTk = "TKBS" Then
            DisplayMessage "0140", msOKOnly, miCriticalError
            Exit Sub
        ElseIf Val(Left$(strBarcode, 3)) < 310 Then         'dntai: sua chi nhan nhung version 310 tro len doi voi 03_TNDN

            If Val(Mid$(strBarcode, 4, 2)) = 3 Then
                DisplayMessage "0140", msOKOnly, miCriticalError
                Exit Sub
            End If

        ElseIf Val(Left$(strBarcode, 3)) < 300 Then         'vttoan: sua chi nhan nhung version 300 tro len

            If Val(Mid$(strBarcode, 4, 2)) = 1 Or Val(Mid$(strBarcode, 4, 2)) = 2 Or Val(Mid$(strBarcode, 4, 2)) = 4 Or Val(Mid$(strBarcode, 4, 2)) = 11 Or Val(Mid$(strBarcode, 4, 2)) = 12 Or Val(Mid$(strBarcode, 4, 2)) = 6 Or Val(Mid$(strBarcode, 4, 2)) = 5 Or Val(Mid$(strBarcode, 4, 2)) = 15 Or Val(Mid$(strBarcode, 4, 2)) = 16 Or Val(Mid$(strBarcode, 4, 2)) = 50 Or Val(Mid$(strBarcode, 4, 2)) = 51 Or Val(Mid$(strBarcode, 4, 2)) = 36 Or Val(Mid$(strBarcode, 4, 2)) = 70 Or Val(Mid$(strBarcode, 4, 2)) = 8 Then
                DisplayMessage "0140", msOKOnly, miCriticalError
                Exit Sub
            End If

        ElseIf Val(Left$(strBarcode, 3)) < 200 Then ' Truong hop to khai thue TNCN duoc in bang phien ban 1.3.1 se khong con hieu luc theo luat thue TNCN moi nam 2009

            If Val(Mid$(strBarcode, 4, 2)) = 22 Or Val(Mid$(strBarcode, 4, 2)) = 23 Then
                DisplayMessage "0105", msOKOnly, miCriticalError
                Exit Sub
            End If

            'chan doi voi cac to an chi
        ElseIf Val(Left$(strBarcode, 3)) < 302 Then

            If Val(Mid$(strBarcode, 4, 2)) = 64 Or Val(Mid$(strBarcode, 4, 2)) = 65 Or Val(Mid$(strBarcode, 4, 2)) = 66 Or Val(Mid$(strBarcode, 4, 2)) = 67 Or Val(Mid$(strBarcode, 4, 2)) = 68 Or Val(Mid$(strBarcode, 4, 2)) = 91 Then
                DisplayMessage "0159", msOKOnly, miCriticalError
                Exit Sub
            End If
        End If
        
        strPrefix = Left$(strBarcode, 36)
        strBarcodeCount = Right$(strPrefix, 6)
        strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 6)
        'lay ma TK,MST,DIA CHI
        'DHDANG
        MATK_PRINT = Mid(strBarcode, 4, 2)
        
        ' Bat dau
        ' To khai 04/TNCN bat dau thu thang 2 se ko nhan nua
        If Left$(strPrefix, 3) = "250" Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 04AB/TNCN thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "39" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "40" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0113", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' To khai 07/TNCN phien ban 2.1.0 bat dau thu thang 2 se ko nhan nua
        If Left$(strPrefix, 3) = "210" Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 07/TNCN thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "36" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0116", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' Doi voi cac to khai thang quy/TNCN nay da bi thay doi ID giua version 210 va 250
        ' Dat lai cho ID cua 210 dung voi 250 de nhan vao QLT_NTK
        If Left$(strPrefix, 3) = "210" Or Left$(strPrefix, 3) = "200" Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 02/TNCN thang cua nam 2009 co ID = 15 thi phai set lai gia tri moi co ID = 53
            If Trim(idToKhai) = "15" Then
                strPrefix = Left$(strPrefix, 3) & "53" & Mid(strPrefix, 6, Len(strPrefix) - 5)
            End If

            ' Neu la to khai 03/TNCN thang cua nam 2009 co ID = 16 thi phai set lai gia tri moi co ID = 54
            If Trim(idToKhai) = "16" And UBound(Split(Mid$(strBarcode, 37), "~")) <> 11 Then
                strPrefix = Left$(strPrefix, 3) & "54" & Mid(strPrefix, 6, Len(strPrefix) - 5)
            End If
        End If
        
        ' To khai 02/TNCN, 03/TNCN bat dau thu thang 2 se ko nhan theo TT84 nua
        If (Left$(strPrefix, 3) = "250") Or (Left$(strPrefix, 3) = "210") Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 02AB/TNCN, 03AB/TNCN  thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "53" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "37" And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "54" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "38" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0115", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' To khai 01/TAIN, 02/TAIN, 03/TAIN bat dau thu thang 2 se ko nhan ND 50 2010 CP� doi voi to 01/TAIN v� 02/TAIN va to khai co nien do 2010
        'dhdang sua
        'ngay 22/07/2010
        idToKhai = Mid(strPrefix, 4, 2)

        If (Trim(idToKhai) = "06" And Val(Mid(strPrefix, 19, 2)) >= 7 And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Or (Trim(idToKhai) = "09" And Val(Mid(strPrefix, 19, 2)) >= 7 And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Then
            DisplayMessage "0121", msOKOnly, miInformation
            Exit Sub
        ElseIf (Trim(idToKhai) = "08" And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Then
            DisplayMessage "0122", msOKOnly, miInformation
            Exit Sub
        End If

        ' Ket thuc
        
        strBarcode = Mid$(strBarcode, 37)
        intBarcodeNo = CInt(Val(Left$(strBarcodeCount, 3)))
        intBarcodeCount = CInt(Val(Right$(strBarcodeCount, 3)))
        
        If intBarcodeNo = 0 Or intBarcodeCount = 0 Then
            MessageBox "0054", msOKOnly, miCriticalError
            Exit Sub
        End If
        
        If strTaxReportInfo = "" Then

            ' Scanning
            If UBound(arrStrElements()) = 0 Then
                ProgressBar1.max = intBarcodeCount
                ProgressBar1.Value = 0
                arrStrElements(0) = strPrefix
                cmdViewNow.Enabled = True
            Else

                If IsDifferent(strPrefix, arrStrElements(0)) Then

                    'Another tax report
                    If MessageBox("0035", msYesNo, miQuestion) = mrYes Then
                        ReDim arrStrElements(0)
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                    Exit Sub
                Else

                    If ProgressBar1.max <> intBarcodeCount Then
                        MessageBox "0062", msOKOnly, miCriticalError
                        Exit Sub
                    End If
                End If
            End If
            
            ReDim Preserve arrStrElements(intBarcodeCount)
            arrStrElements(intBarcodeNo) = strBarcode
            
            ' hlnam Edit
            ' Lay them trong truong hop ko quet het ma vach ma muon hien thi luon
            ReDim Preserve arrBCBuffer(intBarcodeCount)
            arrBCBuffer(intBarcodeNo) = strPrefix & strBarcodeCount & strBarcode
            
            If IsCompleteData(strData) Then
                Dim tmp As String
                
                ' To khai 05/GTGT version < 320
                If Val(Left$(strData, 3)) < 320 And Mid$(strData, 4, 2) = "72" Then
                    tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) - 5)
                    strData = tmp & "~0~" & Right$(strData, Len(strData) - InStr(1, strData, "</S01>", vbTextCompare) + 5)
                    
                End If

                ' Check version <= 3.1.6
                If Val(Left$(strData, 3)) <= 316 Then
                    If Mid$(strData, 4, 2) = "01" Or Mid$(strData, 4, 2) = "02" Or Mid$(strData, 4, 2) = "04" Or Mid$(strData, 4, 2) = "71" Or Mid$(strData, 4, 2) = "36" Or Mid$(strData, 4, 2) = "25" Then
                        If Val(idToKhai) <> 36 Then
                            tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) - 5)
                            strData = tmp & "~0" & Right$(strData, Len(strData) - InStr(1, strData, "</S01>", vbTextCompare) + 5)
                        Else
                            strData = Left$(strData, Len(strData) - 10) & "~0" & Right$(strData, 10)
                        End If

                    ElseIf Mid$(strData, 4, 2) = "68" Then
                        tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) - 5)
                        strData = tmp & "~1" & Right$(strData, Len(strData) - InStr(1, strData, "</S01>", vbTextCompare) + 5)
                    ElseIf Mid$(strData, 4, 2) = "73" Then
                        tmp = Mid(strData, 1, InStr(1, strData, "</S02>", vbTextCompare) - 5)
                        strData = tmp & "~" & Right$(strData, Len(strData) - InStr(1, strData, "</S02>", vbTextCompare) + 5)
                    End If
                End If

                If Val(idToKhai) = 1 Or Val(idToKhai) = 2 Or Val(idToKhai) = 4 Or Val(idToKhai) = 71 Or Val(idToKhai) = 36 Or Val(idToKhai) = 68 Or Val(idToKhai) = 25 Then
                    If Val(idToKhai) = 36 Then
                        LoaiKyKK = LoaiToKhai(strData)
                    Else
                        tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) + 5)
                        LoaiKyKK = LoaiToKhai(tmp)
                    End If
                End If

                lblLoading.Visible = False
                lblConnecting.Visible = True
                frmInterfaces.Refresh

                If Not LoadForm(strData) Then
                    StartReceiveForm
                End If

                'Free memory
                ReDim arrStrElements(0)
                ' Khai bao lai mot mang rong
                ' hlnam Edit
                ReDim arrBCBuffer(0)
                ' hlnam End
            End If
                
        Else ' Form is loaded

            If strTaxReportInfo = strPrefix Then
                MessageBox "0044", msOKOnly, miWarning
                Exit Sub
            Else

                If frmSystem.chkSaveQuestion.Value Then
                    cmdSave_Click

                    If blnSaveSuccess Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                Else

                    If MessageBox("0045", msYesNo, miQuestion) = mrYes Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If

    Else
        'Version 1.1.0 and 1.0.0
        strPrefix = Left$(strBarcode, 25)
        strBarcodeCount = Right$(strPrefix, 4)
        strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 4)
        
        strBarcode = Mid$(strBarcode, 26)
        intBarcodeNo = CInt(Val(Left$(strBarcodeCount, 2)))
        intBarcodeCount = CInt(Val(Right$(strBarcodeCount, 2)))
        
        If intBarcodeNo = 0 Or intBarcodeCount = 0 Then
            MessageBox "0054", msOKOnly, miCriticalError
            Exit Sub
        End If
        
        If strTaxReportInfo = "" Then

            ' Scanning
            If UBound(arrStrElements()) = 0 Then
                ProgressBar1.max = intBarcodeCount
                ProgressBar1.Value = 0
                arrStrElements(0) = strPrefix
            Else

                If IsDifferent(strPrefix, arrStrElements(0)) Then

                    'Another tax report
                    If MessageBox("0035", msYesNo, miQuestion) = mrYes Then
                        ReDim arrStrElements(0)
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                    Exit Sub
                Else

                    If ProgressBar1.max <> intBarcodeCount Then
                        MessageBox "0062", msOKOnly, miCriticalError
                        Exit Sub
                    End If
                End If
            End If
            
            ReDim Preserve arrStrElements(intBarcodeCount)
            arrStrElements(intBarcodeNo) = strBarcode
            
            If IsCompleteData(strData) Then
                lblLoading.Visible = False
                lblConnecting.Visible = True
                frmInterfaces.Refresh

                If Not LoadForm(strData) Then
                    StartReceiveForm
                End If

                'Free memory
                ReDim arrStrElements(0)
            End If
                
        Else ' Form is loaded

            If strTaxReportInfo = strPrefix Then
                MessageBox "0044", msOKOnly, miWarning
                Exit Sub
            Else

                If frmSystem.chkSaveQuestion.Value Then
                    cmdSave_Click

                    If blnSaveSuccess Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                Else

                    If MessageBox("0045", msYesNo, miQuestion) = mrYes Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Barcode_Scaned", Err.Number, Err.Description
End Sub

Private Sub GetCells(xmlSectionTemplate As MSXML.IXMLDOMNode, arrStrValue() As String)
    On Error GoTo ErrHandler
    Dim lCtrl As Long, lCtrl2 As Long
    
    'Fill data from array of data to Cell node
    While lCtrl <= UBound(arrStrValue) And Not xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2) Is Nothing
        If GetAttribute(xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2), "Receive") <> "0" Then
            SetAttribute xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2), "Value", _
                Replace(Replace(arrStrValue(lCtrl), "1" & Chr$(20) & Chr$(20) & "1", "#"), Chr$(20), "~")
        Else
            lCtrl = lCtrl - 1
        End If
        lCtrl = lCtrl + 1
        lCtrl2 = lCtrl2 + 1
    Wend
    
    Exit Sub
ErrHandler:
    SaveErrorLog Me.Name, "GetCells", Err.Number, Err.Description
End Sub

'*******************************************
'Description: GetSection procedure convert data from data string
'               to Dom data.
'Author: ThanhDX
'Date: 21/02/2006
'Input:
'   xmlSectionTemplate: Section template node
'   xmlSectionData : Section data node
'*******************************************
Private Sub GetSection(xmlSectionTemplate As MSXML.IXMLDOMNode, xmlSectionData As MSXML.IXMLDOMNode, blnValidData As Boolean)

On Error GoTo ErrHandler

    Dim lCtrl As Long
    Dim lElementsNo As Long
    Dim lDataNo As Long
    Dim arrStrValue() As String
    Dim idToKhaiCheck As Integer
    
    lElementsNo = GetElementsNo(xmlSectionTemplate.childNodes(0))
    'Get array of data units
    arrStrValue = Split(xmlSectionData.Text, "~")
    ' Lay ve so chi tieu cua chuoi ma vach
    lDataNo = UBound(arrStrValue)
    If lDataNo = -1 Then
        lDataNo = 0
    End If
    ' End
    'get id to khai
    idToKhaiCheck = Val(TAX_Utilities_Svr_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    
    If GetAttribute(xmlSectionTemplate, "Dynamic") = "0" Then
        ' to khai 01/GTGT cho phep nhan 2 CTMV
        If idToKhaiCheck = 1 Then
            
            'Static data
            ' Truong hop chuoi ma vach nhieu chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 > lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 3 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 3 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
'        ElseIf idToKhaiCheck = 11 Then
'             If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 1
'                Exit Sub
'            End If
'            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
'            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
'            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 2
'                Exit Sub
'            End If
'        ElseIf idToKhaiCheck = 12 Then
'             If ((lDataNo + 1 > lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 6)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 1
'                Exit Sub
'            End If
'            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
'            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
'            If ((lDataNo + 1 < lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 6)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 2
'                Exit Sub
'            End If
        ElseIf idToKhaiCheck = 3 Then
           If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 3 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
            
        Else
            
            'Static data
    '        If UBound(arrStrValue) + 1 > lElementsNo Then
    '            blnValidData = False
    '            'DisplayMessage "0070", msOKOnly, miCriticalError
    '            Exit Sub
    '        End If
            ' Truong hop chuoi ma vach nhieu chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 > lElementsNo) And isSheetTk Then
            If (lDataNo + 1 > lElementsNo) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If (lDataNo + 1 < lElementsNo) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
            
        End If

    Else
         ' Kiem tra neu chuoi ma vach bi thieu hoac thua chi tieu se khong cho phep nhan
        ' If ((UBound(arrStrValue) + 1) Mod lElementsNo <> 0) And isSheetTk Then
        If ((lDataNo + 1) Mod lElementsNo <> 0) And isSheetTk Then
            blnValidData = False
            checkSoCT = 3
            Exit Sub
        End If
        ' Dynamic data
        For lCtrl = 2 To IIf((UBound(arrStrValue) + 1) Mod lElementsNo = 0, _
                            (UBound(arrStrValue) + 1) / lElementsNo, (UBound(arrStrValue) + 1) \ lElementsNo + 1)
            'Insert nodes
            InsertNode xmlSectionTemplate
        Next lCtrl
    End If
    
    GetCells xmlSectionTemplate, arrStrValue
    
    blnValidData = True
    
    Exit Sub
ErrHandler:
    blnValidData = False
    SaveErrorLog Me.Name, "GetSection", Err.Number, Err.Description
End Sub

'*******************************************
'Description: GetSections procedure convert data from data string
'               to Dom data.
'Author: ThanhDX
'Date: 21/02/2006
'Input:
'   xmlSectionsTemplate: Sections template node
'   xmlSectionsData : Sections data node
'*******************************************
Private Sub GetSections(xmlSectionsTemplate As MSXML.IXMLDOMNode, xmlSectionsData As MSXML.IXMLDOMNode, blnValidData As Boolean)

On Error GoTo ErrHandler

    Dim xmlSectionNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    If xmlSectionsData.childNodes.length > xmlSectionsTemplate.childNodes.length Then
        blnValidData = False
        'DisplayMessage "0072", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    For lCtrl = 1 To xmlSectionsData.childNodes.length
        GetSection xmlSectionsTemplate.childNodes(lCtrl), xmlSectionsData.childNodes(lCtrl - 1), blnValidData
        If Not blnValidData Then
            blnValidData = False
            Exit Sub
        End If
    Next
    blnValidData = True
    
    Exit Sub
ErrHandler:
    blnValidData = False
    SaveErrorLog Me.Name, "GetSections", Err.Number, Err.Description
End Sub

Private Function GetElementsNo(xmlCellsNode As MSXML.IXMLDOMNode) As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCntElementsNo As Long
    
    For Each xmlCellNode In xmlCellsNode.childNodes
        If GetAttribute(xmlCellNode, "Receive") <> "0" Then
            lCntElementsNo = lCntElementsNo + 1
        End If
    Next
    GetElementsNo = lCntElementsNo
End Function

Private Sub InsertNode(xmlSectionTemplate As MSXML.IXMLDOMNode)
    Dim xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode, xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim lRows As Long, lRow2s As Long
    Dim lRowLBound As Long, lRowUbound As Long
    Dim lRow As Long, lCol As Long
    
    Set xmlCellsNode = xmlSectionTemplate.lastChild
    lRows = GetDynRowCount(fpSpread1, xmlCellsNode, lRow2s, lRowLBound, lRowUbound)
    
    'Increase row value on each cell in Dom data
    IncreaseRowInDOM fpSpread1, xmlSectionTemplate.parentNode.parentNode, lRowUbound + 1, lRows, lRow2s
    
    Set xmlNodeNewCells = xmlCellsNode.cloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
        
'        ' Set new ID2 for node (CellID2)
'        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
'        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
    Next
    
    ' Insert new node to DOM object
    xmlCellsNode.parentNode.appendChild xmlNodeNewCells
    
End Sub

Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDOMdata As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDOMdata Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To 0 Step -1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        End If
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "IncreaseRowInDOM", Err.Number, Err.Description
End Sub
'****************************
'Description: RestoreDataFile function restore
'             data files from data string.
'   Step 1: Cut data string into sheet datas
'   Step 2: Load content of sheet datas to DOM, load template to DOM
'   Step 3: Generate xml string and save it to xml file
'Author: ThanhDX
'Date:20/11/2005
'Input:
'       strBarcodeData: Data string.
'OutPut:
'Return: True if restore data file successfully
'        False if the otherwise.
'****************************
Private Function RestoreDataFile(ByVal strBarcodeData As String) As Boolean ', rsTaxInfor As ADODB.Recordset)
'    Dim strDataRestore As String, strFileName As String
'    Dim lIndex As Long, lCtrl As Long, arrStrData() As String
'    Dim xmlData As New MSXML.DOMDocument, xmlTemplate As New MSXML.DOMDocument
'    Dim fso As New FileSystemObject, tstFile As TextStream
'
'On Error GoTo ErrHandle
'    arrStrData = GetSheetDatas(strBarcodeData)
'
'    If UBound(arrStrData) < TAX_Utilities_Svr_New.NodeValidity.childNodes.length Then
'        RestoreDataFile = False
'        Exit Function
'    End If
'
'    For lIndex = 1 To UBound(arrStrData())
'        xmlTemplate.Load GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "TemplateFolder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & ".xml"
'
'        If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
'            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Svr_New.Month & TAX_Utilities_Svr_New.Year & ".xml"
'        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then
'            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Svr_New.ThreeMonths & TAX_Utilities_Svr_New.Year & ".xml"
'        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
'            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_00" & TAX_Utilities_Svr_New.Year & ".xml"
'        End If
'
'        If arrStrData(lIndex) <> vbNullString Then
'            If Not xmlData.loadXML(arrStrData(lIndex)) Then
'                RestoreDataFile = False
'                Exit Function
'            End If
'
'            'Get data string and structure
'            strDataRestore = GetSections(xmlData.firstChild, xmlTemplate.getElementsByTagName("Sections")(0)) ', rsTaxInfor)
'        Else
'            strDataRestore = xmlTemplate.xml
'        End If
'
'        Set tstFile = fso.CreateTextFile(strFileName, True, True)
'
'        tstFile.Write strDataRestore
'        tstFile.Close
'    Next lIndex
'
'    Set xmlData = Nothing
'    Set xmlTemplate = Nothing
'    Set fso = Nothing
'
'    RestoreDataFile = True
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "RestoreDataFile", Err.Number, Err.Description

    Dim strDataRestore As String, strFileName As String
    Dim lIndex As Long, lCtrl As Long, arrStrData() As String
    Dim xmlData As New MSXML.DOMDocument, xmlTemplate As New MSXML.DOMDocument
    Dim fso As New FileSystemObject, tstFile As TextStream
    Dim blnValidData As Boolean
    
On Error GoTo ErrHandle
    arrStrData = GetSheetDatas(strBarcodeData)
    
    If UBound(arrStrData) < TAX_Utilities_Svr_New.NodeValidity.childNodes.length Then
        RestoreDataFile = False
        Exit Function
    End If
    
    For lIndex = 1 To UBound(arrStrData())
           ' Chi kiem tra so chi tieu tren to khai sheet 1
        If lIndex = 1 Then
            isSheetTk = True
        Else
            isSheetTk = False
        End If
        ' end
    
        xmlTemplate.Load GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "TemplateFolder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & ".xml"
        
        If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Svr_New.Month & TAX_Utilities_Svr_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Svr_New.ThreeMonths & TAX_Utilities_Svr_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_00" & TAX_Utilities_Svr_New.Year & ".xml"
        End If
        
        If arrStrData(lIndex) <> vbNullString Then
            If Not xmlData.loadXML(arrStrData(lIndex)) Then
                RestoreDataFile = False
                Exit Function
            End If
            
            'Get data string and structure
            GetSections xmlTemplate.getElementsByTagName("Sections")(0), xmlData.firstChild, blnValidData
            If Not blnValidData Then
                RestoreDataFile = False
                Exit Function
            Else
                xmlTemplate.save strFileName
            End If
        End If
        
        'Set tstFile = fso.CreateTextFile(strFileName, True, True)
        
        'tstFile.Write strDataRestore
        'tstFile.Close
    Next lIndex
    
    Set xmlData = Nothing
    Set xmlTemplate = Nothing
    Set fso = Nothing
    
    RestoreDataFile = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "RestoreDataFile", Err.Number, Err.Description
End Function

'****************************
'Description: SetNodeMenu procedure set value to menu node.
'Author: ThanhDX
'Date:20/11/2005
'Input:
'       strMenuId: Menu id string.
'OutPut:
'Return:
'****************************
Private Sub SetNodeMenu(strMenuId As String)
    Dim xmlMenuDom As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
On Error GoTo ErrHandle
    xmlMenuDom.Load App.path & "\Menu.xml"
    
    For Each xmlNode In xmlMenuDom.getElementsByTagName("Root")(0).childNodes
        If GetAttribute(xmlNode, "ID") = strMenuId Then
            TAX_Utilities_Svr_New.NodeMenu = xmlNode
            Exit For
        End If
    Next
    
    Set xmlNode = Nothing
    Set xmlMenuDom = Nothing
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetNodeMenu", Err.Number, Err.Description
End Sub

'****************************
'Description: SetPeriod procedure set value to month, threemonth and year property.
'Author: ThanhDX
'Date:20/11/2005
'Input:
'       strValue: Value set.
'OutPut:
'Return:
'****************************
Private Sub SetPeriod(ByVal strValue As String)
    Dim TkID As String

    On Error GoTo ErrHandle
    TkID = Left$(strTaxReportInfo, 2)

    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
        If TkID = "01" Or TkID = "02" Or TkID = "04" Or TkID = "71" Or TkID = "36" Or TkID = "68" Or TkID = "25" Then
            TAX_Utilities_Svr_New.Month = Left$(strValue, 2)
            TAX_Utilities_Svr_New.ThreeMonths = Left$(strValue, 2)
        Else
            TAX_Utilities_Svr_New.Month = Left$(strValue, 2)
            TAX_Utilities_Svr_New.ThreeMonths = ""
        End If
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = 1 Then

        If TkID = "68" Then
            TAX_Utilities_Svr_New.Month = Left$(strValue, 2)
        Else
            TAX_Utilities_Svr_New.Month = ""

        End If

        TAX_Utilities_Svr_New.ThreeMonths = Left$(strValue, 2)
    End If
    
    TAX_Utilities_Svr_New.Year = Right$(strValue, 4)
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetPeriod", Err.Number, Err.Description
End Sub

'****************************
'Description: InitParameters function initialize parameters from data string
'   Step 1:
'   Step 2:
'   Step 3:
'Author:ThanhDX
'Date:25/11/2005
'Input:
'       strData: data string of tax report.
'Output:
'       rsTaxInfor: recordset contain query data.
'Return: true if initialize sucessfully
'        false if the otherwise
'****************************
'Private Function InitParameters(ByVal strData As String, ByRef rsTaxInfor As ADODB.Recordset) As Boolean
Private Function InitParameters(ByVal strData As String, arrStrHeaderData() As String) As Boolean

'ThanhDX modified
'Date: 10/04/06
    Dim strTaxID As String, strID As String
    Dim blnConnected As Boolean
    Dim strValidDate As String, strTempDate As String
    Dim rsParams As ADODB.Recordset
    Dim strPhongXuLy As String
    Dim rsTaxInfor As ADODB.Recordset
    Dim strTaxID_DLT As String
    Dim strTaxReportInfo_DLT As String
On Error GoTo ErrHandle
    
    TAX_Utilities_Svr_New.Month = ""
    TAX_Utilities_Svr_New.ThreeMonths = ""
    TAX_Utilities_Svr_New.Year = ""
    TAX_Utilities_Svr_New.FinanceStartDate = ""
    
'    If Left$(strData, 1) = "0" Then
'        strTaxReportVersion = "1.1.0"
'        lblVersion.caption = "1.1.0"
''**********************************
    
'    If Left$(strData, 3) = "120" Then
'        lblVersion.caption = "1.2.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "130" Then
'    'Version 1.3.0
'        'Get version of application
'        lblVersion.caption = "1.3.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "131" Then
'    'Version 1.3.1
'        'Get version of application
'        lblVersion.caption = "1.3.1"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "200" Then
'    'Version 2.0.0
'        'Get version of application
'        lblVersion.caption = "2.0.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "210" Then
'    'Version 2.1.0
'        'Get version of application
'        lblVersion.caption = "2.1.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "250" Then
'    'Version 2.1.0
'        'Get version of application
'        lblVersion.caption = "2.5.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "252" Then
'        'Version 2.1.0
'        'Get version of application
'        lblVersion.caption = "2.5.2"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    Else
'        'Version 2.5.3
'        'Get version of application
'        lblVersion.caption = "2.5.3"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    End If

    ' 17122010 - sua lai doan lay version cua ung dung in ma vach
    strTaxReportVersion = Left$(strData, 3)
    strData = Mid$(strData, 4)
    lblVersion.caption = Left$(strTaxReportVersion, 1) & "." & Mid$(strTaxReportVersion, 2, 1) & "." & Right$(strTaxReportVersion, 1)
    ' end doan lay version
    
    'Get info of barcode string --25 characters
    strTaxReportInfo = Left$(strData, 21)
    
'    If xmlSQL.url = "" Then
'        xmlSQL.Load App.path & "\SQL.xml"
'    End If
'
    'Get Tax id
    strTaxID = Trim(Mid$(strTaxReportInfo, 3, 13))
'htphuong sua bo dau gach ngang 13 so
'    If Len(strTaxID) = 13 Then
'        strTaxID = Mid$(strTaxID, 1, 10) & "%" & Mid$(strTaxID, 11, 13)
'    End If
    
    'Connect DB and get informations
    Set rsTaxInfor = GetTaxInfo(strTaxID, blnConnected)
    
     'Connect DB fail
    If Not blnConnected Then _
        Exit Function
    strMST = strTaxID
  
   'vttoan: lay mst Dai Ly thue
    strID = Left$(strTaxReportInfo, 2)
   If strID = 80 Then
        strTaxReportInfo_DLT = Left$(strData, 78)
   Else
        strTaxReportInfo_DLT = Left$(strData, 58)
   End If
    
       strTaxID_DLT = Mid(strData, InStr(1, strData, "<S>") + 3, InStr(1, strData, "</S>") - InStr(1, strData, "<S>") - 3)
       strMST_DLT = strTaxID_DLT
   
    If InStr(1, strData, "<S") < 35 Then
        iNgayTaiChinh = 0
        iThangTaiChinh = 0
    Else
        'Ver 1.1.0 and later
        ' Get NgayTaiChinh and ThangTaiChinh
        strTempDate = Mid$(strData, 22, 5)
        iNgayTaiChinh = GetNgayTaiChinh(strTempDate)
        iThangTaiChinh = GetThangTaiChinh(strTempDate)
        TAX_Utilities_Svr_New.FinanceStartDate = strTempDate
    End If
    
    strID = Left$(strTaxReportInfo, 2)
    SetNodeMenu strID
    SetPeriod Right$(strTaxReportInfo, 6)
    TAX_Utilities_Svr_New.NodeValidity = GetValidityNode
    
    '*******************************
    'ThanhDX added
    'Date: 13/02/2006
    'Gan gia tri tu ngay, den ngay.
    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Day") = "1" Then
        TAX_Utilities_Svr_New.FirstDay = Mid$(strData, 37, 10)
        TAX_Utilities_Svr_New.LastDay = Mid$(strData, 47, 10)
    End If
'*******************************
'*******************************
'ThanhDX added
'Date: 16/02/2006
    'Danh sach to khai can kiem tra ngay bat dau nam tai chinh
    On Error GoTo ThamSoErrHandle
    
    'Set rsParams = clsDAO.Execute("select gia_tri from rcv_thamso where ten ='LOAI_TK_TAICHINH'")
    
    On Error GoTo ErrHandle
    'Kiem tra ngay bat dau nam tai chinh doi voi cac loai to
    '   khai co kiem tra ngay bat dau nam tai chinh
'    If InStr(1, "," & rsParams.Fields(0) & ",", "," & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") & ",") <> 0 Then
'        If Not IsNull(rsTaxInfor.Fields("ngay_tchinh")) Then
'            If Mid$(rsTaxInfor("ngay_tchinh"), 1, 5) <> Mid$(strTempDate, 1, 5) Then
'                DisplayMessage "0065", msOKOnly, miCriticalError
'                Exit Function
'            End If
'            'Kiem tra ngay bat dau kinh doanh
'        Else 'Trong DB chua co gia tri ngay bat dau kinh doanh
'            DisplayMessage "0066", msOKOnly, miCriticalError
'            Exit Function
'        End If
'    End If
'    'To khai ke khai tu ngay ... den ngay ...
'    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "FinanceYear") = "1" Then
'        If IsNull(rsTaxInfor("ngay_kdoanh")) Then
'            'Trong DB chua co gia tri ngay bat dau kinh doanh
'            DisplayMessage "0067", msOKOnly, miCriticalError
'            Exit Function
'        Else
'            If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Day") = "1" Then
'                If Mid$(rsTaxInfor("ngay_kdoanh"), 1, 5) <> Mid$(TAX_Utilities_Svr_New.FirstDay, 1, 5) _
'                   And Mid$(rsTaxInfor("ngay_tchinh"), 1, 5) <> Mid$(TAX_Utilities_Svr_New.FirstDay, 1, 5) Then
'                   'Tu ngay phai bang ngay bat dau nam tai chinh
'                   '    hoac ngay bat dau kinh doanh
'                    DisplayMessage "0068", msOKOnly, miCriticalError
'                    Exit Function
'                End If
'                'Ky ke khai lon hon ngay bat dau kinh doanh
'                If CInt(Mid$(rsTaxInfor("ngay_kdoanh"), 7, 4)) > CInt(Mid$(TAX_Utilities_Svr_New.FirstDay, 7, 4)) Then
'                    DisplayMessage "0069", msOKOnly, miCriticalError
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
    
    'Kiem tra cach thuc tinh ky ke khai la tinh theo
    ' nam duong lich hay nam tai chinh
    On Error GoTo ThamSoErrHandle
    
    '    Set rsParams = clsDAO.Execute("select gia_tri from rcv_thamso where ten ='THEO_NAM_TAICHINH'")
    '    blnTinhTheoNamTaiChinh = IIf(rsParams.Fields(0) = 0 Or IsNull(rsParams.Fields(0)), False, True)
    
    On Error GoTo ErrHandle

    'Gan gia tri ngay dau ky
    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
        dNgayDauKy = DateSerial(CInt(TAX_Utilities_Svr_New.Year), CInt(TAX_Utilities_Svr_New.Month), 1)
        dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)

        If Val(strID) = 1 Or Val(strID) = 2 Or Val(strID) = 4 Or Val(strID) = 71 Or Val(strID) = 36 Or Val(strID) = 68 Or Val(strID) = 95 Then
            If LoaiKyKK = True Then
                dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Svr_New.ThreeMonths), CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)
                dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
                dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
            End If

        End If

    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then

        If Val(strID) = 68 And LoaiKyKK = False Then
            dNgayDauKy = DateSerial(CInt(TAX_Utilities_Svr_New.Year), CInt(TAX_Utilities_Svr_New.Month), 1)
            dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
            dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
        Else
            dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Svr_New.ThreeMonths), CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)
            dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
            dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
        End If

    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
        dNgayDauKy = GetNgayDauNam(CInt(TAX_Utilities_Svr_New.Year), iThangTaiChinh, iNgayTaiChinh)
        dNgayCuoiKy = DateAdd("m", 12, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1/2" Then
        dNgayDauKy = GetNgayDauNam(CInt(TAX_Utilities_Svr_New.Year), iThangTaiChinh, iNgayTaiChinh)
        dNgayCuoiKy = DateAdd("m", 12, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    End If

    '*******************************
    '*******************************
    'ThanhDX added
    'Date: 11/01/2006
    'Check validity of start date.
    
    If InStr(1, strData, "<S") < 35 Then
        'Ver 1.0
        strTempDate = Mid$(strData, 22, 8)
        strValidDate = GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "StartDate")
        '        If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))) _
        '                , DateSerial(CInt(Mid$(strTempDate, 5, 4)), CInt(Mid$(strTempDate, 3, 2)), CInt(Mid$(strTempDate, 1, 2)))) = 0 Then
        '            DisplayMessage "0064", msOKOnly, miInformation
        '            Exit Function
        '        End If
        '*******************************
        'Get main content
        strData = Mid$(strData, 30)
    Else

        'Ver 1.1.0 and later
        '        strTempDate = Mid$(strData, 27, 10)
        '        strValidDate = GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "StartDate")
        '        If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))) _
        '                , DateSerial(CInt(Mid$(strTempDate, 7, 4)), CInt(Mid$(strTempDate, 4, 2)), CInt(Mid$(strTempDate, 1, 2)))) = 0 Then
        '            DisplayMessage "0064", msOKOnly, miInformation
        '            Exit Function
        '        End If
        '*******************************
        'Get main content
        If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Day") <> "0" Then
            strData = Mid$(strData, 57)
        Else
            strData = Mid$(strData, 37)
        End If
    End If
    
    'RestoreDataFile (strData)
    If Not RestoreDataFile(strData) Then  ', rsTaxInfor

        'So chi tieu tren ma vach nhieu hon so chi tieu tren to khai
        If checkSoCT = 1 Then
            MessageBox "0133", msOKOnly, miCriticalError
            Exit Function
        End If

        ' So chi tieu tren ma vach it hon so chi tieu tren to khai
        If checkSoCT = 2 Then
            MessageBox "0134", msOKOnly, miCriticalError
            Exit Function
        End If

        ' Kiem tra cac to khai co so dong dong (chi kiem tra duoc khac chu khong phan biet duoc truong hop thieu hoac thua)
        If checkSoCT = 3 Then
            MessageBox "0135", msOKOnly, miCriticalError
            Exit Function
        End If
        
        If blnReceiveByBarcode Then
            MessageBox "0057", msOKOnly, miCriticalError
        Else
            MessageBox "0053", msOKOnly, miCriticalError
        End If

        Exit Function
    End If
    
    '***********************************
    'ThanhDX added
    'Date 26/05/06
    
    'Gan thong tin Header vao mang
    If Not GetHeaderData(rsTaxInfor, arrStrHeaderData) Then
        DisplayMessage "0080", msOKOnly, miCriticalError
        Exit Function
    End If
    
    'Lay thong tin ma so tep va so thu tu to khai.
    If Not GetThongTinTep(strID, arrStrHeaderData) Then
        DisplayMessage "0079", msOKOnly, miCriticalError
        Exit Function
    End If

    '***********************************
    InitParameters = True

    Exit Function
ThamSoErrHandle:
    DisplayMessage "0078", msOKOnly, miCriticalError
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "InitParameters", Err.Number, Err.Description
End Function

Private Sub SetupSpread()
    On Error GoTo ErrHandle
    
    Dim lSheet As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .Sheet = lSheet
            .AllowCellOverflow = False
            .AllowEditOverflow = False
            .Appearance = AppearanceFlat
            .ArrowsExitEditMode = True
            '.GrayAreaBackColor = RGB(238, 238, 238)
            .GrayAreaBackColor = vbButtonFace
            
            .MaxCols = .DataColCnt - 1
            .MaxRows = .DataRowCnt - 1
            .GridShowHoriz = False
            .GridShowVert = False
            
            .EditModePermanent = True
            .EditModeReplace = True
            .ColHeadersShow = False
            .RowHeadersShow = False
            .BorderStyle = BorderStyleNone
            .EditEnterAction = EditEnterActionNext
            .ProcessTab = True
            .ScrollBarExtMode = True
            .ScrollBarTrack = ScrollBarTrackOff
            .ScrollBars = ScrollBarsBoth  'ScrollBarsVertical
            .SetActionKey ActionKeyClear, False, False, 0
            .TabStripPolicy = TabStripPolicyAsNeeded
            .TabStripFont.Name = "Tahoma"
            
            .TextTip = TextTipFloating
        
            If UCase(.SheetName) <> UCase("Header") Then
                .SheetName = GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(.Sheet - 1), "Caption")
            Else
                mHeaderSheet = .Sheet
            End If
            
            
            .SetTextTipAppearance "Tahoma", 8, False, False, RGB(255, 255, 235), &H0
            .Protect = True
        Next
        .ActiveSheet = 1
        .Sheet = 1
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "SetupSpread", Err.Number, Err.Description
End Sub

Private Sub FormatGrid()
    On Error GoTo ErrHandle
    
    Dim lSheet As Long, i As Long, j As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .Sheet = lSheet
            If .SheetVisible = True Then
                For i = 1 To .MaxRows
                    .Row = i
                    If .RowHeight(i) > 10 And .RowHeight(i) < 15 Then .RowHeight(i) = 14
                    For j = 1 To .MaxCols
                        .Col = j
                        
                        If .BackColor = 12632256 Then
                            'Form backcolor
                            '.BackColor = RGB(238, 238, 238)
                            .BackColor = vbButtonFace
                            Me.BackColor = .BackColor
                        End If
                        
                        If .BackColor = 9868950 Then
                            'Grid header backcolor
                            .BackColor = RGB(215, 215, 215)
                        End If
                        
                        If .BackColor = 16777164 Then
                            'Grid hight light 1 backcolor
                            .BackColor = RGB(233, 245, 254)
                        End If
                        
                        If .BackColor = 13434879 Then
                            'Grid hight light 2 backcolor
                            .BackColor = RGB(255, 255, 235)
                        End If
                        
                        If .CellType = CellTypeNumber Then
                            .TypeNumberDecimal = ","
                            .TypeNumberSeparator = "."
                            .TypeNumberNegStyle = TypeNumberNegStyle2
                        End If
                        
                        Select Case Trim(.Text)
                            Case "chk"
                                .CellType = CellTypeCheckBox
                                .TypeCheckCenter = True
                            Case "cbo"
                                .CellType = CellTypeComboBox
                                .Text = ""
                            Case "cmd"
                                .CellType = CellTypeButton
                            Case "picture"
                                .CellType = CellTypePicture
                        End Select
                    Next
                Next
            End If
        Next
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "FormatGrid", Err.Number, Err.Description
End Sub

'****************************
'Description: LoadForm function load tax report to screen
'   Step 1: Convert data string into data files.
'   Step 2: Load excel template and format grid.
'   Step 3: Load and fill data from data files to grid.
'Author:ThanhDX
'Date:25/11/2005
'Input:
'       strData: data string of tax report.
'Output:
'Return: true if load data sucessfully
'        false if the otherwise
'****************************
Private Function LoadForm(ByVal strData As String) As Boolean

    On Error GoTo ErrHandle

    Dim rsHeaderData       As ADODB.Recordset
    Dim arrStrHeaderData() As String
    Dim rs                 As ADODB.Recordset, strSQL As String
    Dim blnConnected       As Boolean
    Dim strPhongXuLy       As String
    Dim result             As Integer
    'dntai them bien de luu ket qua ma vach dung cho cac to 08_TNCN,08A_TNCN
    Dim strCellMavach      As String
    Dim strArrMavach()     As String
    
    Dim LoaiTk1            As String
    'lay id tkhai
    
    Me.MousePointer = vbHourglass
    frmSystem.MousePointer = vbHourglass
    
    'If InitParameters(strData, rsHeaderData) = False Then
    If InitParameters(strData, arrStrHeaderData) = False Then
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
        Exit Function
    End If
          
    mOnLoad = True
    fpSpread1.EventEnabled(EventAllEvents) = False
    LoadTemplate fpSpread1
    SetupSpread

    FormatGrid
    'LoadInitFiles
        
    'dntai 4/8/2011
    'check neu to khai khbs co kykkhai < 06/2011 hoac qui <03/2011 thi khong nhan
    
    LoaiTk1 = Mid(strData, 4, 2)

    If LoaiTk = "TKBS" Then
        If Trim(LoaiTk1) = "74" Or Trim(LoaiTk1) = "75" Then
            'cat section thu 2 cua to 08_TNCN, 08A_TNCN
            '                strCellMavach = Mid(Trim(strData), InStr(InStr(1, Trim(strData), "<S>") + 1, strData, "<S>") + 3, InStr(InStr(1, Trim(strData), "</S>") + 4, Trim(strData), "</S>") - (InStr(InStr(1, Trim(strData), "<S>") + 1, strData, "<S>") + 3))
            strArrMavach = Split(strData, "~")

            If Trim(LoaiTk1) = "74" Then
                If Trim(strArrMavach(2)) = vbNullString And Trim(strArrMavach(3)) = vbNullString Then
                    If checkKyKHBSTo08("Q") = False Then
                        DisplayMessage "0145", msOKOnly, miInformation
                        Exit Function
                    End If

                Else

                    If checkKyKHBSTo08("T" & Trim(strArrMavach(3))) = False Then
                        DisplayMessage "0145", msOKOnly, miInformation
                        Exit Function
                    End If
                End If
            End If

            If Trim(LoaiTk1) = "75" Then
                If Trim(Right(strArrMavach(0), 7)) = "</S><S>" And Trim(strArrMavach(1)) = vbNullString Then
                    If checkKyKHBSTo08("Q") = False Then
                        DisplayMessage "0145", msOKOnly, miInformation
                        Exit Function
                    End If

                Else

                    If checkKyKHBSTo08("T" & Trim(strArrMavach(1))) = False Then
                        DisplayMessage "0145", msOKOnly, miInformation
                        Exit Function
                    End If
                End If
            End If

        Else

            If checkKyKHBS(Val(LoaiTk1)) = False Then
                DisplayMessage "0145", msOKOnly, miInformation
                Exit Function
            End If
        End If
    End If
    
    'end
    
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & "\dtnt\"
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
    End If
    
    strSQL = "SELECT madtnt, tengoi, dchi, nganh, mabpql, dthoai, fax "
    strSQL = strSQL & " FROM dtnt2 where madtnt = '" & strMST & "'"
    Set rs = clsDAO.Execute(strSQL)
    
    If Not rs Is Nothing Then
        strTenGoi = rs.Fields("tengoi")
        strTenGoi = Trim(strTenGoi)
        
        strDchi = rs.Fields("dchi")
        strDchi = Trim(strDchi)
        
        strNganh = rs.Fields("nganh")
        strNganh = Trim(strNganh)
        
        strMaBPQL = rs.Fields("mabpql")
        strMaBPQL = Trim(strMaBPQL)
        
        strDThoai = rs.Fields("dthoai")
        strDThoai = Trim(strDThoai)
        
        strFax = rs.Fields("fax")
        strFax = Trim(strFax)
    Else
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
        Beep 600, 500
        MessageBox "0087", msOKOnly, miCriticalError
        LoadForm = False
        clsDAO.Disconnect
        Exit Function
        
    End If
    
    If strMaBPQL <> vbNullString Then
        strSQL = "SELECT mabpql, tengoi from dmbpql where Mabpql =  '" & strMaBPQL & "'"
        Set rs = clsDAO.Execute(strSQL)

        If Not rs Is Nothing Then
            strTenBpql = rs.Fields("tengoi")
            strTenBpql = Trim(strTenBpql)
        End If
    End If
    
    ' Lay thong tin phong quan ly phuc vu cac mau AC
    If strMaBPQL <> vbNullString Then
        strSQL = "SELECT mabpql, tengoi from dmbpql where Mabpql =  '" & Mid$(strMaBPQL, 1, 2) & "'"
        Set rs = clsDAO.Execute(strSQL)

        If Not rs Is Nothing Then
            strMaPhongQuanLy = Mid$(strMaBPQL, 1, 2)
            strTenPhongQuanLy = Trim(rs.Fields("tengoi"))
        End If
    End If
    
    clsDAO.Disconnect
    
    ' check trien khai PIT
    TAX_Utilities_Svr_New.isCheckPIT = checkActivePIT

    ' end check
    '    --------------------------
    'lay thong tin ve dai ly thue
    'bo qua cac to an chi
    If Val(LoaiTk1) <> 64 And Val(LoaiTk1) <> 65 And Val(LoaiTk1) <> 66 And Val(LoaiTk1) <> 67 And Val(LoaiTk1) <> 68 And Val(LoaiTk1) <> 91 Then
        If getTTDLT = False Then
            If MessageBox("0141", msYesNo, miQuestion) = mrNo Then
                Exit Function
            End If
        End If
    End If

    ' end
    If Trim(GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "Class")) <> vbNullString Then
        Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "Class"))
        Set objTaxBusiness.fps = fpSpread1
        objTaxBusiness.strPhongXuLy = strMaPhongXuLy
        objTaxBusiness.strNgayNhanToKhai = strNgayNhanToKhai
        objTaxBusiness.strNguoiSuDung = strUserID
        
        objTaxBusiness.strMST = Replace(strMST, "%", "")
        objTaxBusiness.strTenGoi = strTenGoi
        objTaxBusiness.strDchi = strDchi
        objTaxBusiness.strNganh = strNganh
        objTaxBusiness.strMaBPQL = strMaBPQL
        objTaxBusiness.strDThoai = strDThoai
        objTaxBusiness.strFax = strFax
        objTaxBusiness.spathVat = spathVat
        objTaxBusiness.hannop = hannop
        objTaxBusiness.strTenBpql = strTenBpql
        ' vttoan: thong tin dai ly thue
        LoaiTk1 = Mid(strData, 4, 2)

        If LoaiTk1 = "01" Or LoaiTk1 = "02" Or LoaiTk1 = "04" Or LoaiTk1 = "11" Or LoaiTk1 = "06" Or LoaiTk1 = "05" Or LoaiTk1 = "15" Or LoaiTk1 = "16" Or LoaiTk1 = "50" Or LoaiTk1 = "36" Or LoaiTk1 = "70" Or LoaiTk1 = "72" Or LoaiTk1 = "86" Or LoaiTk1 = "87" Or LoaiTk1 = "74" Or LoaiTk1 = "75" Or LoaiTk1 = "03" Or LoaiTk1 = "71" Then
            objTaxBusiness.strMST_DLT = Replace(strMST_DLT, "%", "")
            objTaxBusiness.strTen_DLT = strTen_DLT
            objTaxBusiness.strDchi_DLT = strDchi_DLT
            '        objTaxBusiness.strQHuyen_DLT = strQHuyen_DLT
            '        objTaxBusiness.strTTPho_DLT = strTTPho_DLT
            objTaxBusiness.strDthoai_DLT = strDthoai_DLT
            objTaxBusiness.strFax_DLT = strFax_DLT
            objTaxBusiness.strMail_DLT = strMail_DLT
            objTaxBusiness.strSoHD_DLT = strSoHD_DLT
            objTaxBusiness.strNgayHD_DLT = strNgayHD_DLT
            objTaxBusiness.checkDLTTonTai = getTTDLT
        End If

        'checkDLTTonTai
        'end
        
        ' An chi
        LoaiTk1 = Mid(strData, 4, 2)

        If (Val(LoaiTk1) >= 64 And Val(LoaiTk1) <= 68) Or Val(LoaiTk1) = 91 Then
            objTaxBusiness.strSoTTTKhai = getSoTTTK_AC(changeMaToKhai(LoaiTk1), arrStrHeaderData, strData)
            objTaxBusiness.isTKTonTai = isTonTaiAC
            objTaxBusiness.strMaBPQL = strMaPhongQuanLy
            objTaxBusiness.strTenBpql = strTenPhongQuanLy
            objTaxBusiness.strNguoiSuDung = strUserID
        End If
        
        'dhdang
        'ngay 31/08/2010
        'lay thog tin sang form in BB nop cham
        MST_PRINT = strMST
        NNT_PRINT = strTenGoi
        DIACHI_PRINT = strDchi
        'dhdang sua lay loai hs
        'ngay 08/11/10
        LOAihs_PRINT = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Caption")

        If Not objTaxBusiness.Prepared1 Then Exit Function
    End If
            
    SetupData fpSpread1
    '***********************************
    'BacLT added
    'Date: 21/05/06
    'Lay thong tin phong xu ly.
    Set rsPXL = GetPhongXuLy(strPhongXuLy, blnConnected)
    
    If Not blnConnected Then
        Exit Function
    End If

    If rsPXL Is Nothing Then
        DisplayMessage "0077", msOKOnly, miCriticalError
        Exit Function
    End If

    '***********************************
    If Not objTaxBusiness Is Nothing Then
        If Not objTaxBusiness.Prepared2(rsPXL) Then Exit Function
    End If
    
    'Load co quan thue KHBS
    If InStr(tt156, LoaiTk1) > 0 Then

        With fpSpread1
            Dim CQT_CAPCUC    As Variant
            Dim CQT_HOANTHUE  As Variant
            Dim tCQT_CAPCUC   As String
            Dim tCQT_HOANTHUE As String

            If TAX_Utilities_Svr_New.NodeValidity.hasChildNodes Then
                If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 1), "ID") = "KHBS" Then
                    If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 1), "Active") = "1" Then
                        .Sheet = .SheetCount - 1
                        .GetText .ColLetterToNumber("BI"), .MaxRows - 15, CQT_CAPCUC
                        .GetText .ColLetterToNumber("BI"), .MaxRows - 13, CQT_HOANTHUE
                        GetTenCQT CQT_CAPCUC, tCQT_CAPCUC
                        GetTenCQT CQT_HOANTHUE, tCQT_HOANTHUE
                        .Col = .ColLetterToNumber("BE")

                        If tCQT_CAPCUC <> vbNullString Then
                            .Row = .MaxRows - 15
                            .Text = tCQT_CAPCUC

                        End If

                        If tCQT_HOANTHUE <> vbNullString Then
                            .Row = .MaxRows - 13
                            .Text = tCQT_HOANTHUE
 
                        End If
                    End If
                End If
            End If
    
        End With

    End If

    clsDAO.Disconnect
    'Setup header data
    'SetupHeaderData rsHeaderData
    SetupHeaderData arrStrHeaderData

    If Not objTaxBusiness Is Nothing Then
        If Not objTaxBusiness.Prepared3 Then Exit Function
    End If
    
    fpSpread1.EventEnabled(EventAllEvents) = True
    cmdClear.Enabled = True
    cmdSave.Enabled = True
    cmd_insert.Enabled = True
    cmdViewNow.Enabled = False
    fpSpread1.Visible = True
    
    lblLabelVersion.Left = 3630
    lblVersion.Left = 8520
    
    If CLng(Replace$(strTaxReportVersion, ".", "")) < CLng(Replace$(HTKK_LAST_VERSION, ".", "")) Then
        Beep 600, 500
        lblWarning.Visible = True
    Else
        lblLabelVersion.Left = lblLabelVersion.Left + lblWarning.Width
        lblVersion.Left = lblVersion.Left + lblWarning.Width
    End If

    lblLabelVersion.Visible = True
    lblVersion.Visible = True
    
    If frmSystem.chkSaveQuestion = True Then
        cmdClear.SetFocus
    Else
        cmdSave.SetFocus
    End If
    
    mOnLoad = False
    LoadForm = True
    
    frmSystem.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "LoadForm", Err.Number, Err.Description
End Function

'****************************
'Description: IsCompleteData function check whether barcode data is complete
'Author:ThanhDX
'Date:19/11/2005
'Input:
'Output:
'       strData: complete data string.
'Return: true if the string is complete data
'        false if otherwise
'****************************
Private Function IsCompleteData(ByRef strData As String) As Boolean
    Dim blnReturn As Boolean
    Dim intCtrl As Integer, intCount As Integer
    Dim strTemp As String

On Error GoTo ErrHandle
    blnReturn = True
    strTemp = arrStrElements(0)
    ' Check Version
    If Left$(strTemp, 1) = "0" Then
    'Version 1.1.0 and 1.0.0
    Else
    'Version 1.2.0 and later
        'Remove 6 character of printting session
        strTemp = Mid$(strTemp, 1, Len(strTemp) - 6)
    End If
    '*********************************
    

    
    For intCtrl = 1 To UBound(arrStrElements())
        If Trim(arrStrElements(intCtrl)) = vbNullString Then
            blnReturn = False
        Else
            strTemp = strTemp & arrStrElements(intCtrl)
            intCount = intCount + 1
        End If
    Next intCtrl
    
    'Get all of data string
    If blnReturn Then
        strData = strTemp
    End If
    
    If intCount > ProgressBar1.Value Then
        ProgressBar1.Value = ProgressBar1.Value + 1
        lblFilePath.caption = ProgressBar1.Value & " / " & ProgressBar1.max
    End If
    
    IsCompleteData = blnReturn
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "IsCompleteData", Err.Number, Err.Description
End Function

'****************************
'Description: IsDifferent function compare two string
'Author:ThanhDX
'Date:23/11/2005
'Input:
'       strValue1: The first string.
'       strValue2: The second string.
'Output:
'Return: 0 if two string is the same
'        1 if two string not is the same
'****************************
Private Function IsDifferent(ByVal strValue1 As String, ByVal strValue2 As String) As Integer
    Dim intReturn As Integer

On Error GoTo ErrHandle
    If strValue1 = strValue2 Then Exit Function
    
    'If Mid$(strValue1, 3) <> Mid$(strValue2, 3) Or Left$(strValue1, 2) <> Left$(strValue2, 2) Then
        intReturn = 1 'Another tax report
        'GoTo exitFunction
    'End If

'exitFunction:
    IsDifferent = intReturn
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "IsDifferent", Err.Number, Err.Description
End Function

'****************************
'Description: GetSheetDatas function divide data string into sheet datas.
'Author:ThanhDX
'Date:23/11/2005
'Input:strBarcodeData: Data string.
'Output:
'Return: array of data sheets.
'****************************
Private Function GetSheetDatas(ByVal strBarcodeData As String) As String()
    Dim arrStrData() As String ', strSheetId As String , strTemp As String
    Dim intIndex As Integer, intLoc1 As Long, intLoc2 As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    
On Error GoTo ErrHandle
    For Each xmlNode In TAX_Utilities_Svr_New.NodeValidity.childNodes
        SetAttribute xmlNode, "Active", "0"
    Next
    
    ReDim arrStrData(0)
    
'    Do
'        strSheetId = Mid$(strBarcodeData, 2, 3)
'        strTemp = strBarcodeData
'        For Each xmlNode In TAX_Utilities_Svr_New.NodeValidity.childNodes
'            If GetAttribute(xmlNode, "ID") = Right$(strSheetId, 2) Then
'                SetAttribute xmlNode, "Active", "1"
'                intLoc = InStr(1, strBarcodeData, "</" & strSheetId & ">")
'
'                intIndex = intIndex + 1
'                ReDim Preserve arrStrData(intIndex)
'                arrStrData(intIndex) = Mid$(strBarcodeData, 1, intLoc + 5)
'                strBarcodeData = Right$(strBarcodeData, Len(strBarcodeData) - Len(arrStrData(intIndex)))
'                Exit For
'            ElseIf GetAttribute(xmlNode, "Active") = "0" Then
'                intIndex = intIndex + 1
'                ReDim Preserve arrStrData(intIndex)
'            End If
'        Next
'
'        If strTemp = strBarcodeData Then
'            blnErr = True
'            Exit Do
'        End If
'    Loop Until strBarcodeData = vbNullString
    
    For Each xmlNode In TAX_Utilities_Svr_New.NodeValidity.childNodes
        Dim i As Integer
        i = Len(GetAttribute(xmlNode, "ID"))
        intLoc1 = InStr(1, strBarcodeData, "<S" & GetAttribute(xmlNode, "ID") & ">")
        
        If intLoc1 = 0 Then
            intIndex = intIndex + 1
            ReDim Preserve arrStrData(intIndex)
        Else
            intLoc2 = InStr(1, strBarcodeData, "</S" & GetAttribute(xmlNode, "ID") & ">")
            If intLoc2 > intLoc1 Then
                SetAttribute xmlNode, "Active", "1"
                intIndex = intIndex + 1
                ReDim Preserve arrStrData(intIndex)
                arrStrData(intIndex) = Mid$(strBarcodeData, intLoc1, intLoc2 + i + 3)
                strBarcodeData = Replace(strBarcodeData, arrStrData(intIndex), "")
            End If
        End If
    Next
    
    If strBarcodeData = "" Then
        If UBound(arrStrData) < TAX_Utilities_Svr_New.NodeValidity.childNodes.length Then
            ReDim Preserve arrStrData(TAX_Utilities_Svr_New.NodeValidity.childNodes.length)
        End If
    Else
        ReDim arrStrData(0)
    End If
    
    GetSheetDatas = arrStrData()
    Set xmlNode = Nothing

    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetSheetDatas", Err.Number, Err.Description
End Function

'****************************
'Description: GetTaxInfo function get data from DB
'Author:ThanhDX
'Date:23/11/2005
'Input:
'       strTaxIDString: Id string
'Output:
'Return: a Recordset contain query data.
'****************************
Private Function GetTaxInfo(ByVal strTaxIDString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'
'    ' Get SQL statement from DOM
'    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlMST")
'    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
'    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
'
'    Set rsReturn = clsDAO.Execute(strSQL)
'
'    Set GetTaxInfo = rsReturn
'
    Set rsReturn = Nothing
    
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetTaxInfo", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function
Private Function GetPhongXuLy(ByVal strPXLString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & "\DTNT\"
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
    End If
'
'
'    ' Get SQL statement from DOM
'    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlPhongXuLy")
'
'    '*************************************
'    'ThanhDX added
'    'Date: 30/05/06
'    strSQL = Replace$(strSQL, "MA_CQT", strTaxOfficeId)
'    '*************************************
''    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
''    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
'
   'htphuong sua lay dsach BPQL
   ' strSQL = "SELECT mabpql, tengoi FROM dmbpql WHERE len(alltrim(mabpql)) = 2"
   
    strSQL = "SELECT mabpql, tengoi FROM dmbpql WHERE len(alltrim(macaptren)) =0"
    Set rsReturn = clsDAO.Execute(strSQL)

    Set GetPhongXuLy = rsReturn
    
    Set rsReturn = Nothing
    
    'Connect DB success
    blnSuccess = True
    'clsDAO.Disconnect
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    clsDAO.Disconnect
    SaveErrorLog Me.Name, "GetPXL", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

Private Sub SetupHeaderData(arrStrHeaderData() As String)
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    
On Error GoTo ErrHandle
        fpSpread1.Sheet = lCtrl + 1
        For lIndex = 0 To UBound(arrStrHeaderData) 'TAX_Utilities_Svr_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes.length
            'If lIndex < UBound(arrStrHeaderData) Then
                If Not arrStrHeaderData(lIndex) = vbNullString Then
                    SetAttribute TAX_Utilities_Svr_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex) _
                        , "Value", TAX_Utilities_Svr_New.Convert(arrStrHeaderData(lIndex), TCVN, UNICODE)
                    ParserCellID fpSpread1, GetAttribute(TAX_Utilities_Svr_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex), "CellID"), lCol, lRow
                    fpSpread1.SetText lCol, lRow, TAX_Utilities_Svr_New.Convert(arrStrHeaderData(lIndex), TCVN, UNICODE)
                    fpSpread1.RowHeight(lRow) = fpSpread1.MaxTextRowHeight(lRow)
                End If
            'Else
                'Exit For
            'End If
        Next lIndex
        
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetupHeaderData", Err.Number, Err.Description
End Sub

Function GenerateSQL_Details(xmlDOMdata As MSXML.DOMDocument, strSQL_DTL As String, vHdrID As Variant, lPos As Long) As String
    Dim xmlListSection As MSXML.IXMLDOMNodeList
    Dim xmlNodeSection As MSXML.IXMLDOMNode
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim lPosition As Long, strCondition As String
    Dim i As Long, j As Long, strLoaiDL As String
    
On Error GoTo ErrHandle
    iRowID = 0
    Set xmlListSection = xmlDOMdata.getElementsByTagName("Section")
    For Each xmlNodeSection In xmlListSection
        If Trim(xmlNodeSection.Attributes.getNamedItem("Dynamic").nodeValue) = "1" Then
            For i = 0 To xmlNodeSection.childNodes.length - 1
                iRowID = iRowID + 1
                For j = 0 To xmlNodeSection.childNodes(i).childNodes.length - 1
                    Set xmlAttribute = xmlDOMdata.createAttribute("RowID")
                    xmlAttribute.Value = iRowID
                    Set xmlNode = xmlNodeSection.childNodes(i).childNodes(j).Attributes.setNamedItem(xmlAttribute)
                    Set xmlAttribute = Nothing
                Next
            Next
        End If
    Next
    
    strLoaiDL = Trim(TAX_Utilities_Svr_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue) & Trim(TAX_Utilities_Svr_New.NodeValidity.childNodes(lPos).Attributes.getNamedItem("ID").nodeValue)
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    If xmlList.length > 0 Then GenerateSQL_Details = ""
    For Each xmlNode In xmlList
        If Not xmlNode.Attributes.getNamedItem("MCT") Is Nothing Then
             If Trim(xmlNode.Attributes.getNamedItem("MCT").nodeValue) <> "" Then
                strSQL = strSQL_DTL
                strSQL = strSQL & "'" & vHdrID & "',"
                strSQL = strSQL & "'" & strLoaiDL & "',"
                strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("MCT").nodeValue & "',"
                strSQL = strSQL & "'" & Trim(Replace(TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Value").nodeValue, UNICODE, TCVN), "'", "''")) & "',"
                If Not xmlNode.Attributes.getNamedItem("RowID") Is Nothing Then
'**********************************
'ThanhDX modified
'Date: 26/07/2006
                    'strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("RowID").nodeValue & "');"
                    strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("RowID").nodeValue & "')"
'**********************************
                Else
'**********************************
'ThanhDX modified
'Date: 26/07/2006
                    'strSQL = strSQL & "''); "
                    strSQL = strSQL & "'') "
                End If
'ThanhDX inserted
'Date: 26/07/2006
     '           clsDAO.Execute strSQL
'**********************************
                GenerateSQL_Details = GenerateSQL_Details & vbCrLf & strSQL
             End If
        End If
    Next
    If Trim(GenerateSQL_Details) <> "" Then GenerateSQL_Details = GenerateSQL_Details & vbCrLf
    Set xmlDOMdata = Nothing
    Set xmlList = Nothing
    Set xmlListSection = Nothing
    Exit Function

ErrHandle:
    SaveErrorLog Me.Name, "GenerateSQL_Details", Err.Number, Err.Description
End Function

''' CheckValidData description
''' Check all formula in last sheet, if error put the notetext into cellnode
''' No parameter
''' Return True if no error checking
''' Return False if one or more error occur
Private Function CheckValidData() As Boolean
    
    Dim i As Long
    Dim strCellString As String
    
    Dim vFunction As Variant, vCell As Variant
    Dim vMsg As Variant, vWarning As Variant
    Dim vOrder As Variant, vFormulaFunc As Variant
    Dim cOrder As New Collection
    
On Error GoTo ErrHandle

    CheckValidData = True
    
    With fpSpread1
        .ReDraw = False
        If .SheetCount = 1 Then Exit Function
        .Sheet = mHeaderSheet
        
        For i = 12 To .MaxRows
            .Sheet = mHeaderSheet
            .Col = .ColLetterToNumber("B")
            .Row = i
            If .Formula <> vbNullString Then
                .Col = .Col + 1 'Column B
                strCellString = .Formula
                If Trim(strCellString) <> vbNullString Then SetCellNote strCellString, .BackColor, ""
            End If
        Next
        
        'set error note for cell
        If Not objTaxBusiness Is Nothing Then
            CheckValidData = objTaxBusiness.CheckValidData
        End If
        
        .Sheet = mHeaderSheet
        For i = 12 To .MaxRows
            .Sheet = mHeaderSheet
            .Col = 2
            .Row = i
            vFormulaFunc = .Formula
            If Trim(.Text) <> "" Then
                .GetText .ColLetterToNumber("B"), i, vFunction
                .GetText .ColLetterToNumber("E"), i, vMsg
                .GetText .ColLetterToNumber("S"), i, vWarning
                .GetText .ColLetterToNumber("T"), i, vOrder
                .Col = .Col + 1
                vCell = .Formula
                If vFormulaFunc <> vbNullString Then
                    If Val(vFunction) <> 1 Then
                        SetCellNote vCell, .BackColor, "> " & vMsg
                        If Trim(vCell) <> "" Then cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                        If UCase(Trim(vWarning)) = "Y" Then CheckValidData = False
                    End If
                Else 'Dynamic
                    If Val(vFunction) <> 1 Then
                        If Trim(vCell) <> "" Then cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                        If UCase(Trim(vWarning)) = "Y" Then CheckValidData = False
                    End If
                End If
            End If
        Next
        
        'focus on the first error cell
        Dim min As Integer, X As Long, strCell As String
        Dim lSheet As Long, lCol As Long, lRow As Long
        
        
        If cOrder.Count > 0 Then
            min = Val(Left(cOrder(1), InStr(cOrder(1), "[]")))
            strCell = Right(cOrder(1), Len(cOrder(1)) - InStr(cOrder(1), "[]") - 1)
            For i = 2 To cOrder.Count
                X = Val(Left(cOrder(i), InStr(cOrder(i), "[]")))
                If min >= X Then
                    min = X
                    strCell = Right(cOrder(i), Len(cOrder(i)) - InStr(cOrder(i), "[]") - 1)
                End If
            Next
            'focus cell here
            getCellPosition strCell, lSheet, lCol, lRow
            .SetFocus
            .ActiveSheet = lSheet
            .SetActiveCell lCol, lRow
        End If
        
        .ReDraw = True
    End With
    Exit Function
    
ErrHandle:
    SaveErrorLog Me.Name, "CheckValidData", Err.Number, Err.Description
End Function


''' get Sheet, Col, Row from Cell Formula
'''Parameter: Cell Formula string
'''Parameter: sheet integer
'''parameter: Col integer
'''parameter: Row integer
Private Sub getCellPosition(pCellString As String, lSheet As Long, lCol As Long, lRow As Long)
        
    Dim lAnchor As Long
    Dim lSheetName As String, lCellString As String, lStringTemp As String
    Dim i As Long
    
On Error GoTo ErrHandle
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For i = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, i)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, i)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For i = 1 To .SheetCount
            .Sheet = i
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
                lSheet = i
                Exit For
            End If
        Next
    End With
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "getCellPosition", Err.Number, Err.Description
End Sub

''' SetCellNote description
''' Set CellNote for error cell
''' Parser pCellString (containt sheetname and cellID)
''' Parameter1 pCellString  : containt sheetname and cellID
''' Parameter2 pNoteText    : the string input into cellnote
Private Sub SetCellNote(ByVal pCellString As String, ByVal lNoErrColor As Long, ByVal pNoteText As String)
    
    Dim lAnchor As Long
    Dim lSheetName As String, lCellString As String, lStringTemp As String
    Dim lCol As Long, lRow As Long, i As Long
    
On Error GoTo ErrHandle
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For i = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, i)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, i)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For i = 1 To .SheetCount
            .Sheet = i
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
                .Col = lCol
                .Row = lRow
                
                If Trim(pNoteText) = "" Then
                    .CellNote = ""
                ElseIf Trim(.CellNote) = "" Then
                    .CellNote = pNoteText
                Else
                    .CellNote = .CellNote & vbCrLf & pNoteText
                End If
                If Trim(.CellNote) <> vbNullString Then
                    .BackColor = &HC0C0FF   'VB 'vbRed
                Else
                    .BackColor = lNoErrColor
                End If
                Exit For
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "SetCellNote", Err.Number, Err.Description
End Sub

'****************************
'Description: StartReceiveForm initialize intput data screen.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'Output:
'Return:
'****************************
Private Sub StartReceiveForm()
    
On Error GoTo ErrHandle
    fpSpread1.Visible = False
    
    lblLabelVersion.Visible = False
    lblVersion.Visible = False
    lblWarning.Visible = False
    
    strTaxReportInfo = ""
    TAX_Utilities_Svr_New.xmlDataReDim (0)
    cmdClear.Enabled = False
    cmdSave.Enabled = False
    cmd_insert.Enabled = False
    cmdViewNow.Enabled = True
    Set objTaxBusiness = Nothing
    frmSystem.MousePointer = vbDefault
    Me.MousePointer = vbDefault
        
    If blnReceiveByBarcode Then
        lblFile.Visible = False
        lblBarcode.Visible = True
        lblFilePath.caption = "0/0"
        lblLoading.Visible = True
        lblConnecting.Visible = False
        lblExit.Visible = False
        ProgressBar1.Value = 0
        ReDim arrStrElements(0)
        If Not blnOnLoadEvent Then cmdExit.SetFocus
        blnOnLoadEvent = False
    Else
        If UBound(arrStrElements) > 0 Then
            lblFile.Visible = True
            lblBarcode.Visible = False
            lblFilePath.caption = arrStrElements(UBound(arrStrElements))
            lblLoading.Visible = False
            lblConnecting.Visible = True
            lblExit.Visible = False
            
            LoadFormByFileName
        ElseIf Not blnOnLoadEvent Then
            Unload Me
            frmTreeviewMenu.Show
        Else
            lblLoading.Visible = False
            lblConnecting.Visible = False
            lblExit.Visible = True
        End If
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "StartReceiveForm", Err.Number, Err.Description
End Sub

'****************************
'Description: TrimString function cut not valid characters
'             at the begin of string.
'Author:ThanhDX
'Date:27/11/2005
'Input: strValue: data string.
'OutPut:
'Return: Data string cut (if it exist not valid characters)
'****************************
Private Function TrimString(ByVal strValue As String) As String
    Dim lCtrl As Long, strNumber As String
    
On Error GoTo ErrHandle
    strNumber = "0123456789"
    For lCtrl = 1 To Len(strValue)
        If InStr(1, strNumber, Mid$(strValue, lCtrl, 1)) <> 0 Then
            TrimString = Mid$(strValue, lCtrl)
            Exit Function
        End If
    Next lCtrl
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "TrimString", Err.Number, Err.Description
End Function

'****************************
'Description: ShowFormReceiveFromBarcode procedure initialize
'             form which is waiting barcode reader.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub ShowFormReceiveFromBarcode()
On Error GoTo ErrHandle
    StartBarcodeReader
    StartReceiveForm
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "ShowFormReceiveFromBarcode", Err.Number, Err.Description
End Sub

'****************************
'Description: ShowFormReceiveFromFile procedure initialize
'             form which get content form file.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub ShowFormReceiveFromFile()
    ProgressBar1.max = UBound(arrStrElements)
    StartReceiveForm
End Sub

'****************************
'Description: LoadFormByFileName procedure load data from file
'             then call load form.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub LoadFormByFileName()
    Dim intUbound As Integer
    Dim strData As String, strFileName As String
    
On Error GoTo ErrHandle
    intUbound = UBound(arrStrElements)
    If intUbound = 0 Then
        'Unload Me
        Exit Sub
    End If
    
    strFileName = arrStrElements(intUbound)
    ReDim Preserve arrStrElements(intUbound - 1)
    strData = GetDataFormFile(strFileName)
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If Not LoadForm(strData) Then
        'If UBound(arrStrElements) > 0 Then
            StartReceiveForm
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "LoadFormByFileName", Err.Number, Err.Description
End Sub

Public Sub SetReceiveByBarcode(ByVal blnValue As Boolean)
    blnReceiveByBarcode = blnValue
End Sub

Public Sub SetArrayElements(arrStrValue() As String)
    arrStrElements = arrStrValue
End Sub

'****************************
'Description: GetDataFromFile function get content in file by name.
'Author:ThanhDX
'Date:25/11/2005
'Input:
'   strFileName: name of file contain data
'OutPut:
'Return: Data string contained in the file
'****************************
Private Function GetDataFormFile(ByVal strFileName As String) As String
    Dim fso As New FileSystemObject
    Dim tstFile As TextStream
    
On Error GoTo ErrHandle
    Set tstFile = fso.OpenTextFile(strFileName, ForReading)
    While Not tstFile.AtEndOfStream
        GetDataFormFile = GetDataFormFile & tstFile.ReadLine
    Wend
    GetDataFormFile = TAX_Utilities_Svr_New.Convert(GetDataFormFile, TCVN, UNICODE)
    tstFile.Close
    Set fso = Nothing
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetDataFromFile", Err.Number, Err.Description
End Function

''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Private Function UpdateCell(ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String) As Boolean
    On Error GoTo ErrHandle
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fpSpread1, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_Svr_New.Data(fpSpread1.ActiveSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow))
    
    If GetAttribute(xmlNodeCell, "Value") <> pValue Then
        SetAttribute xmlNodeCell, "Value", pValue
        UpdateCell = True
    End If
    
    Set xmlNodeCell = Nothing
    
    Exit Function
    
ErrHandle:
    SaveErrorLog Me.Name, "UpdateCell", Err.Number, Err.Description
End Function

'****************************
'Description: MessageBox function stop barcode reader then call message box
'             and start barcode reader (If received data method  is bacode reader)
'Author:ThanhDX
'Date:24/11/2005
'Input:
'   strMsgId: Message Id
'   intMsgStyle: Style of message
'   intMsgIcon: Type of icon message
'Output:
'Return:Action user
'****************************
Private Function MessageBox(strMsgId As String, intMsgStyle As MsgBoxStyle, intMsgIcon As MsgBoxIcon) As MsgBoxResult
    Dim intReturn As Integer
    
On Error GoTo ErrHandle
    If blnReceiveByBarcode Then StopBarcodeReader
    
    MessageBox = DisplayMessage(strMsgId, intMsgStyle, intMsgIcon)
    
    If blnReceiveByBarcode Then StartBarcodeReader
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "MessageBox", Err.Number, Err.Description
End Function

''' LoadInitFiles description
''' Set max len for string type cell
''' Set min value for numeric type cell
''' Set max value for numeric type cell
''' Call after load template
''' No parameter
Private Sub LoadInitFiles()
    On Error GoTo ErrorHandle
    Dim i As Long, lCol As Long, lRow As Long
    Dim xmlNodeListIni As MSXML.IXMLDOMNodeList
    Dim xmlNodeIni As MSXML.IXMLDOMNode
    
    For i = 0 To fpSpread1.SheetCount - 2
        ReDim Preserve xmlDocumentInit(i)
        Set xmlDocumentInit(i) = New MSXML.DOMDocument
        xmlDocumentInit(i).Load GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(i), "Ini"))
        Set xmlNodeListIni = xmlDocumentInit(i).getElementsByTagName("Cell")
        For Each xmlNodeIni In xmlNodeListIni
            fpSpread1.Sheet = i + 1
            ParserCellID fpSpread1, GetAttribute(xmlNodeIni, "CellID"), lCol, lRow
            fpSpread1.Col = lCol
            fpSpread1.Row = lRow
            If Val(GetAttribute(xmlNodeIni, "MaxLen")) <> 0 Then
                fpSpread1.TypeMaxEditLen = Val(GetAttribute(xmlNodeIni, "MaxLen"))
            End If
            If fpSpread1.CellType = CellTypeNumber Then
                fpSpread1.TypeNumberMin = Val(GetAttribute(xmlNodeIni, "MinValue"))
                fpSpread1.TypeNumberMax = Val(GetAttribute(xmlNodeIni, "MaxValue"))
            End If
            fpSpread1.CellTag = GetAttribute(xmlNodeIni, "HelpContexID") & fpSpread1.CellTag
        Next
    Next
    
    Set xmlNodeIni = Nothing
    Set xmlNodeListIni = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "LoadInitFiles", Err.Number, Err.Description
End Sub
'
''****************************************************
''Description: GetDynRowCount function get count of interface rows in
''             one Cells node.
''Author: ThanhDX
''Date:14/12/2006
''Input:
''       pGrid: fpSpread
''       xmlNodeCells: Cells node in dynamic section
''       lReportRows: Count of report rows in Cells node
''       lMinRow: Min row in Cells node
''       lMaxRow: Max row in Cells node
''****************************************************
'Public Function GetDynRowCount(pGrid As fpSpread, xmlNodeCells As MSXML.IXMLDOMNode, Optional ByRef lMinRow As Long, Optional lMaxRow As Long)
'    Dim xmlNodeCell As MSXML.IXMLDOMNode
'    Dim lRow As Long, lCol As Long
'
'    lMinRow = 100000
'    lMaxRow = 0
'
'    If Not xmlNodeCells Is Nothing Then
'        For Each xmlNodeCell In xmlNodeCells.childNodes
'            'Get CellID
'            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
'
'            'Get max row
'            If lRow > lMaxRow Then
'                lMaxRow = lRow
'            End If
'
'            'Get min row
'            If lRow < lMinRow Then
'                lMinRow = lRow
'            End If
'        Next
'
'        GetDynRowCount = lMaxRow - lMinRow + 1
'    End If
'End Function

Private Function GetThongTinTep(ByVal strID As String, arrStrHeaderData() As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String, strMaTkhaiQLT As String
    Dim strPrefixMaTep As String, strMatep As String
    Dim strSTT As String
    
    On Error GoTo ErrHandle
    
   ' lngIndex = UBound(arrStrHeaderData)
    
    On Error GoTo ConnectErrHandle
    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
    
    'Lay ma to khai theo QLT
'    strSQL = "Select hso.loai_hoso " & _
'            "From rcv_map_tkhai tkhai," & _
'            "qlt_map_hoso_tkhai hso " & _
'            "Where (tkhai.nhom_hso = hso.nhom) " & _
'            "And (tkhai.ma_tkhai_qlt = hso.loai_tkhai) " & _
'            "And (tkhai.ma_tkhai = '" & strID & "')"
'
'    Set rsResult = clsDAO.Execute(strSQL)
    strMaTkhaiQLT = "04" ' rsResult.Fields(0).Value
    
    'La^'y chuo^~i tie^`n to^' cu?a ma~ te^.p
'    strSQL = "Select To_Char(Sysdate,'RRMM')||'" & strMaTkhaiQLT & _
'            "' From Dual"
'
'    Set rsResult = clsDAO.Execute(strSQL)
    strPrefixMaTep = "a" ' rsResult.Fields(0).Value
    
    'Lay so thu tu lon nhat cua tep (hau to)
'    strSQL = "Select nvl(max(To_Number(Substr(So_Hieu_tep,8,3))),1) " & _
'            "From rcv_tkhai_hdr " & _
'            "Where So_Hieu_Tep Like '" & strPrefixMaTep & "' || '%'"
'
'    Set rsResult = clsDAO.Execute(strSQL)
    strMatep = strPrefixMaTep & "-" & "ab" 'rsResult.Fields(0).Value
    
'    'Lay so to khai lon nhat trong tep tim dc
'    strSQL = "Select nvl(max(so_tt_tk),0) + 1 " & _
'            "From rcv_tkhai_hdr " & _
'            "Where So_Hieu_tep = '" & strMatep & "'"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strSTT = rsResult.Fields(0).Value
'
'    If CInt(strSTT) > 50 Or CInt(strSTT) = 1 Then
'        'Dong tep, sinh tep moi
'        'Lay ma tep tu DB QHS
'        Dim intSuffixMaTep As Integer
'        strSQL = "Select nvl(max(To_Number(Substr(So_Hieu,8,3))),0)+1 " & _
'                "From Qhs_Tep_Hoso " & _
'                "Where So_Hieu Like '" & strPrefixMaTep & "' || '%'"
'
'        Set rsResult = clsDAO.Execute(strSQL)
'        intSuffixMaTep = CInt(rsResult.Fields(0).Value)
'
'        'Kie^?m tra te^.p du?a va`o da~ co' du?~ lie^.u hay chu?a.
'        strSQL = "Select so_hoso " & _
'                "From Qhs_Tep_Hoso " & _
'                "Where (so_hieu = '" & strPrefixMaTep & "-" & (intSuffixMaTep - 1) & "')"
'
'        Set rsResult = clsDAO.Execute(strSQL)
'
'        If Not rsResult Is Nothing Then
'            If rsResult.Fields(0).Value <> "0" Or IsNull(rsResult.Fields(0)) Then
'CreateFile:
'                strMatep = strPrefixMaTep & "-" & intSuffixMaTep
'                strSTT = "1"
'
'                'Insert tep moi vao QHS
'                strSQL = "Insert Into Qhs_Tep_Hoso (So_Hieu, Dhs_Ma, " & _
'                        "Kykk_Tu_Ngay, Kykk_Den_Ngay, Ngay_Tao, So_Hoso)" & _
'                        "Values ('" & strMatep & "'," & _
'                        "'" & strMaTkhaiQLT & "'," & _
'                        "To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')," & _
'                        "To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')," & _
'                        "Sysdate,0)"
'                clsDAO.Execute (strSQL)
'                strSQL = "commit"
'                clsDAO.Execute (strSQL)
'            Else
'                strMatep = strPrefixMaTep & "-" & (intSuffixMaTep - 1)
'                strSTT = "1"
'            End If
'        Else
'            GoTo CreateFile
'        End If
'    End If
'
'    On Error GoTo ErrHandle
'
'    'Ghep ma so tep vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 1)
    arrStrHeaderData(lngIndex + 1) = strMatep
'
'    'Ghep so thu tu to khai vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 2)
    arrStrHeaderData(lngIndex + 2) = "" 'strSTT
    
    Set rsResult = Nothing
    GetThongTinTep = True
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetThongTinTep", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "GetThongTinTep", Err.Number, Err.Description
End Function

Private Function GetHeaderData(ByVal rsTaxInfor As ADODB.Recordset, arrStrHeaderData() As String) As Boolean
    Dim arrStrData() As String
    Dim lCtrl As Long
       
    On Error GoTo ErrHandle
    
'    If rsTaxInfor Is Nothing Then
'        Exit Function
'    End If
'
'    If rsTaxInfor.Fields.Count = 0 Then
'        Exit Function
'    End If
'
'    For lCtrl = 0 To rsTaxInfor.Fields.Count - 2
'        ReDim Preserve arrStrData(lCtrl)
'        If Not IsNull(rsTaxInfor.Fields(lCtrl + 1).Value) Then
'            arrStrData(lCtrl) = rsTaxInfor.Fields(lCtrl + 1).Value
'        End If
'    Next lCtrl
'
'    'Loai bo gia tri Ngay bat dau nam TC va Ngay bat dau KD
'    ReDim Preserve arrStrData(UBound(arrStrData) - 2)
'
'    arrStrHeaderData = arrStrData
    GetHeaderData = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetHeaderData", Err.Number, Err.Description
End Function
'dhdang them ham Prepare data de insert QLT

Private Function Prepare_QLT() As String
    Dim sSQL           As String
    Dim sSQLCol        As String
    Dim sSQLVal        As String
    Dim rs             As ADODB.Recordset
    Dim MATKHAI        As Variant
    Dim KYLBO          As Variant
    Dim NGNOP          As Variant
    Dim NGNHAP         As Variant
    Dim KYKKHAI        As Variant
    Dim maDTNT         As Variant
   
    Dim DAGHI          As Variant
    Dim LOAITIEN       As Variant
    Dim MAVACH         As Variant
    Dim SHTEP          As Variant
    Dim HANNOP2        As Variant
    Dim THUETKY        As Variant
    Dim THUETKY2       As Variant
    Dim MAMUC          As Variant
    Dim MATM           As Variant
    Dim CTHUC          As Variant
    Dim BSUNG          As Variant
    Dim LANBS          As Variant
    Dim TRICH_YEU      As String
    Dim vKYLBO         As Variant
    Dim CHKGIAHAN      As Variant
    Dim bln            As Boolean
    'dhdang them bien
    'Dim DHS_MA As Variant
    Dim PHONG_XL       As Variant
    Dim PHONG_XL_X     As Variant
    Dim PHONG_XL_Y     As Variant
    Dim SO_HOSO        As Variant
    Dim NGAY_XL        As Date
    Dim NGAY_HEN       As Date
    Dim NGAY_NHAN      As Date
    Dim ID_TK          As Variant
    Dim MST            As Variant
   
    Dim GHICHU_U       As Variant
    Dim DIA_CHI_U      As Variant
    Dim NGUOI_NOP_U    As Variant
    Dim NGUOI_NOP      As Variant
    Dim GHICHU         As Variant
    Dim DIA_CHI        As Variant
    'Dim NGUOI_NOP As Variant
    Dim strSQL         As String
    Dim LOAI_HS        As String
    Dim HTHUC_NOP      As String
    Dim TRANG_THAI     As String
    Dim SO_HOSO_BSUNG  As String
    Dim MA_DLT         As Variant
    Dim TEN_DLT        As Variant
    Dim NGAY_HDONG_DLT As Variant
    
    With fpSpread1
        .Sheet = 1
        'dhdang sua loi lay thong tin hearder cua TK TAIN
        menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")
        Dim strToaDo As String

        'vttoan: sua lai cel va them thong tin dai ly thue
        '30/07/2011 theo thong tu 28
        If menuId = 1 Or menuId = 71 Or menuId = 74 Or menuId = 75 Then
            .GetText .ColLetterToNumber("F"), 10, maDTNT
                
            .GetText .ColLetterToNumber("F"), 10, MST
                
            .GetText .ColLetterToNumber("H"), 8, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("F"), 12, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("F"), 20, MA_DLT
                
            .GetText .ColLetterToNumber("H"), 18, TEN_DLT
            TEN_DLT = TAX_Utilities_Svr_New.Convert(Trim(TEN_DLT), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("N"), 28, NGAY_HDONG_DLT
                
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
                
            .GetText .ColLetterToNumber("E"), 32, NGNOP
                
            .GetText .ColLetterToNumber("M"), 32, NGNHAP
                
            .GetText .ColLetterToNumber("M"), 36, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
                
        ElseIf menuId = 2 Or menuId = 4 Or menuId = 11 Or menuId = 59 Then
        
            .GetText .ColLetterToNumber("F"), 10, maDTNT
                
            .GetText .ColLetterToNumber("F"), 10, MST
                
            .GetText .ColLetterToNumber("H"), 8, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("F"), 12, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("F"), 20, MA_DLT
                
            .GetText .ColLetterToNumber("H"), 18, TEN_DLT
            TEN_DLT = TAX_Utilities_Svr_New.Convert(Trim(TEN_DLT), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("N"), 28, NGAY_HDONG_DLT
                
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
                
            .GetText .ColLetterToNumber("E"), 32, NGNOP
                
            .GetText .ColLetterToNumber("M"), 32, NGNHAP
                
            .GetText .ColLetterToNumber("M"), 34, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
        ElseIf menuId = 12 Then
            .GetText .ColLetterToNumber("F"), 10, maDTNT
                
            .GetText .ColLetterToNumber("F"), 10, MST
                
            .GetText .ColLetterToNumber("G"), 8, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("F"), 12, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("F"), 20, MA_DLT
                
            .GetText .ColLetterToNumber("H"), 18, TEN_DLT
            TEN_DLT = TAX_Utilities_Svr_New.Convert(Trim(TEN_DLT), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("N"), 28, NGAY_HDONG_DLT
                
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
                
            .GetText .ColLetterToNumber("E"), 32, NGNOP
                
            .GetText .ColLetterToNumber("M"), 32, NGNHAP
                
            .GetText .ColLetterToNumber("M"), 34, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
                
        ElseIf menuId = 15 Or menuId = 16 Or menuId = 50 Or menuId = 51 Or menuId = 36 Then
            .GetText .ColLetterToNumber("G"), 8, maDTNT
                
            .GetText .ColLetterToNumber("G"), 8, MST
                
            .GetText .ColLetterToNumber("H"), 7, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("G"), 9, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("G"), 15, MA_DLT
                
            .GetText .ColLetterToNumber("H"), 14, TEN_DLT
            TEN_DLT = TAX_Utilities_Svr_New.Convert(Trim(TEN_DLT), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("S"), 19, NGAY_HDONG_DLT
                
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
                
            .GetText .ColLetterToNumber("E"), 24, NGNOP
                
            .GetText .ColLetterToNumber("M"), 24, NGNHAP
                
            .GetText .ColLetterToNumber("E"), 28, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
        ElseIf menuId = 6 Or menuId = 5 Or menuId = 70 Then
            .GetText .ColLetterToNumber("H"), 7, maDTNT
                
            .GetText .ColLetterToNumber("H"), 7, MST
                
            .GetText .ColLetterToNumber("H"), 5, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("H"), 9, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("H"), 15, MA_DLT
                
            .GetText .ColLetterToNumber("H"), 13, TEN_DLT
            TEN_DLT = TAX_Utilities_Svr_New.Convert(Trim(TEN_DLT), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("R"), 21, NGAY_HDONG_DLT
                
            .GetText .ColLetterToNumber("E"), 23, vKYLBO
                
            .GetText .ColLetterToNumber("E"), 25, NGNOP
                
            .GetText .ColLetterToNumber("R"), 25, NGNHAP
                
            .GetText .ColLetterToNumber("E"), 29, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
        ElseIf menuId = 8 Or menuId = 9 Then
            .GetText .ColLetterToNumber("K"), 4, maDTNT
            .GetText .ColLetterToNumber("I"), 9, vKYLBO
            .GetText .ColLetterToNumber("I"), 11, NGNOP
            .GetText .ColLetterToNumber("K"), 4, MST
                
            .GetText .ColLetterToNumber("K"), 6, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
            'Ghi chu
            .GetText .ColLetterToNumber("T"), 13, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("K"), 5, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("T"), 11, NGNHAP
                 
        ElseIf menuId = 80 Or menuId = 81 Or menuId = 82 Or menuId = 89 Or menuId = 3 Then
            '         "maDTNT~vKYLBO~NGNOP~MST~DIA_CHI~GHICHU~NGUOI_NOP~NGNHAP"
            strToaDo = "F-10~E-30~E-32~F-10~F-12~M-36~H-8~M-32"
            ThongTin_DLT strToaDo, maDTNT, vKYLBO, NGNOP, MST, DIA_CHI, GHICHU, NGUOI_NOP, NGNHAP
            
        ElseIf menuId = 86 Or menuId = 87 Or menuId = 72 Then
            strToaDo = "G-8~E-22~E-24~G-8~G-9~M-28~G-7~M-24"
            ThongTin_DLT strToaDo, maDTNT, vKYLBO, NGNOP, MST, DIA_CHI, GHICHU, NGUOI_NOP, NGNHAP
        ElseIf menuId = 77 Then
        
            strToaDo = "H-10~E-30~E-32~H-10~H-12~E-36~H-8~R-32"
            ThongTin_DLT strToaDo, maDTNT, vKYLBO, NGNOP, MST, DIA_CHI, GHICHU, NGUOI_NOP, NGNHAP
            
        ElseIf menuId = 73 Then
            '         "maDTNT~vKYLBO~NGNOP~MST~DIA_CHI~GHICHU~NGUOI_NOP~NGNHAP"
            strToaDo = "F-10~E-42~E-44~F-10~F-12~M-48~H-8~M-44"
            ThongTin_DLT strToaDo, maDTNT, vKYLBO, NGNOP, MST, DIA_CHI, GHICHU, NGUOI_NOP, NGNHAP
        Else
               
            .GetText .ColLetterToNumber("G"), 4, maDTNT
                
            .GetText .ColLetterToNumber("E"), 10, KYLBO
        
            .GetText .ColLetterToNumber("E"), 12, NGNOP
            'MST
            .GetText .ColLetterToNumber("G"), 4, MST
            'USE
            USER = strFile(1) & "_NTKCC"
            .GetText .ColLetterToNumber("G"), 6, DIA_CHI
            DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
            'Ghi chu
            .GetText .ColLetterToNumber("M"), 14, GHICHU
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
                
            .GetText .ColLetterToNumber("G"), 5, NGUOI_NOP
            NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
            .GetText .ColLetterToNumber("M"), 12, NGNHAP
        
        End If
        
        'USE
        USER = strFile(1) & "_NTKCC"
       
        ID_TK = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")
        LOAI_HS = changeToLoaiToKhaiQHS(ID_TK)
        DHS_MA = changeToKhaiQHS(ID_TK)
        SO_HOSO = SinhSoHoSo(DHS_MA)
        
        'NGNHAP = Date
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            'NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
        
        If Trim(maDTNT) = vbNullString Then
            maDTNT = "''"
        Else
            maDTNT = "'" & maDTNT & "'"
        End If
        
        If Trim(MA_DLT) = vbNullString Then
            MA_DLT = "''"
        Else
            MA_DLT = "'" & MA_DLT & "'"
        End If

        If Trim(NGAY_HDONG_DLT) = vbNullString Then
            NGAY_HDONG_DLT = "CTOD('')"
        Else
            NGAY_HDONG_DLT = "CTOD('" & format(NGAY_HDONG_DLT, "mm/dd/yyyy") & "')"
        End If
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
       
        NGNOP_S = NGNOP

        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            'NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = DateSerial(Int(Mid$(NGNOP, 7, 4)), Int(Mid$(NGNOP, 4, 2)), Int(Mid$(NGNOP, 1, 2)))
            NGAYNOP_PRINT = NGNOP ' Ngay nop hien thi tren man hinh in BB nop cham (dhdang sua)
        End If
            
        If ID_TK = "01" Or ID_TK = "02" Or ID_TK = "04" Or ID_TK = "71" Or ID_TK = "36" Or ID_TK = "68" Or ID_TK = "25" Then
            If LoaiKyKK = False Then
                KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
                Tinhkykekkhaithang (Mid$(KYKKHAI, 2, 7))
            Else
                KYKKHAI = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
                Tinhkykekkhaiquy (Mid$(KYKKHAI, 2, 7))
            End If
            
        Else
            If (Trim(TAX_Utilities_Svr_New.Month) <> vbNullString Or Trim(TAX_Utilities_Svr_New.Month) <> "") And (Trim(TAX_Utilities_Svr_New.ThreeMonths) = vbNullString Or Trim(TAX_Utilities_Svr_New.ThreeMonths) = "") Then
                KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
                Tinhkykekkhaithang (Mid$(KYKKHAI, 2, 7))
            ElseIf (Trim(TAX_Utilities_Svr_New.Month) = vbNullString Or Trim(TAX_Utilities_Svr_New.Month) = "") And (Trim(TAX_Utilities_Svr_New.ThreeMonths) <> vbNullString Or Trim(TAX_Utilities_Svr_New.ThreeMonths) <> "") Then
                KYKKHAI = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
                Tinhkykekkhaiquy (Mid$(KYKKHAI, 2, 7))
            Else
                KYKK_TU_NGAY = "01/01/" & TAX_Utilities_Svr_New.Year
                KYKK_TU_NGAY_F = "01/01/" & TAX_Utilities_Svr_New.Year
                KYKK_DEN_NGAY = "12/31/" & TAX_Utilities_Svr_New.Year
        End If

        End If
        'TAX_Utilities_Svr_New.ThreeMonths
        
        
       'NGNHAN = Date
        NGAY_NHAN = GetNgayNhap
        NGAY_XL = NGAY_NHAN
        NGAY_HEN = NGAY_XL
'        If Trim(NGAY_NHAN) = vbNullString Then
'            NGAY_NHAN = "CTOD('')"
'        Else
'            'NGAY_NHAN = ToDate(Trim(NGAY_NHAN), DDMMYYYY)
'            NGAY_NHAN = "CTOD('" & format(NGAY_NHAN, "mm/dd/yyyy") & "')"
'        End If

        'dhdang xu ly lay ma phong xu ly tren Form
        'ngay 05-08-2010
        If Not objTaxBusiness Is Nothing Then
        'Get Params
            PHONG_XL_X = objTaxBusiness.PHONG_XU_LY_X1
            PHONG_XL_Y = objTaxBusiness.PHONG_XU_LY_Y1
        End If
        If PHONG_XL_X <> "" And PHONG_XL_Y <> "" Then
            .GetText .ColLetterToNumber(PHONG_XL_X), PHONG_XL_Y, PHONG_XL
            If PHONG_XL <> vbNullString Then
                PHONG_XL = Mid(PHONG_XL, InStr(1, PHONG_XL, "{") + 1, InStr(1, PHONG_XL, "}") - InStr(1, PHONG_XL, "{") - 1)
            End If
        Else
            PHONG_XL = ""
        End If
         Dim F As String
         F = "F"
         HTHUC_NOP = "02"
         'Trang thai ho so chinh thuc bo xung
         
         
            strSQL = "Select top 1 SO_HOSO from QHSCC.dbo.QHS_SO_HOSO where TTHAI_HOSO <> '02' and DHS_MA = '" + DHS_MA + "'and TIN = '" + MST + "' and KYKK_TU_NGAY = '" + KYKK_TU_NGAY + "' order by ID desc"
            If clsDAO.Connected = False Then
            clsDAO.Connect
            End If
            Set rs = clsDAO.Execute(strSQL)
            If Not rs Is Nothing Then
           
                 SO_HOSO_BSUNG = rs(0)
                 If ID_TK = "01" Or ID_TK = "02" Or ID_TK = "71" Then
                    .GetText .ColLetterToNumber("M"), 6, BSUNG
                 ElseIf ID_TK = "04" Then
                    .GetText .ColLetterToNumber("L"), 6, BSUNG
                 ElseIf ID_TK = "72" Then
                    .GetText .ColLetterToNumber("J"), 5, BSUNG
                 Else
                    .GetText .ColLetterToNumber("O"), 2, BSUNG
                 End If
                 If Trim(BSUNG) = "X" Then
                    TRANG_THAI = "01"
                    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & format(NGAY_XL, "mm/dd/yyyy") & "' where ID = '" & rs(0) & "'"
                    bln = clsDAO.ExecuteDLL(strSQL)
                 Else
                    TRANG_THAI = "02"
                    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & format(NGAY_XL, "mm/dd/yyyy") & "' where ID = '" & rs(0) & "'"
                    bln = clsDAO.ExecuteDLL(strSQL)
                 End If
                'ssss = hannop()
                'ssss = Mid$(ssss, 4, 2) + "/" + Mid$(ssss, 1, 2) + "/" + Mid$(ssss, 7, 4)
                'If (NGNOP > CDate(ssss)) Then
                '    TRANG_THAI = "02"
                '    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & NGAY_XL & "' where ID = '" & rs(0) & "'"
                '    bln = clsDAO.ExecuteDLL(strSQL)
                'Else
                    
                'End If
      
'                If CTHUC = "1" And BSUNG = "" Then
'                    TRANG_THAI = "03"
'                    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & NGAY_XL & "' where ID = '" & rs(0) & "'"
'                    bln = clsDAO.ExecuteDLL(strSQL)
'                ElseIf CTHUC = "" And BSUNG = "1" Then
'                    TRANG_THAI = "02"
'                    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & NGAY_XL & "' where ID = '" & rs(0) & "'"
'                    bln = clsDAO.ExecuteDLL(strSQL)
'                Else
'                    TRANG_THAI = "03"
'                    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & NGAY_XL & "' where ID = '" & rs(0) & "'"
'                    bln = clsDAO.ExecuteDLL(strSQL)
'                End If
            Else
                    TRANG_THAI = "01"
            End If
         
         
         
        strSQL = "Select top 1 SO_TEP from QHSCC.dbo.QHS_SO_HOSO where SO_HIEU_TEP = '' and DHS_MA = '" + DHS_MA + "' and HTHUC_NOP = '02' and NGUOI_NHAP = '" + USER + "' order by ID desc"
        If clsDAO.Connected = False Then
            clsDAO.Connect
        End If
        
        Set rs = clsDAO.Execute(strSQL)
        
        If rs Is Nothing Then
            SO_TEP = "0"
        Else
            SO_TEP = rs(0)
        End If
        
        If SO_TEP = "" Then SO_TEP = "0"
        SO_TEP = Trim(str(Int(SO_TEP) + 1))
        TRICH_YEU = TinhPhuLucTk
         
        sSQLCol = "DHS_MA, SO_HOSO, TIN,TEN,DIA_CHI,KYKK_TU_NGAY,KYKK_DEN_NGAY, NGAY_NHAN,NGUOI_NOP,NGAY_NHAP,NGUOI_NHAP,HAN_XULY,NGAY_HEN,PHONG_XLY,PHONG_XLY_HIENTAI,GHI_CHU,NGAY_NOP,TTHAI_HOSO,GUI_BD,HTHUC_NOP,SO_TEP,SO_HOSO_BSUNG,TRICH_YEU"
        sSQLVal = DHS_MA & ",'" & SO_HOSO & "','" & MST & "','" & NGUOI_NOP & "','" & DIA_CHI & "','" & KYKK_TU_NGAY & "','" & KYKK_DEN_NGAY & "','" & format(NGAY_NHAN, "mm/dd/yyyy") & "','" & NGUOI_NOP & "','" & format(NGAY_NHAN, "mm/dd/yyyy") & "','" & USER & "','" & format(NGAY_XL, "mm/dd/yyyy") & "','" & format(NGAY_XL, "mm/dd/yyyy") & "','" & PHONG_XL & "','" & PHONG_XL & "','" & GHICHU & "','" & format(NGNOP, "mm/dd/yyyy") & "','" & TRANG_THAI & "','" & F & "','" & HTHUC_NOP & "','" & SO_TEP & "','" & SO_HOSO_BSUNG & "','" & TRICH_YEU & "'"
       
        sSQL = "INSERT INTO QHSCC.dbo.QHS_SO_HOSO" & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
        'bln = clsDAO.ExecuteDLL(sSQL)
        
        'dhdang
    'in bien ban phat nop cham
    'ngay 20/09
    '-----------------------------------------------------------
        Dim kieukykk As String
        If ID_TK = "01" Or ID_TK = "02" Or ID_TK = "04" Or ID_TK = "71" Or ID_TK = "36" Or ID_TK = "68" Or ID_TK = "25" Then
            If LoaiKyKK = False Then
                kieukykk = "M"
            Else
                kieukykk = "Q"
            End If
        Else
            If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
                kieukykk = "M"
            ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then
                kieukykk = "Q"
            ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
                kieukykk = "Y"
            Else
                kieukykk = "Y"
            End If
        End If
        
        If CheckThanhTraKiemTra(MST_PRINT, TinhLoaiThue(DHS_MA), KYKK_TU_NGAY_F, KYKK_DEN_NGAY) = True Then
'            If MessageBox("0130", msYesNo, miQuestion) = mrYes Then
'                    frmInBienBanPhatNopCham.Show 1
'            End If
            MessageBox "0132", msOKOnly, miWarning
            Exit Function
        End If
            
       Dim TK_PS As Variant
     
       If ID_TK = "73" Then
            .GetText .ColLetterToNumber("Q"), 49, TK_PS
       End If
            
        If KiemTraNopCham(KYKK_TU_NGAY_F, kieukykk, NGNOP_S) = True And TK_PS <> "1" Then
            If MessageBox("0130", msYesNo, miQuestion) = mrYes Then
                    frmInBienBanPhatNopCham.Show 1
            End If
        End If
        '--------------------------------------
    End With
     
     
    Prepare_QLT = sSQL
    'clsDAO.Disconnect
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       NTKCC
' Procedure  :       ThongTin_DLT
' Description:       Set lai thong tin cho dai ly thue
' Created by :       Project Administrator
' Machine    :       NKHOAN-PC
' Date-Time  :       12/7/2011-10:43:58
'
' Parameters :       strToaDo: la chuoi chua toa do cua lan luot cac bien so:
'                    - maDTNT,vKYLBO, NGNOP, MST, DIA_CHI, GHICHU, NGUOI_NOP, NGNHAP
'                    - strToaDo co dang: "A-10~AB-20....."
'--------------------------------------------------------------------------------
'</CSCM>
Private Function ThongTin_DLT(strToaDo As Variant, _
                              maDTNT As Variant, _
                              KYLBO As Variant, _
                              NGNOP As Variant, _
                              MST As Variant, _
                              DiaChi As Variant, _
                              GHICHU As Variant, _
                              NguoiNop As Variant, _
                              NgayNhap As Variant)
    Dim iRow, iCol, arrToaDo, arr As Variant
    arrToaDo = Split(strToaDo, "~")
                
    With fpSpread1
        .Sheet = 1
        arr = Split(arrToaDo(0), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, maDTNT
        
        arr = Split(arrToaDo(1), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, KYLBO
        
        arr = Split(arrToaDo(2), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, NGNOP
        
        arr = Split(arrToaDo(3), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, MST
        

        arr = Split(arrToaDo(4), "-")
       iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, DiaChi
        If DiaChi <> vbNullString Then
            DiaChi = TAX_Utilities_Svr_New.Convert(Trim(DiaChi), UNICODE, TCVN)
        End If
        
        
        arr = Split(arrToaDo(5), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, GHICHU
        If GHICHU <> vbNullString Then
            GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
        End If
      
        
        arr = Split(arrToaDo(6), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, NguoiNop
         If NguoiNop <> vbNullString Then
            NguoiNop = TAX_Utilities_Svr_New.Convert(Trim(NguoiNop), UNICODE, TCVN)
        End If
        
        arr = Split(arrToaDo(7), "-")
        iRow = Val(arr(1))
        iCol = arr(0)
        .GetText .ColLetterToNumber(iCol), iRow, NgayNhap
        
               
    End With

    
End Function

Private Function SinhSoHoSo(DHS_MA) As String
    Dim rs As ADODB.Recordset
    Dim s, s1 As String
    Dim i As Integer
    Dim SQ As String
    Dim D As Date
    Dim strSQL As String
  
    Set rs = New ADODB.Recordset
     If clsDAO.Connected = False Then
            Me.MousePointer = vbHourglass
            frmSystem.MousePointer = vbHourglass
            clsDAO.CreateConnectionStringSQL spathQHSCC
            clsDAO.Connect
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
    End If
    
    strSQL = "Select Top 1 SO_HOSO from QHSCC.dbo.QHS_SO_HOSO  where DHS_MA = '" & DHS_MA & "' Order By ID desc"
    Set rs = clsDAO.Execute(strSQL)
    
     If rs Is Nothing Then
        SQ = 0
     Else
        
        s = rs(0)
        i = InStrRev(s, "/")
        s1 = Right(s, Len(s) - i)
        SQ = CLng(s1)
     End If
    SQ = SQ + 1

    strSQL = "Select getdate() as DATE_NOW"
    Set rs = clsDAO.Execute(strSQL)
    D = rs(0)
    clsDAO.Disconnect

    s = format(D, "YYMMDD") + "/" + DHS_MA + "/"
    s1 = Trim(str(SQ))
    SinhSoHoSo = s + s1
End Function

Private Function changeToKhaiQHS(strMaToKhai) As String
Dim DHS_MA As String
Dim strSQL As String
Dim tkPhatSinh As Variant

    menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")
     With fpSpread1
        .Sheet = 1
        If menuId = 70 Then
            .GetText .ColLetterToNumber("AD"), 3, tkPhatSinh
        End If
        
        If menuId = 81 Then
            .GetText .ColLetterToNumber("Q"), 37, tkPhatSinh
        End If

        If menuId = 73 Then
            .GetText .ColLetterToNumber("Q"), 49, tkPhatSinh
        End If

    End With

    On Error Resume Next
    
         Select Case strMaToKhai
            Case "48"
                 DHS_MA = "425"
            Case "46"
                 DHS_MA = "423"
            Case "47"
                 DHS_MA = "424"
            Case "49"
                 DHS_MA = "426"
            Case "14"
                 DHS_MA = "173"
            Case "07"
                 DHS_MA = "33"
            Case "03"
                 DHS_MA = "31"
            Case "08"
                 DHS_MA = "80"
            Case "04"
            If LoaiKyKK = True Then
                DHS_MA = "529"
            Else
                DHS_MA = "17"
          
            End If

        Case "05"
            DHS_MA = "81"

        Case "38"
            DHS_MA = "271"

        Case "54"
            DHS_MA = "25"

        Case "09"
            DHS_MA = "177"

        Case "02"

            If LoaiKyKK = True Then
                DHS_MA = "528"
            Else
                DHS_MA = "30"
            End If
                 
        Case "06"
            DHS_MA = "27"

        Case "37"
            DHS_MA = "354"

        Case "53"
            DHS_MA = "22"

        Case "11"
            DHS_MA = "174"

        Case "12"
            DHS_MA = "75"

        Case "01"

            If LoaiKyKK = True Then
                DHS_MA = "527"
            Else
                DHS_MA = "16"
            End If

        Case "36"
            DHS_MA = "23"

        Case "40"
            DHS_MA = "372"

        Case "39"
            DHS_MA = "24"

        Case "50"
            DHS_MA = "25"

        Case "51"
            DHS_MA = "371"

        Case "16"
            DHS_MA = "161"

        Case "15"
            DHS_MA = "21"

        Case "17"
            DHS_MA = "36"

        Case "80"
            DHS_MA = "83"

        Case "81"

            If tkPhatSinh = vbNullString Then
                DHS_MA = "472"
            Else
                DHS_MA = "473"
            End If

        Case "82"
            DHS_MA = "474"

        Case "86"
            DHS_MA = "181"

        Case "87"
            DHS_MA = "182"
            
        Case "89"
            DHS_MA = "180"

        Case "70"

            If tkPhatSinh = vbNullString Then
                DHS_MA = "63"
            Else
                DHS_MA = "351"
            End If

        Case "73"

            If tkPhatSinh = vbNullString Then
                DHS_MA = "447"
            Else
                DHS_MA = "448"
            End If

        Case "71"

            If LoaiKyKK = True Then
                DHS_MA = "530"
            Else
                DHS_MA = "526"
            End If
            
        Case "72"
            DHS_MA = "441"
            
        Case "74"
            DHS_MA = "449"
            
        Case "75"
            DHS_MA = "454"
            
        Case "59"
            DHS_MA = "387"
            
        Case "77"
            DHS_MA = "450"
            
        Case Else
            DHS_MA = ""
            
    End Select

    changeToKhaiQHS = DHS_MA
End Function

Private Function changeToLoaiToKhaiQHS(strMaToKhai) As String
    Dim DHS_MA     As String
    Dim strSQL     As String
    Dim tkPhatSinh As Variant
    menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")

    With fpSpread1
        .Sheet = 1

        If menuId = 70 Then
            .GetText .ColLetterToNumber("AD"), 3, tkPhatSinh
        End If
        
        If menuId = 81 Then
            .GetText .ColLetterToNumber("Q"), 37, tkPhatSinh
        End If
        
        If menuId = 73 Then
            .GetText .ColLetterToNumber("Q"), 49, tkPhatSinh
        End If

    End With

    On Error Resume Next
    
    Select Case strMaToKhai

        Case "37"
            DHS_MA = "200514"

        Case "53"
            DHS_MA = "200503"

        Case "11"
            DHS_MA = "200201"

        Case "01"

            If LoaiKyKK = True Then
                DHS_MA = "200121"
            Else
                DHS_MA = "200101"
            End If

        Case "02"

            If LoaiKyKK = True Then
                DHS_MA = "200122"
            End If

        Case "04"

            If LoaiKyKK = True Then
                DHS_MA = "200123"
            End If

        Case "71"

            If LoaiKyKK = True Then
                DHS_MA = "200124"
            Else
                DHS_MA = "200105"
            End If

        Case "12"
            DHS_MA = "200202"

        Case "36"
            DHS_MA = "200531"

        Case "40"
            DHS_MA = "200517"

        Case "39"
            DHS_MA = "200504"

        Case "50"
            DHS_MA = "200507"

        Case "51"
            DHS_MA = "200516"

        Case "16"
            DHS_MA = "200502"

        Case "15"
            DHS_MA = "200501"

        Case "80"
            DHS_MA = "300110"

        Case "81"

            If tkPhatSinh = vbNullString Then
                DHS_MA = "200905"
            Else
                DHS_MA = "200906"
            End If

        Case "82"
            DHS_MA = "300122"

        Case "86"
            DHS_MA = "200804"

        Case "87"
            DHS_MA = "200805"
            
        Case "89"
            DHS_MA = "200803"

        Case "70"

            If tkPhatSinh = vbNullString Then
                DHS_MA = "200902"
            Else
                DHS_MA = "200904"
            End If

        Case Else
            DHS_MA = ""
    End Select

    changeToLoaiToKhaiQHS = DHS_MA
End Function
Private Function GetNgayNhap() As Date
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
On Error GoTo ErrHandle
    Set rsReturn = New ADODB.Recordset
    'connect to database QLT
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionStringSQL spathQHSCC
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
    End If
   
    strSQL = "SELECT Getdate()"
    Set rsReturn = clsDAO.Execute(strSQL)

    GetNgayNhap = rsReturn(0)
    
     Exit Function
ErrHandle:
    'Connect DB fail
    clsDAO.Disconnect
    SaveErrorLog Me.Name, "GetPXL", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function
Private Sub Tinhkykekkhaithang(KyKK As String)
    Dim ss1 As String
    Dim ss2 As String
        
                ss1 = "01/" + Trim(KyKK)
                Select Case Left$(KyKK, 2)
                    Case "01", "03", "05", "07", "08", "10", "12": ss2 = "31/" + KyKK
                    Case "04", "06", "09", "11": ss2 = "30/" + Trim(KyKK)
                    Case "02":
                        If (Right$(KyKK, 4) Mod 4) = 0 Then
                                ss2 = "29/" + Trim(KyKK)
                            Else
                                ss2 = "28/" + Trim(KyKK)
                            End If
                End Select
        
        KYKK_TU_NGAY = Mid$(ss1, 4, 2) + "/" + Mid$(ss1, 1, 2) + "/" + Mid$(ss1, 7, 4)
        KYKK_TU_NGAY_F = Mid$(ss1, 1, 2) + "/" + Mid$(ss1, 4, 2) + "/" + Mid$(ss1, 7, 4)
        KYKK_DEN_NGAY = Mid$(ss2, 4, 2) + "/" + Mid$(ss2, 1, 2) + "/" + Mid$(ss2, 7, 4)
End Sub
Private Sub Tinhkykekkhaiquy(KyKK As String)
    Dim s1 As String
    Dim s2 As String
    Dim ss1 As String
    Dim ss2 As String
    
    ss1 = Mid$(KyKK, 1, 2)
    ss2 = Mid$(KyKK, 4, 4)
    
    
               Select Case ss1
                    Case "01":
                        s1 = "01/01/" + ss2
                        s2 = "31/03/" + ss2
                    Case "02":
                        s1 = "01/04/" + ss2
                        s2 = "30/06/" + ss2
                    Case "03":
                        s1 = "01/07/" + ss2
                        s2 = "30/09/" + ss2
                    Case "04":
                        s1 = "01/10/" + ss2
                        s2 = "31/12/" + ss2
                End Select
        
        KYKK_TU_NGAY = Mid$(s1, 4, 2) + "/" + Mid$(s1, 1, 2) + "/" + Mid$(s1, 7, 4)
        KYKK_TU_NGAY_F = Mid$(s1, 1, 2) + "/" + Mid$(s1, 4, 2) + "/" + Mid$(s1, 7, 4)
        KYKK_DEN_NGAY = Mid$(s2, 4, 2) + "/" + Mid$(s2, 1, 2) + "/" + Mid$(s2, 7, 4)
End Sub

Private Sub Insert_QHS()

    On Error GoTo ErrHandle

    Dim strSQL           As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String, strSQL_KHBS As String
    Dim HdrID            As Variant, strDate() As String, dDate As Date
    Dim rs               As ADODB.Recordset, i As Long
    Dim blHoiTonTai      As Integer
    Dim blUpdateTHUETKY2 As Boolean
    Dim bln              As Boolean
    Dim blnKTRB          As Integer
    Dim sSaiCT11         As String
    Dim vKYLBO           As Variant
    Dim vNGAYQUET        As Variant
    Dim vNGAY_DAU_KYLBO  As Variant
    Dim sSQL             As String
    'Dim menuId As Integer
    Dim NGAY_HIENTAI     As Date
    Dim s                As String
    Dim TEP_ID           As String
    
    'NGAY_HIENTAI = GetNgayNhap
    'Set rs = New ADODB.Recordset
    sSaiCT11 = ""

    '***************************
    'ThanhDX added
    'Date:23/11/2005
    If TAX_Utilities_Svr_New.Data(0) Is Nothing Then Exit Sub
    '***************************
       
    blnSaveSuccess = False
    
    CallFinish

    ' Kiem tra xem da khoa so trong ky lap bo nay chua
    ' hlnam edit
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionStringSQL spathQHSCC
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
    End If

    menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")

    With fpSpread1
        .Sheet = 1

        If menuId = 8 Or menuId = 9 Then
            .GetText .ColLetterToNumber("I"), 9, vKYLBO
            ' vttoan: lay KyLapBo
        ElseIf menuId = 15 Or menuId = 16 Or menuId = 50 Or menuId = 51 Or menuId = 36 Or menuId = 72 Or menuId = 86 Or menuId = 87 Or menuId = 72 Then
            .GetText .ColLetterToNumber("E"), 22, vKYLBO
        ElseIf menuId = 5 Or menuId = 6 Or menuId = 70 Then
            .GetText .ColLetterToNumber("E"), 23, vKYLBO
        ElseIf menuId = 1 Or menuId = 2 Or menuId = 4 Or menuId = 11 Or menuId = 12 Or menuId = 80 Or menuId = 81 Or menuId = 82 Or menuId = 89 Or menuId = 71 Or menuId = 3 Or menuId = 59 Or menuId = 74 Or menuId = 75 Or menuId = 77 Then
            .GetText .ColLetterToNumber("E"), 30, vKYLBO
        ElseIf menuId = 73 Then
            .GetText .ColLetterToNumber("E"), 42, vKYLBO
        Else
            .GetText .ColLetterToNumber("E"), 10, vKYLBO
        End If
        
        vNGAY_DAU_KYLBO = "01/" & IIf(Len(Trim(vKYLBO)) = 6, "0" & vKYLBO, vKYLBO) ' Lay ngay dau cua ky lap bo de xem ngay quet co phu hop voi ky khoa so hay khong?
        
        If Trim(vKYLBO) = vbNullString Or Trim(vKYLBO) = "../...." Then
           
            DisplayMessage "0106", msOKOnly, miCriticalError
            Exit Sub
           
        Else

            If Len(Trim(vKYLBO)) = 6 Then
                vKYLBO = "'0" & vKYLBO & "'"
            Else
                vKYLBO = "'" & vKYLBO & "'"
            End If
        End If
          
        strSQL_DTL = Prepare_QLT
        
        If clsDAO.Connected = False Then
            Me.MousePointer = vbHourglass
            frmSystem.MousePointer = vbHourglass
            clsDAO.CreateConnectionStringSQL spathQHSCC
            clsDAO.Connect
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
        End If
              
        If Trim(strSQL_DTL) <> vbNullString Then
            
            bln = clsDAO.ExecuteDLL(strSQL_DTL)
            
            '        ' Dong tep
            '        If SO_TEP = "50" Then
            '
            '            'Sinh so hieu tep
            '
            '             s = format(NGAY_HIENTAI, "YYMM")
            '             s = s + DHS_MA
            '
            '                 If clsDAO.Connected = False Then
            '                        Me.MousePointer = vbHourglass
            '                        frmSystem.MousePointer = vbHourglass
            '                        clsDAO.CreateConnectionStringSQL spathQHSCC
            '                        clsDAO.Connect
            '                        frmSystem.MousePointer = vbDefault
            '                        Me.MousePointer = vbDefault
            '                End If
            '                strSQL = "Select top 1 SO_HIEU, NGAY_TAO from QHSCC.dbo.QHS_TEP_HOSO where SO_HIEU like '" & s & "%' order by ID DESC "
            '                Set rs = clsDAO.Execute(strSQL)
            '
            '            If rs Is Nothing Then
            '                    s = s + "-1"
            '                Else
            '                    If Left$(rs(0), 4) <> format(NGAY_HIENTAI, "YYMM") Then
            '                        s = s + "-1"
            '                    Else
            '                        i = Right$(rs(0), Len(rs(0)) - InStr(1, rs(0), "-"))
            '                        i = i + 1
            '                        s = s & "-" & i
            '                    End If
            '                End If
            '
            '                TEP_ID = s
            '                If clsDAO.Connected = False Then
            '                        Me.MousePointer = vbHourglass
            '                        frmSystem.MousePointer = vbHourglass
            '                        clsDAO.CreateConnectionStringSQL spathQHSCC
            '                        clsDAO.Connect
            '                        frmSystem.MousePointer = vbDefault
            '                        Me.MousePointer = vbDefault
            '            End If
            '                'Update QHS_SO_HOSO
            '                strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set SO_HIEU_TEP = '" & s & "' where SO_HIEU_TEP = '' and DHS_MA = '" + DHS_MA + "' and HTHUC_NOP = '02' and NGUOI_NHAP = '" + USER + "'"
            '                bln = clsDAO.ExecuteDLL(strSQL)
            '                ' insert QHS_TEP_HOSO
            '                strSQL = "insert into QHSCC.dbo.QHS_TEP_HOSO (SO_HIEU, DHS_MA, KYKK_TU_NGAY, KYKK_DEN_NGAY, NGAY_TAO, SO_HOSO, TTHAI, NGUOI_TAO) values ('" & s & "', '" & DHS_MA & "', " & KYKK_TU_NGAY & ", " & KYKK_DEN_NGAY & ", '" & format(NGAY_HIENTAI, "mm/dd/yyyy") & "', '" & SO_TEP & "', '', '" & USER & "')"
            '                bln = clsDAO.ExecuteDLL(strSQL)
            '        End If
          
            'Debug.Print strSQL_DTL
        End If
        
        clsDAO.Disconnect

        If Not objTaxBusiness Is Nothing Then
            'Get Params
            objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
        End If
   
        StartReceiveForm
        '***************************
        Set rs = Nothing
        blnSaveSuccess = True
        Exit Sub
ErrHandle:

        If Err.Number = -2147217865 Then
            MessageBox "0094", msOKOnly, miCriticalError
        ElseIf Err.Number = 53 Then
            'MessageBox "0096", msOKOnly, miCriticalError
            ' "0109" Thong bao Truoc khi chay ban hay khoi tao ky ke khai ben UD VATCC truoc roi moi nhan bang NTKCC
            MessageBox "0109", msOKOnly, miCriticalError
        Else
            MessageBox "0049", msOKOnly, miCriticalError
            SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
        End If

        On Error GoTo ExitErr
        'Rollback
        'clsDAO.RollbackTrans
        clsDAO.Disconnect
        Set rs = Nothing
        blnSaveSuccess = True
        Exit Sub
ExitErr:
        Set rs = Nothing
        SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
        MessageBox "0049", msOKOnly, miCriticalError
        blnSaveSuccess = True
    End With

End Sub


'dhdang tao ham check connection toi QHS
'06/07/2010

Private Function CheckConnection() As Boolean
    Dim flag As Boolean
    
    clsDAO.CreateConnectionStringCheckSQL spathQHSCC
    clsDAO.Connect_qhs
    flag = clsDAO.Connected_qhs
    clsDAO.DisConnect_qhs
    CheckConnection = flag

End Function
'dhdang tao ham tinh so phu luc cua to khai
'ngay 06/07/2010
Private Function TinhPhuLucTk() As String
    Dim str As String
    Dim soPl As String
    Dim i As Integer
    
    soPl = TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 2
    For i = 1 To soPl
        If TAX_Utilities_Svr_New.NodeValidity.childNodes(i).Attributes.getNamedItem("Active").nodeValue = 1 Then
            str = str & "[" & TAX_Utilities_Svr_New.NodeValidity.childNodes(i).Attributes.getNamedItem("Caption").nodeValue & "];"
        End If
    Next
    If str <> "" Then
        str = "Phu Luc :" & str
    End If
    TinhPhuLucTk = str
End Function

' Cac ham dung de in BB nop cham va check thanh tra kiem tra (dhdang sua)
Public Function InitParametersPrint() As Boolean

'ThanhDX modified
'Date: 10/04/06
    Dim strTaxID As String, strID As String
    Dim blnConnected As Boolean
    Dim strValidDate As String, strTempDate As String
    Dim rsParams As ADODB.Recordset
    Dim strPhongXuLy As String
    Dim rsTaxInfor As ADODB.Recordset
    
On Error GoTo ErrHandle
    
    strID = 54
    SetNodeMenu strID
    'SetPeriod Right$(strTaxReportInfo, 6)
    TAX_Utilities_Svr_New.NodeValidity = GetValidityNode
    
    '*******************************
    'RestoreDataFile (strData)
'    If Not RestoreDataFile(strData) Then  ', rsTaxInfor
'        If blnReceiveByBarcode Then
'            MessageBox "0057", msOKOnly, miCriticalError
'        Else
'            MessageBox "0053", msOKOnly, miCriticalError
'        End If
'        Exit Function
'    End If
    
    InitParametersPrint = True

    Exit Function
ThamSoErrHandle:
    DisplayMessage "0078", msOKOnly, miCriticalError
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "InitParameters", Err.Number, Err.Description
End Function
'ham kiem tra nop cham
'dhdang
'ngay 31\09\2010

Private Function KiemTraNopCham(kykekhaitu As String, _
                                kieuky As String, _
                                ngayNop As String) As Boolean
    Dim b       As Boolean
    Dim D1      As Date
    Dim d2      As Date
    Dim s       As String
    Dim sss     As String
    Dim s1      As String
    Dim s2      As String
    Dim j       As String
    Dim rs      As ADODB.Recordset
    Dim rs1     As ADODB.Recordset
    Dim strSQL  As String
    Dim strSQL1 As String
    
    'To khai theo nam
    If kieuky = "Y" Then
        If MATK_PRINT = "03" Or MATK_PRINT = "07" Or MATK_PRINT = "08" Then
            CAN_CU1 = GetAttribute(GetMessageCellById("0131"), "Msg")
        Else
            CAN_CU1 = GetAttribute(GetMessageCellById("0124"), "Msg")
        End If

        'CAN_CU1 = TAX_Utilities_Svr_New.Convert(CAN_CU1, TCVN, UNICODE)
        CAN_CU2 = GetAttribute(GetMessageCellById("0125"), "Msg")
        KyKeKhai = "N�m " + Mid$(kykekhaitu, 7, 4)

        s = Mid$(kykekhaitu, 7, 4)
        s = "31/03/" + Trim(str(Int(s) + 1))
        D1 = StringToDate(s)
        d2 = StringToDate(ngayNop)
        '    End If

        ' Quyet toan thue

        '     If (Mid$(txtMaLoaiHoSo.Text, 1, 1) = "3") Then
        '        CAN_CU1 = "C?n c? kho?n 2, ?i?u 32 lu?t Qu?n l� thu? quy ??nh th?i h?n n?p h? s? quy?t to�n thu?: ""Ch?m nh?t l� ng�y th? ch�n m??i k? t? ng�y k?t th�c n?m d??ng l?ch ho?c n?m t�i ch�nh ??i v?i h? s? quy?t to�n thu? n?m"""
        '        CAN_CU2 = "C�c h�nh vi tr�n ?� vi ph?m v�o ?i?u 9 Ch??ng 1 c?a ngh? ??nh s? 98/2007/N?-CP ng�y 07 th�ng 06 n?m 2007 quy ??nh v? ""x? l� vi ph?m ph�p lu?t v? thu? v� c??ng ch? thi h�nh quy?t ??nh h�nh ch�nh thu?""."
        '        KyKeKhai = "N�m " + Mid$(kykekhaitu, 7, 4)
        '
        '        s = Mid$(kykekhaitu, 7, 4)
        '        s = "31/03/" + Trim(str(Int(s) + 1))
        '        D1 = StringToDate(s)
        '        d2 = StringToDate(ngaynop)
        '    End If

        ' To Khai Quy
    ElseIf kieuky = "Q" Then
        CAN_CU1 = GetAttribute(GetMessageCellById("0126"), "Msg")
        CAN_CU2 = GetAttribute(GetMessageCellById("0127"), "Msg")
        s2 = Mid$(kykekhaitu, 4, 2)
        s1 = Mid$(kykekhaitu, 7, 4)

        Select Case s2

            Case "01":
                s = "30/04/" + s1
                KyKeKhai = "Qu� 1/" + Mid$(kykekhaitu, 7, 4)

            Case "04":
                s = "31/07/" + s1
                KyKeKhai = "Qu� 2/" + Mid$(kykekhaitu, 7, 4)

            Case "07":
                s = "30/10/" + s1
                KyKeKhai = "Qu� 3/" + Mid$(kykekhaitu, 7, 4)

            Case "10":

                If Int(s1) = 2013 Then
                    s = "06/02/" + Trim(str(Int(s1) + 1))
                    KyKeKhai = "Qu� 4/" + Mid$(kykekhaitu, 7, 4)

                Else
                    s = "31/01/" + Trim(str(Int(s1) + 1))
                    KyKeKhai = "Qu� 4/" + Mid$(kykekhaitu, 7, 4)
    
                End If

        End Select

        D1 = StringToDate(s)
        d2 = StringToDate(ngayNop)
        'End If

        'To khai thang
    ElseIf kieuky = "M" Then
        CAN_CU1 = GetAttribute(GetMessageCellById("0128"), "Msg")
        CAN_CU2 = GetAttribute(GetMessageCellById("0129"), "Msg")
        KyKeKhai = "th�ng " + Mid$(kykekhaitu, 4, 7)

        s2 = Mid$(kykekhaitu, 4, 2)
        s1 = Mid$(kykekhaitu, 7, 4)

        If Int(s2) < 12 Then
            sss = Trim(str(Int(s2) + 1))
            s = "20/" + Switch(Len(sss) = 1, "0" + sss, Len(sss) = 2, sss) + "/" + s1
        Else
            s = "20/01/" + Trim(str(Int(s1) + 1))
        End If

        D1 = StringToDate(s)
        d2 = StringToDate(ngayNop)
    End If

    If clsDAO.Connected_qhs = False Then
        clsDAO.Connect_qhs
    End If
    
    Set rs = New Recordset
    Set rs1 = New Recordset
    strSQL = "Select Count(*) from QHS_DM_NGAYNGHI  where NgayNghi = '" & format(D1, "MM/DD/YYYY") & "'"
    Set rs = clsDAO.Execute_Qhs(strSQL)
    j = rs(0)
    While (j <> 0) Or (Weekday(D1) = 1) Or (Weekday(D1) = 7)
        D1 = D1 + 1
        Set rs1 = New Recordset
        strSQL1 = "Select Count(*) from QHS_DM_NGAYNGHI  where NgayNghi = '" & format(D1, "MM/DD/YYYY") & "'"
        Set rs1 = clsDAO.Execute_Qhs(strSQL1)
        j = rs1(0)
    Wend

    clsDAO.DisConnect_qhs

    If (D1 + 5) <= d2 Then
        KiemTraNopCham = True
    Else
        KiemTraNopCham = False
    End If

    'If ((txtMaLoaiHoSo.Text >= "100123") And (txtMaLoaiHoSo.Text <= "100132")) Or (txtMaLoaiHoSo.Text = "100153") Then
    '    If D1 < d2 Then KiemTraNopCham = True
    'End If
    'KiemTraNopCham = True
    HAN_NOP = D1
End Function

'dhdang
'ngay 01/09/2010
Private Function Prepare_In() As String
    Dim strSQL As String
    Dim strSQL1 As String
   Dim rs As ADODB.Recordset
   Dim NGNOP As Variant
        fpSpread1.GetText fpSpread1.ColLetterToNumber("E"), 12, NGNOP
        'NGNOP_S = NGNOP
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            'NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = DateSerial(Int(Mid$(NGNOP, 7, 4)), Int(Mid$(NGNOP, 4, 2)), Int(Mid$(NGNOP, 1, 2)))
            NGAYNOP_PRINT = "CTOD('" & format(NGNOP, "mm/dd/yyyy") & "')"
        End If
             
End Function

Private Function StringToDate(s As String) As Date
        If Trim(s) = vbNullString Then
            StringToDate = "CTOD('')"
        Else
            StringToDate = DateSerial(Int(Mid$(s, 7, 4)), Int(Mid$(s, 4, 2)), Int(Mid$(s, 1, 2)))
        End If
End Function
Public Function CheckThanhTraKiemTra(TIN As String, LOAI_THUE As String, KYKK_TU_NGAY As String, KYKK_DEN_NGAY As String) As Boolean
    Dim k1 As String, k2 As String, strSQL As String
    Dim b As Boolean
    Dim rs As ADODB.Recordset
    
    b = False
    If clsDAO.Connected_qhs = False Then
            clsDAO.Connect_qhs
    End If
    
    Set rs = New Recordset
    strSQL = "SELECT QHS_TTRA_KTRA_DTL.LOAI_THUE, QHS_TTRA_KTRA_DTL.KY_TT_TU, QHS_TTRA_KTRA_DTL.KY_TT_DEN FROM QHS_TTRA_KTRA_HDR INNER JOIN QHS_TTRA_KTRA_DTL ON QHS_TTRA_KTRA_HDR.ID = QHS_TTRA_KTRA_DTL.HDR_ID WHERE  (QHS_TTRA_KTRA_HDR.MA_DTNT = '" + TIN + "') AND (QHS_TTRA_KTRA_DTL.LOAI_THUE =  '" + LOAI_THUE + "')"
    Set rs = clsDAO.Execute_Qhs(strSQL)
    'rs.Open s, cn
    If rs Is Nothing Then
        CheckThanhTraKiemTra = b
        clsDAO.DisConnect_qhs
        Exit Function
    End If
    
                Dim D1 As Date
                Dim d2 As Date
                Dim d3 As Date
                Dim d4 As Date
                
                D1 = StringToDate(KYKK_TU_NGAY)
                d2 = StringToDate(KYKK_DEN_NGAY)
                
                d3 = rs!KY_TT_TU
                d4 = rs!KY_TT_DEN
                
                If Not ((d4 < D1) Or (d3 > d2)) Then
                    b = True
                End If
    
    clsDAO.DisConnect_qhs
    CheckThanhTraKiemTra = b
End Function

Public Function TinhLoaiThue(MAHS As String) As String
    Dim k1 As String, k2 As String, strSQL As String, loaithue As String
    Dim b As Boolean
    Dim rs As ADODB.Recordset
    
    b = False
    If clsDAO.Connected_qhs = False Then
            clsDAO.Connect_qhs
    End If
    
    Set rs = New Recordset
    strSQL = "select loai_thue from QHS_DM_HOSO where MA = '" + MAHS + "'"
    Set rs = clsDAO.Execute_Qhs(strSQL)
    'rs.Open s, cn
    loaithue = rs(0)
    clsDAO.DisConnect_qhs
    TinhLoaiThue = loaithue
End Function

' Ham lay ve so tt quet An chi
Private Function getSoTTTK_AC(ByVal strID As String, _
                              arrStrHeaderData() As String, _
                              strData As String) As String
    Dim rsResult     As ADODB.Recordset
    Dim strSQL       As String
    Dim clsConn      As New TAX_Utilities_Svr_New.clsADO
    
    Dim arrDeltail() As String
    Dim arrDate()    As String
    Dim dTempDate    As Date
    Dim dTempDate1   As Date
    
    Dim strSTT       As String
    Dim vTuNgay      As Date
    Dim vDenNgay     As Date
    Dim vMaSoThue    As String
    On Error GoTo ErrHandle
    
    On Error GoTo ConnectErrHandle

    'connect to database VAT
    If Not clsConn.Connected Then
        clsConn.CreateConnectionString spathVat & "\NTK_TG\"
        clsConn.Connect
    End If

    'Lay so TT to khai trong HDR
    vMaSoThue = Trim(Mid$(strData, 6, 13))

    If Len(vMaSoThue) = 13 Then
        vMaSoThue = Left(CStr(vMaSoThue), 10) & "-" & Right(CStr(vMaSoThue), 3)
    End If

    If strID = "01_TBAC" Then
        arrDeltail = Split(strData, "~")
        arrDeltail(UBound(arrDeltail) - 3) = Trim(arrDeltail(UBound(arrDeltail) - 3))

        'check neu TIN_DV_CQ = 13 thi tach
        If Len(arrDeltail(UBound(arrDeltail) - 3)) = 13 Then
            arrDeltail(UBound(arrDeltail) - 3) = Left(CStr(arrDeltail(UBound(arrDeltail) - 3)), 10) & "-" & Right(CStr(arrDeltail(UBound(arrDeltail) - 3)), 3)
        End If

        arrDate = Split(arrDeltail(UBound(arrDeltail) - 1), "/")
        dTempDate = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & _
        "Where tkhai.tin = '" & vMaSoThue & "'" & _
        "And tkhai.LOAI_BC = '" & strID & "' " & _
        " And tkhai.NGAY_BC=CTOD('" & format(dTempDate, "mm/dd/yyyy") & "')" & _
        " And tkhai.TIN_DV_CQ='" & Trim(arrDeltail(UBound(arrDeltail) - 3)) & "'"
    ElseIf strID = "03_TBAC" Then
        arrDeltail = Split(strData, "~")
        arrDate = Split(Left$(arrDeltail(UBound(arrDeltail)), 10), "/")
        dTempDate = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & " And tkhai.NGAY_BC=CTOD('" & format(dTempDate, "mm/dd/yyyy") & "')"
    ElseIf strID = "04_TBAC" Then
        arrDeltail = Split(strData, "~")
        arrDate = Split(arrDeltail(UBound(arrDeltail) - 1), "/")
        dTempDate = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        arrDate = Split(Right$(arrDeltail(UBound(arrDeltail) - 5), 10), "/")
        dTempDate1 = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & " And tkhai.NGAY_BC=CTOD('" & format(dTempDate, "mm/dd/yyyy") & "')" & " And tkhai.NGAY_TB_PH=CTOD('" & format(dTempDate1, "mm/dd/yyyy") & "')"
        
    ElseIf strID = "BC21_AC" Then
        arrDeltail = Split(strData, "~")
        arrDate = Split(Left$(arrDeltail(UBound(arrDeltail)), 10), "/")
        dTempDate = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & " And tkhai.NGAY_BC=CTOD('" & format(dTempDate, "mm/dd/yyyy") & "')"
        
    ElseIf strID = "01_AC" Then
        arrDeltail = Split(strData, "~")
        arrDate = Split(CStr(arrDeltail(1)), "/")
        dTempDate = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        arrDate = Split(Left$(arrDeltail(2), 10), "/")
        dTempDate1 = DateSerial(Val(arrDate(2)), Val(arrDate(1)), Val(arrDate(0)))
        
        strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & "And tkhai.TU_NGAY = CTOD('" & format(dTempDate, "mm/dd/yyyy") & "')" & "And tkhai.DEN_NGAY = CTOD('" & format(dTempDate1, "mm/dd/yyyy") & "')"
    ElseIf strID = "BC26_AC" Then

        If LoaiKyKK = False Then
            strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & "And tkhai.QUY_BC = CTOD('" & format$(dNgayDauKy, "mm/dd/yyyy") & "')"
        Else
            strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & "And tkhai.TU_NGAY = CTOD('" & format$(dNgayDauKy, "mm/dd/yyyy") & "')" & "And tkhai.DEN_NGAY = CTOD('" & format$(dNgayCuoiKy, "mm/dd/yyyy") & "')"

        End If

    Else
        strSQL = "select max(so_tt_tk) from tmp_bcao_hdr_ac tkhai " & "Where tkhai.tin = '" & vMaSoThue & "'" & "And tkhai.LOAI_BC = '" & strID & "' " & "And tkhai.TU_NGAY = CTOD('" & format$(dNgayDauKy, "mm/dd/yyyy") & "')" & "And tkhai.DEN_NGAY = CTOD('" & format$(dNgayCuoiKy, "mm/dd/yyyy") & "')"
    End If
    
    Set rsResult = clsConn.Execute(strSQL)
    If rsResult Is Nothing Then
        strSTT = 0
        isTonTaiAC = False
    Else
        strSTT = rsResult.Fields(0).Value + 1
        isTonTaiAC = True
    End If
    
    Set rsResult = Nothing
    clsConn.Disconnect
    getSoTTTK_AC = strSTT
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK_AC", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK_AC", Err.Number, Err.Description
End Function
Private Function getTTDLT() As Boolean

     Dim rs As ADODB.Recordset, strSQL As String
     
     If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & "\dtnt\"
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
     End If
    'vttoan: lay thong tin dai ly thue
     getTTDLT = True
     If strMST_DLT <> vbNullString Then
        strSQL = "SELECT madtnt, madlt, tengoi, dchi,dthoai, fax, email, sohd,ngayhd "
        strSQL = strSQL & " FROM DTNT_DLT where madlt = '" & strMST_DLT & "' and madtnt = '" & strMST & "'"
        Set rs = clsDAO.Execute(strSQL)
        If rs Is Nothing Then
            getTTDLT = False
            Exit Function
        Else
            strTen_DLT = rs.Fields("tengoi")
            strTen_DLT = Trim(strTen_DLT)

            strDchi_DLT = rs.Fields("dchi")
            strDchi_DLT = Trim(strDchi_DLT)

'            strQHuyen_DLT = rs.Fields("mahuyen")
'            strQHuyen_DLT = Trim(strQHuyen_DLT)

'            strTTPho_DLT = rs.Fields("matinh")
'            strTTPho_DLT = Trim(strTTPho_DLT)

            strDthoai_DLT = rs.Fields("dthoai")
            strDthoai_DLT = Trim(strDthoai_DLT)

            strFax_DLT = rs.Fields("fax")
            strFax_DLT = Trim(strFax_DLT)

            strMail_DLT = rs.Fields("email")
            strMail_DLT = Trim(strMail_DLT)

            strSoHD_DLT = rs.Fields("sohd")
            strSoHD_DLT = Trim(strSoHD_DLT)

            strNgayHD_DLT = rs.Fields("ngayhd")
            strNgayHD_DLT = Trim(format(strNgayHD_DLT, "dd/mm/yyyy"))
'        End If
'        If Not rs Is Nothing Then

'        Else
'            frmSystem.MousePointer = vbDefault
'            Me.MousePointer = vbDefault
'            Beep 600, 500
'            MessageBox "0137", msOKOnly, miCriticalError
'            LoadForm = False
'            clsDAO.Disconnect
'            Exit Function
        End If
    Else
    'set lai "" cho cac gia tri
        strTen_DLT = ""
        strDchi_DLT = ""
        strDthoai_DLT = ""
        strFax_DLT = ""
        strMail_DLT = ""
        strSoHD_DLT = ""
        strNgayHD_DLT = ""
    End If
        'end
    
    clsDAO.Disconnect
End Function
Private Function checkKyKHBS(ByVal menuId As Integer) As Boolean
    Dim vMonth As Integer
    Dim vThreeMonth As Integer
    
    vMonth = Val(TAX_Utilities_Svr_New.Month)
    vThreeMonth = Val(TAX_Utilities_Svr_New.ThreeMonths)
    
    checkKyKHBS = True
    
    'check nhung to co ky kkhai tinh theo quy'
    If menuId = 16 Or menuId = 51 Or menuId = 51 Or menuId = 11 Or menuId = 12 Or menuId = 73 Then
        If vThreeMonth < 3 And Int(TAX_Utilities_Svr_New.Year) <= 2011 Then
            checkKyKHBS = False
            Exit Function
        End If
    End If
    'check nhung to co ky kkhai tinh theo thang
    If menuId = 15 Or menuId = 50 Or menuId = 36 Or menuId = 1 Or menuId = 2 Or menuId = 4 Or menuId = 5 Or menuId = 6 Or menuId = 70 _
       Or menuId = 71 Or menuId = 72 Or menuId = 81 Or menuId = 86 Or menuId = 89 Then
        If vMonth < 7 And Int(TAX_Utilities_Svr_New.Year) <= 2011 Then
            checkKyKHBS = False
            Exit Function
        End If
    End If
    'check nhung to co ky kkhai tinh theo nam
    If menuId = 3 Or menuId = 59 Or menuId = 77 Or menuId = 80 Or menuId = 82 Or menuId = 87 Then
        If Int(TAX_Utilities_Svr_New.Year) < 2011 Then
            checkKyKHBS = False
            Exit Function
        End If
    End If
End Function
'ham check ky KHBS cho to 2 to 08
Private Function checkKyKHBSTo08(ByVal menuId As String) As Boolean
    Dim vMonth As Integer
    Dim vThreeMonth As Integer
    Dim vYear  As Integer
    
    vMonth = Val(TAX_Utilities_Svr_New.Month)
    vThreeMonth = Val(TAX_Utilities_Svr_New.ThreeMonths)
    
    checkKyKHBSTo08 = True
    
    'check nhung to co ky kkhai tinh theo quy'
    If InStr(1, UCase(menuId), "Q") > 0 Then
        If vThreeMonth < 3 And Int(TAX_Utilities_Svr_New.Year) <= 2011 Then
            checkKyKHBSTo08 = False
            Exit Function
        End If
    End If
    'check nhung to co ky kkhai tinh theo thang
    If InStr(1, UCase(menuId), "T") > 0 Then
        vMonth = Val(Mid(menuId, 2, 2))
        vYear = Val(Right(menuId, 4))
        If vMonth < 7 And vYear <= 2011 Then
            checkKyKHBSTo08 = False
            Exit Function
        End If
    End If
    'check nhung to co ky kkhai tinh theo nam
End Function

Private Function LoaiToKhai(ByVal strData As String) As Boolean
    Dim LoaiTk      As String
    Dim tmp         As String
    Dim Tk04_GTGT() As String
    On Error GoTo ErrHandle
    
    '    tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) + 5)
    '    tmp = Left$(tmp, Len(tmp) - 10)
    'LoaiTk = Right$(tmp, 1)
    LoaiTk = Mid$(strData, 4, 2)

    If LoaiTk = "71" Then
        LoaiTk = Left$(strData, Len(strData) - 10)
        Tk04_GTGT = Split(LoaiTk, "~")

        If UBound(Tk04_GTGT) > 0 Then
            LoaiTk = Tk04_GTGT(UBound(Tk04_GTGT) - 1)
            
        End If

    Else
        LoaiTk = Left$(strData, Len(strData) - 10)
        LoaiTk = Right$(LoaiTk, 1)
  
    End If
    
    If LoaiTk = "1" Then
        LoaiToKhai = True
    Else
        LoaiToKhai = False
    End If
    
    Exit Function
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "LoaiToKhai", Err.Number, Err.Description

    If Err.Number = -2147467259 Then MessageBox "0063", msOKOnly, miCriticalError
End Function

