VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmInterfaces_bb 
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterfaces_bb.frx":0000
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
         Visible         =   0   'False
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
         Left            =   8790
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
Attribute VB_Name = "frmInterfaces_bb"
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
Public KYKK_DEN_NGAY As String
Public SO_TEP As Variant
Public DHS_MA As String
Dim USER As Variant
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
    Dim rs As ADODB.Recordset, I As Long
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

    Dim strSQL As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String, strSQL_KHBS As String
    Dim HdrID As Variant, strDate() As String, dDate As Date
    Dim rs As ADODB.Recordset, I As Long
    Dim blHoiTonTai As Integer
    Dim blUpdateTHUETKY2 As Boolean
    Dim bln As Boolean
    Dim blnKTRB As Integer
    Dim sSaiCT11 As String
    Dim vKYLBO As Variant
    Dim vNGAYQUET As Variant
    Dim vNGAY_DAU_KYLBO As Variant
    Dim vTHANG_CUOI_KYKK As Variant
    Dim sSQL As String
    'Dim menuId As Integer
    Dim CHKGIAHAN As Variant
    
        
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
        clsDAO.CreateConnectionString spathVat & "\DB_HT\"
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
     End If
     menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")
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
        
        ' Ngay dau ky lap bo chua khoa so
        If Trim(vNGAY_DAU_KYLBO) = vbNullString Or Trim(vNGAY_DAU_KYLBO) = "01/../...." Then
            vNGAY_DAU_KYLBO = "CTOD('')"
        Else
            vNGAY_DAU_KYLBO = DateSerial(Int(Mid$(vNGAY_DAU_KYLBO, 7, 4)), Int(Mid$(vNGAY_DAU_KYLBO, 4, 2)), Int(Mid$(vNGAY_DAU_KYLBO, 1, 2)))
            'dhdang sua loi ky lap bo bang ky ke khai
            'ngay 20/07/2010
            If (TAX_Utilities_Svr_New.Month <> vbNullString) And (TAX_Utilities_Svr_New.Month <> "") Then
                If (Month(vNGAY_DAU_KYLBO) = Int(TAX_Utilities_Svr_New.Month)) And (Year(vNGAY_DAU_KYLBO) = TAX_Utilities_Svr_New.Year) Then
                    DisplayMessage "0120", msOKOnly, miCriticalError
                    Exit Sub
                End If
            End If
        
            vNGAY_DAU_KYLBO = "CTOD('" & format(vNGAY_DAU_KYLBO, "mm/dd/yyyy") & "')"
        End If
        
        ' Lay thang cuoi cung cua ky ke khai
        'dhdang edit
        
        If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") <> "11" And GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") <> "12" And GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") <> "03" Then
            If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
                vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
            ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
                vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
            Else
                vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
            End If
        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") = "11" Then
            .GetText .ColLetterToNumber("E"), 17, CHKGIAHAN
            If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
                    If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                         vTHANG_CUOI_KYKK = "01/02/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                         vTHANG_CUOI_KYKK = "04/05/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                         vTHANG_CUOI_KYKK = "30/07/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                         vTHANG_CUOI_KYKK = "01/11/2010"
                        End If
                    ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                         vTHANG_CUOI_KYKK = "30/07/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                         vTHANG_CUOI_KYKK = "30/11/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                         vTHANG_CUOI_KYKK = "31/01/2011"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                         vTHANG_CUOI_KYKK = "03/05/2011"
                        End If
                    Else
                        If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
                            vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
                        ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
                            vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
                        Else
                            vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
                        End If
                    End If
              Else
                    If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
                        vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
                    ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
                        vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
                    Else
                        vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
                    End If
              End If
        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") = "12" Then
            .GetText .ColLetterToNumber("E"), 17, CHKGIAHAN
            If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
                    If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                         vTHANG_CUOI_KYKK = "01/02/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                         vTHANG_CUOI_KYKK = "04/05/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                         vTHANG_CUOI_KYKK = "30/07/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                         vTHANG_CUOI_KYKK = "01/11/2010"
                        End If
                    ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
                        If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                         vTHANG_CUOI_KYKK = "30/07/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                         vTHANG_CUOI_KYKK = "30/11/2010"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                         vTHANG_CUOI_KYKK = "31/01/2011"
                        ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                         vTHANG_CUOI_KYKK = "03/05/2011"
                        End If
                    Else
                        If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
                            vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
                        ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
                            vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
                        Else
                            vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
                        End If
                    End If
            Else
                If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
                    vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
                ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
                    vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
                Else
                    vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
                End If
            End If
        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID") = "03" Then
               .GetText .ColLetterToNumber("E"), 15, CHKGIAHAN
               If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
                   If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
                       vTHANG_CUOI_KYKK = "02/11/2010"
                   ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
                       vTHANG_CUOI_KYKK = "30/06/2011"
                   End If
               Else
                    If TAX_Utilities_Svr_New.Month <> vbNullString Or TAX_Utilities_Svr_New.Month <> "" Then
                        vTHANG_CUOI_KYKK = "01/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
                    ElseIf TAX_Utilities_Svr_New.ThreeMonths <> vbNullString Or TAX_Utilities_Svr_New.ThreeMonths <> "" Then
                        vTHANG_CUOI_KYKK = "01/" & GetLastMonthOfThreeMonth(TAX_Utilities_Svr_New.ThreeMonths) & "/" & TAX_Utilities_Svr_New.Year
                    Else
                        vTHANG_CUOI_KYKK = "01/03" & "/" & Val(TAX_Utilities_Svr_New.Year) + 1
                    End If
               End If
        End If
        
        'vTHANG_CUOI_KYKK = format(vTHANG_CUOI_KYKK, "dd/mm/yyyy")
        
        vTHANG_CUOI_KYKK = DateSerial(Int(Mid$(vTHANG_CUOI_KYKK, 7, 4)), Int(Mid$(vTHANG_CUOI_KYKK, 4, 2)), Int(Mid$(vTHANG_CUOI_KYKK, 1, 2)))
        
        'CDate(vTHANG_CUOI_KYKK)
        vTHANG_CUOI_KYKK = DateAdd("M", 1, vTHANG_CUOI_KYKK)
        vTHANG_CUOI_KYKK = "CTOD('" & format(vTHANG_CUOI_KYKK, "mm/dd/yyyy") & "')"
        ' Ngay quet
        If menuId = 5 Then
            .GetText .ColLetterToNumber("T"), 12, vNGAYQUET
        ElseIf menuId = 6 Or menuId = 8 Then '01_TAIN, 03_TAIN
            .GetText .ColLetterToNumber("T"), 11, vNGAYQUET
        ElseIf menuId = 9 Then ' 02_TAIN
            .GetText .ColLetterToNumber("Q"), 11, vNGAYQUET
        ElseIf menuId = 17 Then ' 04_TNCN
            .GetText .ColLetterToNumber("L"), 12, vNGAYQUET
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
    Dim fso As New FileSystemObject
    strFileName = spathVat & "\DB_HT\" & "KHOASO.DBF"
    If fso.FileExists(strFileName) = False Then
        DisplayMessage "0111", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    Set rs = clsDAO.Execute(sSQL)
    If Not rs Is Nothing Then
        DisplayMessage "0107", msOKOnly, miInformation
        Exit Sub
    Else
        If vNGAYQUET < vNGAY_DAU_KYLBO Then
           DisplayMessage "0108", msOKOnly, miInformation
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
            If vNGAY_DAU_KYLBO > vTHANG_CUOI_KYKK Then
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
       Insert_QHS
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
    Dim I As Integer
    With fpSpread1
        .Visible = False
        .ReDraw = False
        .EditMode = False
        iActiveSheet = .ActiveSheet
        lActiveCol = .ActiveCol
        lActiveRow = .ActiveRow
        
        
        For I = 1 To .SheetCount
            .ActiveSheet = I
            .Sheet = .ActiveSheet
            .Row = 1
            .Col = 1
            .Lock = False
            .SetActiveCell 1, 1
            .EditMode = True
        Next

        For I = 1 To .SheetCount
            .ActiveSheet = I
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
    Dim I, j, t, counter As Integer
    Dim chkToKhai As Boolean
    
    For I = 1 To UBound(arrBCBuffer)
        If arrBCBuffer(I) <> vbNullString Then
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
    
    ' 01/GTGT thu trong phien ban HTKK 2.1.0
'    str1 = "aa250010101724672   04201000100100100101/0114/06/2006<S01><S>~1110~1110~0~322330~32230~0~0~2320~0~0~0~0~2320~0~0~0~0~2320~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0</S></S01>"
'    Barcode_Scaned str1
'
    'str1 = "aa250010101724672   03201000100100100101/0114/06/2006<S01><S>~0~154225000000~15422500~154225000000~15422500~0~0~0~0~0~0~15422500~2000000~2615000000~18500000~15000000~2600000000~18500000~400000000~700000000~3500000~1500000000~15000000~0~0~0~0~2615000000~18500000~16500000~0~0~0</S></S01>"
'    str1 = "aa250010101724672   03201000100100100101/0114/06/2006<S01><S>~0~150000000~1500000~150000000~1500000~0~0~0~0~0~0~1500000~200000~965000000~72500000~200000000~765000000~72500000~15000000~50000000~2500000~700000000~70000000~0~0~0~0~965000000~72500000~72300000~0~0~0</S></S01>"
'    Barcode_Scaned str1
'    str1 = "aa210172300103987   00200800300600100201/0101/01/190001/04/200831/12/2008<S04><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~60000000000~10000000000~50000000000~1250000000~250000000~1000000000~0~"
'    Barcode_Scaned str1
'    str2 = "aa210172300103987   0020080030060020020~0~60000000000~10000000000~50000000000~1250000000~250000000~1000000000~0~0~0~1250000000~250000000~1000000000</S><S>~~0~0~~</S><S>~~0~0~~</S><S>~06/08/2009</S></S04>"
'    Barcode_Scaned str2

' 02/GTGT thu trong phien ban HTKK 2.1.0
'    str1 = "aa250020101724672   03201000100100100101/0114/06/2006<S01><S>NGUYEN~10000000~0~0~0~0~0~0~0~0~0~0~0~10000000~0~0~10000000</S></S01>"
'    Barcode_Scaned str1

'01TBH/TNCN
'   str1 = "aa250460101724672   02201000300300100101/0101/01/2010<S01><S>100000000~2000000~300000</S><S>~~1~~</S></S01>"
'   Barcode_Scaned str1
'01ATADB
'   str1 = "aa250050101724672   03201000100100100101/0101/01/2010<S01><S>~210000000~127272727.00~0~0~82727273</S><S>10103~Kg~3000.00~210000000~127272727.00~65.0~0~0~82727273</S><S>1000~833.00~0~0~167</S><S>20400~~1100.00~1000~833.00~20.0~0~0~167</S><S>11111</S><S>10203~LÝt~111.00~11111</S><S>~~0.00~0</S><S>~~0.00~0</S><S>210012111~127273560.00~0~0~82727440</S></S01>"
'   Barcode_Scaned str1


'01QBH/TNCN
'   str1 = "aa250470101724672   01201000100300100101/0101/01/2010<S01><S>10000000000~200000000~30000000</S><S>~~1~~</S></S01>"
'   Barcode_Scaned str1
   
'01TXS/TNCN
'   str1 = "aa250480101724672   01201000100200100101/0101/01/2010<S01><S>100000000~2000000~400000</S><S>~~1~~</S></S01>"
'   Barcode_Scaned str1
'01QXS/TNCN
'   str1 = "aa250490101724672   01201000100100100101/0101/01/2010<S01><S>100000000~2000000~63000</S><S>~~1~~</S></S01>"
'   Barcode_Scaned str1
'
'023/TNDN
'   str1 = "aa250030101724672   00200901201200100101/0114/06/200601/01/200931/12/2009<S03><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~~~~~~~x~NguyÔn V¨n An~21/04/2010</S></S03>"
 '  Barcode_Scaned str1
   
'02B/TNCN

'    str1 = "aa250160101724672   02201000200200100101/0101/01/2010<S01><S>100000000~10000000~100000~200000000~20000000~2000000~30000~6000</S><S>~~1~~</S></S01>"
'    Barcode_Scaned str1


'03B/TNCN
'      str1 = "aa250510200765535   01201000100100100101/0101/01/2010<S01><S>100000000~5000000~200000000~200000~300000000~15000000~400000000~40000000~500000000~0</S><S>~~1~~</S></S01>"
'      Barcode_Scaned str1
'03A/TNCN
'     str1 = "aa250500101724672   02201000200200100101/0101/01/2010<S01><S>123000000~5000000~200000000~200000~300000000~15000000~400000000~40000000~500000000~0</S><S>NguyÔn V¨n An~08/03/2010~1~~</S></S01>"
'     Barcode_Scaned str1
'04A/TNCN
    'str1 = "aa250390101724672   02201000400400100101/0101/01/2009<S01><S>111111111112~111111111111~1~22222222224~22222222222~2~3333333~3333333~0</S><S>~~1~~</S></S01>"
    'Barcode_Scaned str1
'04B/TNCN
     'str1 = "aa250400200765535   01201000200200100101/0101/01/2009<S01><S>3~1~2~700000000~300000000~400000000~50000006~50000000~6</S><S>~~1~~</S></S01>"
     'Barcode_Scaned str1
'07/TNCN
'     str1 = "aa250360101724672   02201000700700100101/0101/01/2010<S07><S>100000000~904000000~4000000~200000000~300000000~400000000~0~0~500000~0~0~x</S><S>NguyÔn V¨n An~05/03/2010~1~~</S></S07>"
'     Barcode_Scaned str1
    
    
    
'01 GTGT
'  blnBangke = False
   'str1 = "aa210012300103987   08200900400400100701/0114/06/2006<S01><S>~0~6100000~115000~6100000~115000~0~0~0~0~0~0~115000~1000000~24000000~1250000~1000000~23000000~1250000~8000000~5000000~250000~10000000~1000000~0~100000~0~11000000~24000000~-9650000~0~10650000~0~10650000</S></S01>"
   'Barcode_Scaned str1
'   str2 = "aa130012300103987   062008004004002007<S01_1><S>00001~2000000~01/01/2008~ten nguoi mua~~mat hang 1~1000000~7~780707~</S><S>0002~00002~04/02/2008~ten nguoi mua 1~~mat hang 2~2000000~0%~78707~~01100~00003~01/02/2008~ten nguoi mua 11~~mat hang 22~6000000~0%~0~</S><S>000003~00003~01/01/2008~ten nguoi mua 2~~mat hang 3~5000000~5%~250000~</S><S>00004~00004~01/01/2008~ten nguoi mua 2~~mat hang 4~10000000~10%~1000000~</S><S>24000000~2109414</S></S01_1>"
'   Barcode_Scaned str2
'   str3 = "aa130012300103987   062008004004003007<S01_2><S>1~1~01/01/0008~ha the phuong~1100201110~hat hang~5000000~0.00~0~</S><S>2~2~01/01/2008~ha the phuong~~hat hang1~500000~0.00~60000~~45735~4573~01/02/2008~ha the phuong~~~100000~5.00~5000~</S><S>634634~34636~01/02/2008~ha the phuong~~~500000~10.00~50000~</S><S>6100000~115000</S></S01_2>"
'   Barcode_Scaned str3
'   str4 = "aa130012300103987   062008004004004007<S01_3><S>01/2007~01/01/2008~100000~0~12/2007~01/02/2007~0~6000000</S></S01_3>"
'   Barcode_Scaned str4
'   str5 = "aa130012300103987   062008004004005007<S01_4A><S>9000000~1000000~3000000~5000000~100000000~60000000~60.00~5000000~3000000</S></S01_4A>"
'   Barcode_Scaned str5
'   str6 = "aa130012300103987   062008004004006007<S01_4B><S>2007~6600000~6000000~500000~100000~1500000~1600000~106.67~100000~106670~12000~94670</S></S01_4B>"
'   Barcode_Scaned str6
'   str7 = "aa130012300103987   062008004004007007<S01_5><S>000001~01/01/2000~Ha noi~hai phong~5000000</S></S01_5>"
'   Barcode_Scaned str7

'01/GTGT 1.3.1

'   str1 = "aa131012300103987   07200801201300100401/0114/06/2006<S01><S>~1000000~1593000~132750~1593000~132750~0~0~0~0~0~0~132750~0~114911565~7578176~21109070~93802495~7578176~479980~35081520~1754076~58240995~5824100~0~0~0~0~114911565~7578176~6578176~0~0~0</S></S01>"
'   Barcode_Scaned str1
'   str2 = "aa131012300103987   072008012013002004<S01_1><S>YR2007N~0171886~15/04/2008~Cty TNHH Uni- President VN~3700306630~M× Vôn ~21109070~~0~</S>"
'   str2 = str2 & "<S>YR2007N~170789~14/02/2008~T &amp; F IMPORT EXPORT CO., LTD~~Thöïc aªn Gia Suïc~436346~0%~0~~YR2007N~171722~25/02/2008~T &amp; F IMPORT EXPORT CO., LTD~~Thöïc aªn Gia Suïc~43634~0%~0~</S><S>XV2007N~140694~01/02/2008~Hoµng V¨n H¹nh~"
'   Barcode_Scaned str2
'   str3 = "aa131012300103987   072008012013003004~Thøc ¨n Gia sóc~15881520~5%~794080~~XV2007N~140695~01/02/2008~ Hå ThÞ Nga~~Thøc ¨n Gia sóc~19200000~5%~960000~</S><S>XV2007N~140548~05/02/2008~§inh V¨n T©m~~M× ¨n liÒn~54193450~10%~5419341~~XV2007N~140549~11/02/2008~Cty TNHH METROCASH &amp; CARRY VN ~0302249586~M× ¨n liÒn~4047545~10%~4454664~</S><S>114911565~7578176</S></S01_1>"
'   Barcode_Scaned str3
'   str4 = "aa131012300103987   072008012013004004<S01_2><S>XX/2007N~0000110193~26/01/2008~Cty TNHH TM &amp; DV Ñöïc Tua¸n~0302696785-001~Daµu DO~531000~10%~53100~49/6 QL1A, Ño©ng Höng Thuaän, Q.12, TPHCM~XG/2007N~70232~28/01/2008~DNTN Nam Thaïi~3700145486001~Daµu DO~531000~5%~26550~Huyeän Thuaän An, Tænh B×nh Dö«ng~XX/2007N~110302~29/01/2008~Cty TNHH TM &amp; DV Ñöïc Tua¸n~0302696785-001~Daµu DO~531000~10%~53100~49/6 QL1A, Ño©ng Höng Thuaän, Q.12, TPHCM</S><S>~~~~~~0~~0~</S><S>~~~~~~0~~0~</S><S>1593000~132750</S></S01_2>"
'   Barcode_Scaned str4




'    str1 = "aa200172300103987   00200800100100105001/0101/01/190001/01/200831/12/2008<S04><S>0~362~0~0~362~0~28789059299~28789059299~0~714814570~714814570~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0"
'    Barcode_Scaned str1
'    str2 = "aa200172300103987   002008001001002050~0~0~28789059299~28789059299~0~714814570~714814570~0~0~0~0~714814570~714814570~0</S><S>~~0~0~~</S><S>~~0~0~~</S><S>Hµ ThÕ Ph­¬ng~02/03/2009</S></S04>"
'    Barcode_Scaned str2
'
'
'
'    str1 = "aa200172300103987   0020080010010060507777306~6758910~175891~2110692~1847000~263692~~NguyÔn Minh Dòng~~73309373~26347622~1800000~45161751~6109114~110911~1330932~1122000~208932~~Lª Hoµng Trung ~~74656240~27653184~1200000~45803056~6221353~122135~1465620~1280000~185620~~TrÇn Ngäc Lanh ~~70469821~25589093~1800000~43080728~5872485~87249~1046988~1061000~-14012~~§ç Thanh Liªm~~71574506~25624214~800000~45150292~5964542~96454~1157448~1083000~74448~~NguyÔn Ph­íc Toµn~2100349221~81184615~29421605~800000~50963010~6765385~176539~2118468~1658000~460468~~Lª ThÞ KiÒu Th¾m ~~70063789~26189974~1050000~42823815~5838649~83865~1006380~1050000~-43620~~TrÇn Khëi NghÜa ~2100349239~96295426~35641469~700000~59953957~8024619~302462~3629544~2741000~888544~~L"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100100505000~79749441~10921856~592186~7106232~5039000~2067232~~D­¬ng Hoµng~2100261707~93320880~37002034~700000~55618846~7776740~277674~3332088~2784000~548088~~TrÇn ThÕ Ph­¬ng ~2100271293~82844849~33773821~700000~48371028~6903737~190374~2284488~1890000~394488~~Vâ Hång Khanh~2100271374~108698326~40010525~3050000~65637801~9058194~405819~4869828~3511000~1358828~~Qu¸ch H¶i Hå~2100271381~111022168~38641235~3300000~69080933~9251847~425185~5102220~3636000~1466220~~Ng« Quèc Phong~2100271399~62069083~27905002~700000~33464081~5172424~17242~206904~322000~-115096~~Cao NguyÔn Quèc H­ng~2100271416~63954113~26984898~700000~36269215~5329509~32951~395412~470000~-74588~~L©m Thanh T©m ~2100349253~81106925~32529619~800000~4"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001004050444~1865328~1683000~182328~~§Æng Huy Hoµng ~2100261714~98950157~39746921~1200000~58003236~8245846~324585~3895020~3092000~803020~~NguyÔn T Thanh TuyÒn~~63013472~23535287~800000~38678185~5251123~25112~301344~617000~-315656~~Hå ThÞ Hång Ph­¬ng~~60174692~22565198~880000~36729494~5014558~1456~17472~302000~-284528~~Hå Minh TuÊn~~66061376~25243573~700000~40117803~5505115~50512~606144~858000~-251856~~Phan Minh Nh©n ~~64326323~25217439~870000~38238884~5360527~36053~432636~474000~-41364~~NguyÔn V¨n D­îc~~62715725~27696360~700000~34319365~5226310~22631~271572~452000~-180428~~Lª V¨n Lao~~65630096~27645222~940000~37044874~5469175~46918~563016~614000~-50984~~Huúnh ChÝ H¶i~2100261915~131062269~47112828~42000"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001003050<S04-1><S>NguyÔn Thanh Long~2100261601~235145486~81515073~25900000~127730413~19595457~1919091~23029092~16718000~6311092~~§Æng V¨n D×nh~2100261619~203938855~71631919~19200000~113106936~16994905~1398981~16787772~12944000~3843772~~NguyÔn Thanh Hoµng~2100261640~179815267~67320475~17600000~94894792~14984606~998461~11981532~10126000~1855532~~Huúnh VÜnh Phóc~2100349214~106999770~34846293~17900000~54253477~8916648~391665~4699980~4247000~452980~~Lª ChÝ H­íng~2100271261~69380638~29013045~830000~39537593~5781720~78172~938064~976000~-37936~~Ph¹m Thanh Tïng~2100262059~84912096~36072366~700000~48139730~7076008~207601~2491212~2112000~379212~~T¹ ThÞ YÕn~2100271279~78653245~30537899~850000~47265346~6554437~155"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001010050~Phan Minh S¬n~2100261778~117932342~46505076~2800000~68627266~9827695~482770~5793240~4197000~1596240~~NguyÔn.T.Ph­¬ng §«ng~2100261785~90469645~35518321~700000~54251324~7539137~253914~3046968~2372000~674968~~L©m.T. Ngäc Xu©n~2100261802~87009555~33305935~1110000~52593620~7250796~225080~2700960~2182000~518960~~NguyÔn V¨n Tr­êng~2100261827~92534880~36305149~810000~55419731~7711240~271124~3253488~2752000~501488~~Mai BÝch Dung~2100271430~70227969~28686737~700000~40841232~5852331~85233~1022796~1023000~-204~~NguyÔn Ngäc HiÖp~2100271448~84491542~33850754~800000~49840788~7040962~204096~2449152~2037000~412152~~Huúnh V¨n BÐ~2100271455~82515118~32529619~1050000~48935499~6876260~187626~2251512~1942000~30951"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001009050370862~4450344~3386000~1064344~~HÇu V¨n §øc ~2100271303~78890810~30376878~800000~47713932~6574234~157423~1889076~1746000~143076~~D­¬ng ThÞ Méng Lµnh ~2100271328~79775542~31438448~930000~47407094~6647962~164796~1977552~1742000~235552~~TrÇn Thµnh Nam ~2100261721~86372245~33606316~700000~52065929~7197687~219769~2637228~2160000~477228~~Lª V¨n Ch¸nh ~~73459761~30683256~950000~41826505~6121647~112165~1345980~1163000~182980~~Ng« Thµnh Duy Bót ~~77778252~30105363~700000~46972889~6481521~148152~1777824~1703000~74824~~TrÇn V¨n Thµnh ~~67138320~25734565~700000~40703755~5594860~59486~713832~914000~-200168~~Lª ViÖt Dòng~2100261760~122510520~46749869~4100000~71660651~10209210~520921~6251052~4544000~1707052~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100100805056~3000000~40198899~5475021~47502~570024~582000~-11976~~L­¬ng ViÖt Lang ~2100261658~107067499~40993590~3800000~62273909~8922292~392229~4706748~3749000~957748~~TrÇn M¹nh HiÒn ~2100271342~72353831~28972009~800000~42581822~6029486~102949~1235388~1151000~84388~~NguyÔn ThÕ TuyÒn ~2100261746~115993660~46156891~800000~69036769~9666138~466614~5599368~4166000~1433368~~Bïi V¨n Vò Em ~~64765171~25296412~800000~38668759~5397098~39710~476520~739000~-262480~~TrÇn V¨n NhuÇn~~73785422~29307775~800000~43677647~6148785~114879~1378548~1251000~127548~~NguyÔn Hoµng Minh ~~64865326~25821712~700000~38343614~5405444~40544~486528~736000~-249472~~NguyÔn Quang Anh ~2100271310~104503427~40364997~3900000~60238430~8708619~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001007050©m N TrÇn Quang~2100348348~93884945~35936438~700000~57248507~7823745~282375~3388500~2548000~840500~~Phan Thµnh §øc ~~65433900~26997272~700000~37736628~5452825~45283~543396~564000~-20604~~Lª Duy Ph­¬ng~~72699872~25756192~800000~46143680~6058323~105832~1269984~1185000~84984~~TrÇn §×nh Khang~~68487553~25715987~700000~42071566~5707296~70730~848760~718000~130760~~Ph¹m H÷u §ång ~2100271335~112558164~40287606~6600000~65670558~9379847~437985~5255820~3845000~1410820~~Lª Thanh TuÊn~2100271568~91780551~32992319~2550000~56238232~7648379~264838~3178056~2423000~755056~~NguyÔn Thµnh Trung ~2100349172~100477503~37205880~4100000~59171623~8373125~337313~4047756~2856000~1191756~~Huúnh TuÊn KiÖt ~~65700255~225013"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010140502000~1757416~~§Æng NhËt Khiªm ~2100349278~101825911~38117882~800000~62908029~8485493~348549~4182588~2981000~1201588~~TrÇn Minh ThiÖn~~70190646~25979835~700000~43510811~5849221~84922~1019064~1016000~3064~~Ng« ThÕ H­¬ng ~2100349207~103972724~40689019~3800000~59483705~8664394~366439~4397268~3456000~941268~~Huúnh §øc Quang ~2100271582~86723448~34563960~1250000~50909488~7226954~222695~2672340~2129000~543340~~Ng« V¨n TuÊn~~77659181~30567993~800000~46291188~6471598~147160~1765920~1684000~81920~~NguyÔn Thµnh Th¸i ~~77278306~30780575~700000~45797731~6439859~143986~1727832~1407000~320832~~§inh Hoµng Dòng~2100262034~108377545~42938737~1880000~63558808~9031462~403146~4837752~3440000~1397752~~Lª V¨n ThuËn~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001013050112284~~§inh C«ng B×nh~2100261947~117538114~46248435~700000~70589679~9794843~479484~5753808~4128000~1625808~~TrÇn Thanh V©n~2100261954~118994257~46575927~700000~71718330~9916188~491619~5899428~4249000~1650428~~NguyÔn H÷u B×nh ~2100271631~108331577~42084354~700000~65547223~9027631~402763~4833156~3612000~1221156~~TrÇn Ngäc HËu~2100271617~109868769~43160604~700000~66008165~9155731~415573~4986876~3645000~1341876~~L©m Kú Dao~2100271624~112551643~44493874~700000~67357769~9379304~437930~5255160~3784000~1471160~~TrÇn ViÕt Vinh~2100349260~107306969~41136399~950000~65220570~8942247~394225~4730700~3546000~1184700~~NguyÔn L Nhùt T©n~2100271575~120494175~48053186~700000~71740989~10041181~504118~6049416~429"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010120507230~800000~46948684~6699660~169966~2039592~1586000~453592~~Lª Minh TrÝ~~60640593~21592483~800000~38248110~5053383~5338~64056~322000~-257944~~NguyÔn §×nh Thµnh ~2100271504~80799161~30193251~700000~49905910~6733263~173326~2079912~1647000~432912~~Lª V¨n Hïng~2100262080~100004859~35958055~1170000~62876804~8333738~333374~4000488~3153000~847488~~L©m Tr­êng ThÞnh ~~65099335~23516110~700000~40883225~5424945~42495~509940~594000~-84060~~§Æng TÊn Duy ~~65098798~23300735~800000~40998063~5424900~42490~509880~596000~-86120~~Lª ThÞ Mü Dung ~~85621014~31120514~1830000~52670500~7135085~213509~2562108~2093000~469108~~Vâ Thanh Hïng~2100261922~136152796~50085869~4100000~81966927~11346066~634607~7615284~5503000~2"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010110502~~Cao ThÞ HiÒn ~~76576421~30554116~800000~45222305~6381368~138137~1657644~1373000~284644~~Huúnh K Minh ChuÈn~~67044173~26185490~1000000~39858683~5587014~58701~704412~985000~-280588~~NguyÔn M¹nh Hïng~2100271529~82263233~31139889~3700000~47423344~6855269~185527~2226324~1406000~820324~~Vâ ThÞ Trang Chi~2100271550~86848170~34563960~1050000~51234210~7237348~223735~2684820~2157000~527820~~Bïi V¨n ViÖt Hoµ~2100271536~84672699~34338289~900000~49434410~7056058~205606~2467272~2036000~431272~~Kiªn Thanh Phong ~~64190905~22888247~700000~40602658~5349242~34924~419088~510000~-90912~~§inh.H.Lª Quèc Dòng~~75092617~27122059~700000~47270558~6257718~125772~1509264~1302000~207264~~NguyÔn Thanh T©m~~80395914~3264"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001018050¶o Quèc~~71585904~20352684~896000~50337220~5965492~96549~1158588~1137357~21231~~NguyÔn V¨n T Tó~~65694065~19717387~760000~45216678~5474505~47451~569412~840285~-270873~~Tèng Hoµng Vò~~62138795~18088274~1296000~42754521~5178233~17823~213876~364004~-150128~~D­¬ng Hïng~2100349285~95855321~29062945~1715777~65076599~7987943~298794~3585528~2607260~978268~~NguyÔn Träng Phóc~2100271649~88717391~27720941~1046000~59950450~7393116~239312~2871744~2890559~-18815~~NguyÔn Kh¶I Hoµng~~67656830~19849911~796000~47010919~5638069~63807~765684~885814~-120130~~Hå Thanh Hïng~~60763691~16922927~892307~42948457~5063641~6364~76368~319054~-242686~~NguyÔn V¨n Gol~~62446755~17406140~1128307~43912308~5203896~20390~244680~40"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001017050460~183046~2196552~1815861~380691~~Tr­¬ng V¨n TuÊn~2100261841~99311282~32475697~896000~65939585~8275940~327594~3931128~3052531~878597~~Ph¹m Thµnh VÞ~2100261898~111389585~32264557~896000~78229028~9282465~428247~5138964~4025109~1113855~~Ph¹m v¨n H¶I~2100261908~96122226~26186637~826000~69109589~8010186~301019~3612228~3095664~516564~~Th«i TrÇn A TuÊn~~66621185~19686717~796000~46138468~5551765~55177~662124~850645~-188521~~Phan Minh §øc~2100349292~81310390~23540648~986000~56783742~6775866~177587~2131044~1854989~276055~~V­¬ng ChÝ T©m~~65992242~20067497~896000~45028745~5499354~49935~599220~809859~-210639~~NguyÔn Thanh Long~~61808476~19890203~796000~41122273~5150706~15071~180852~378342~-197490~~Huúnh B"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001016050021946~30557300~800000~63664646~7918496~291850~3502200~2844468~657732~~D­¬ng V¨n Ngé~~62878831~19613259~796000~42469572~5239903~23990~287880~613201~-325321~~NguyÔn §×nh Dòng~2100349327~82392585~27465854~796000~54130731~6866049~186605~2239260~1874605~364655~~Mai Hång Cóc~2100349302~81564909~25599010~800000~55165899~6797076~179708~2156496~1893755~262741~~Hå V¨n Th¾ng~2100261993~99504889~32475697~1000000~66029192~8292074~329207~3950484~3059688~890796~~Tr­¬ng Xu©n Vâ~~73201174~23356165~796000~49049009~6100098~110010~1320120~1171295~148825~~Vâ Minh Phong~2100271367~85272185~26910138~1066000~57296047~7106015~210602~2527224~2131487~395737~~Lª V¨n HiÓn~2100271462~81965515~24187013~796000~56982502~6830"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010150502100262066~103166817~40758261~700000~61708556~8597235~359724~4316688~3118000~1198688~~TrÇn H÷u LËp~2100262073~103691091~40983112~1060000~61647979~8640924~364092~4369104~3151000~1218104~~Lª V¨n Chµo~2100262098~94393985~36853980~1080000~56460005~7866165~286617~3439404~2617000~822404~~Trang V¨n §Ó~2100261834~152778931~46158201~4506529~102114201~12731578~773158~9277896~6731338~2546558~~NguyÔn Ph­íc HËu~2100261672~108924161~35483183~1284000~72156978~9077013~407701~4892412~3628315~1264097~~Huúnh TÊn Ph¸t~~61000520~20467647~796000~39736873~5083377~8338~100056~718101~-618045~~Liªu ThÞ T H­¬ng~2100271487~78318204~24347702~800000~53170502~6526517~152652~1831824~1668824~163000~~TrÇn ThÞ Yªm~2100261866~95"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010220509~6980995~198100~2377200~1989623~387577~~D­¬ng V¨n NghØnh~2100271783~74861984~22823598~916000~51122386~6238499~123850~1486200~1224101~262099~~NguyÔn V¨n Ný~2100271800~75032155~23263308~926000~50842847~6252680~125268~1503216~1315989~187227~~§ç Hïng C­êng~~65260428~19717387~730000~44813041~5438369~43837~526044~809860~-283816~~Lý Anh TuÊn~~74274065~23275809~1042000~49956256~6189505~118951~1427412~1462551~-35139~~Lª Thanh §iÒn~~73305515~23474914~922000~48908601~6108793~110879~1330548~1421867~-91319~~Th¸i Quèc §¹t~~70557390~21388743~1016000~48152647~5879783~87978~1055736~1060883~-5147~~§µo Phan Quèc Anh~~72414426~21476726~964000~49973700~6034536~103454~1241448~1154707~86741~~NguyÔn ViÖt Toµn~~75812"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100102105044000~78958658~9843532~484353~5812236~4194767~1617469~~L­¬ng Vò Lang~2100271663~83349782~27098851~796000~55454931~6945815~194582~2334984~2030541~304443~~Lª V¨n Phông~2100262027~96597468~31808760~934000~63854708~8049789~304979~3659748~2824603~835145~~Lª V¨n Ký~2100271670~88677083~29009083~956000~58712000~7389757~238976~2867712~2155945~711767~~NguyÔn TÊn Hïng~2100271695~85381083~28014602~1112000~56254481~7115090~211509~2538108~2133502~404606~~Phan NhËt Thanh~2100271705~84493579~27586015~1064000~55843564~7041132~204113~2449356~2012127~437229~~Huúnh Phóc Vò~2100271744~80860131~24944171~1016000~54899960~6738344~173834~2086008~1596993~489015~~Huúnh V¨n T­~2100271744~83771941~25988182~1592000~5619175"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001020050 Kh¸nh~2100271751~99586263~31250874~2419777~65915612~8298855~329886~3958632~2800862~1157770~~D­¬ng KÝnh~2100262228~120793486~36526839~1360000~82906647~10066124~506612~6079344~4529007~1550337~~Th¹ch BuneThuone~~63767015~18974289~896000~43896726~5313918~31392~376704~508330~-131626~~NguyÔn Thµnh Nh©n~~78245975~22471494~996000~54778481~6520498~152050~1824600~1489457~335143~~Lª Quèc Kh¸nh~~63387531~20218364~944000~42225167~5282294~28229~338748~479596~-140848~~NguyÔn V¨n ChÝnh~~72679378~20928854~896000~50854524~6056615~105661~1267932~1155689~112243~~NguyÔn Thanh H¶I~2100261979~109235101~38734591~1112000~69388510~9102925~410293~4923516~3299153~1624363~~NguyÔn V¨n §«ng~2100261986~118122379~38019721~11"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010190506140~-161460~~Lª V¨n Hïng~~68509612~20049601~1086000~47374011~5709134~70913~850956~954907~-103951~~Th¹ch Quan Tiªn~~71529592~20663355~1002000~49864237~5960799~96080~1152960~1069414~83546~~Ph¹m Hoµng Kha~2100271769~67074807~21006239~896000~45172568~5589567~58957~707484~852487~-145003~~Phan ThÞ Thu Duyªn~2100271286~85631089~25265253~700000~59665836~7135924~213592~2563104~2139394~423710~~Bïi ThÞ Thu Hµ~2100261792~90885061~28498062~1200000~61186999~7573755~257376~3088512~2584933~503579~~NguyÔn Thanh D©n~2100271494~87481734~28220190~1697558~57563986~7290145~229014~2748168~2159820~588348~~D­¬ng HiÕu NghÜa~2100261880~97439524~32701231~1000000~63738293~8119960~311996~3743952~2618596~1125356~~§Æng Phóc"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001026050247700~-120416~~Lª Quang Vinh~~93395354~39731798~3424630~50238926~7782947~278295~3339540~2137773~1201767~~§ç Ch©u Tïng~~70257360~26788619~2961852~40506889~5854781~85478~1025736~988311~37425~~NguyÔn V¨n ThËt~~62032114~24037877~2478587~35515650~5169343~16934~203208~300598~-97390~~§Æng V¨n ThuËn~~76369509~31401368~2597926~42370215~6364126~136413~1636956~1562982~73974~~NguyÔn V¨n §¼ng~~76690942~31157871~3131852~42401219~6390912~139091~1669092~1555207~113885~~D­¬ng Thanh S¬n~~77438396~31333898~3391852~42712646~6453200~145320~1743840~1601437~142403~~TrÇn Thanh Phong~~74773030~31043216~2938546~40791268~6231086~123109~1477308~1238659~238649~~NguyÔn T©n C­¬ng~~70076565~28038932~2929203~39108430~5839714"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001025050261~-151049~~Bïi Huy Vò~~67790718~24055475~5709484~38025759~5649227~64923~779076~614774~164302~~T¹ H÷u Toµn~2100261873~66208116~26021232~568436~39618448~5517344~51734~620808~1187656~-566848~~Ph¹m V¨n MicSol~~66225787~25889536~1666739~38669512~5518816~51882~622584~878165~-255581~~Thi Quèc Dòng~~68422743~26009121~4233328~38180294~5701896~70190~842280~683716~158564~~Huúnh B¶o Huy~~78326796~28380343~3587771~46358682~6527234~152723~1832676~1328795~503881~~NguyÔn Hoµn CÇu~~67152238~27031284~4339871~35781083~5596020~59602~715224~595248~119976~~Lª Quèc TuÊn~~67372168~26137108~3792567~37442493~5614348~61435~737220~606296~130924~~NguyÔn Hång Ngäc~~61272790~22911864~4170099~34190827~5106066~10607~127284~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100102405065~86111382~23049470~988000~62073912~7175949~217595~2611140~1990960~620180~~Ph¹m H÷u Léc~2100261961~107772371~42036225~7974471~57761675~8981031~398103~4777236~3229042~1548194~~Trang §«ng H¹~~77586964~31508008~2294699~43784257~6465581~146558~1758696~1405392~353304~~D­¬ng TuÊn Khanh~~62181799~24166903~1281112~36733784~5181817~18182~218184~603964~-385780~~T¹ Hoµng Dòng~2100261697~96304592~37460492~4908373~53935727~8025383~302538~3630456~2848571~781885~~§Æng Duy Th¸i~~69866034~26359275~3374788~40131971~5822170~82217~986604~955149~31455~~TrÇn Thanh Hé~~65797462~23839422~3188367~38769673~5483122~48312~579744~542830~36914~~NguyÔn Thµnh Dòng~~65412100~25749354~2760425~36902321~5451009~45101~541212~692"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001023050067~24092646~1064000~50655421~6317672~131767~1581204~1554127~27077~~NguyÔn Quèc Hïng~~74677402~23921779~964000~49791623~6223117~122312~1467744~1495959~-28215~~TrÇn V¨n Nhí~~71139288~22121535~1194000~47823753~5928274~92827~1113924~1082938~30986~~Huúnh Quèc Du~~70000030~21351991~1094000~47554039~5833336~83334~1000008~1034185~-34177~~Bïi Minh Quang~~63902481~19504305~944000~43454176~5325207~32521~390252~464108~-73856~~Viªn VÜnh Lîi~~65463867~19465867~1026000~44972000~5455322~45532~546384~543488~2896~~Phan TuÊn Khanh~~62382113~19818704~892000~41671409~5198509~19851~238212~445204~-206992~~NguyÔn Duy C­êng~~61425129~18619720~996000~41809409~5118761~11876~142512~375564~-233052~~D­¬ng V¨n H¶i~21002616"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001030050~~77735062~30077560~1094000~46563502~6477922~147792~1773504~998148~775356~~Phan V¨n Phóc~~78160618~30325680~1275000~46559938~6513385~151338~1816056~649520~1166536~~TrÇn Thanh H¶i~~65284623~22273926~1990000~41020697~5440385~44039~528468~368110~160358~~Vâ Quèc Hïng~~65111405~23458599~950000~40702806~5425950~42595~511140~309560~201580~~NguyÔn Hoµng T©n~~63004449~22386207~800000~39818242~5250371~25037~300444~251014~49430~~§Æng V¨n Long~~70854593~27312361~1114000~42428232~5904549~90455~1085460~438715~646745~~Th¹ch Kim U«l~~61022314~23287324~890000~36844990~5085193~8519~102228~185306~-83078~~Phan V¨n ThÕ~~61123011~23398662~930000~36794349~5093584~9358~112296~187437~-75141~~Tr­¬ng Minh HËu~~69725654~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001029050~31013965~920000~47348252~6606851~160685~1928220~650884~1277336~~NguyÔn Thanh Hïng~~70664418~25807862~1567896~43288660~5888702~88870~1066440~458390~608050~~S¬n B¶y~2100272201~79026265~31470293~994000~46561972~6585522~158552~1902624~995913~906711~~Th¸i V¨n Biªn~~72544357~25993306~1640821~44910230~6045363~104536~1254432~513470~740962~~S¬n Ph­¬ng §¹t~~72317942~29430657~1034000~41853285~6026495~102650~1231800~478918~752882~~S¬n Thµnh Quang~~64728546~25511312~994000~38223234~5394046~39405~472860~297486~175374~~Chung TiÕn Vò~~82987611~33253804~1050000~48683807~6915634~191563~2298756~764027~1534729~~Kiªn S« §­îc~~77657558~29948334~1034000~46675224~6471463~147146~1765752~1002641~763111~~Mai Trung Kiªn"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001028050388~37868762~4038000~64194626~8841782~384178~4610136~1838738~2771398~~NguyÔn Vò Tr­êng~2100272096~81133254~32724543~994000~47414711~6761105~176110~2113320~1049630~1063690~~L©m TÊn HiÖp~2100272113~93793811~34276169~1985000~57532642~7816151~281615~3379380~1461683~1917697~~Liªu Kinh H¹n~2100272138~75905204~29257937~940000~45707267~6325434~132543~1590516~612033~978483~~Ph¹m B¸ Dáng~2100272145~72247904~25237480~1150000~45860424~6020659~102066~1224792~508883~715910~~TrÇn V¨n HiÕu~2100272152~79095707~30318558~1538731~47238418~6591309~159131~1909572~663776~1245796~~NguyÔn Hoµng T¨ng~2100272177~80572848~30937708~1100000~48535140~6714404~171440~2057280~684728~1372552~~TriÖu Phó C­êng~2100272184~79282217"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001027050~83971~1007652~1041289~-33637~~NguyÔn Quèc Huy~~75754177~30962570~3331852~41459755~6312849~131285~1575420~1514292~61128~~Huúnh VÜnh Léc~~96864265~43142940~4505957~49215368~8072023~307202~3686424~2542364~1144060~~Tr­¬ng B¶o B×nh~~73953806~30823487~3561852~39568467~6162818~116282~1395384~962302~433082~~NguyÔn Thµnh T©n~~77744466~31668898~3291848~42783720~6478706~147871~1774452~1415250~359202~~Lª Thanh ThuËn B×nh~~67761300~27547688~2832513~37381099~5646776~64678~776136~624754~151382~~NguyÔn H÷u S©m~~70304775~28199039~3081842~39023894~5858732~85873~1030476~985033~45443~~T¹ Anh Sa~2100262179~103427006~37858384~5328000~60240622~8618917~361892~4342704~1713386~2629318~~TrÇn §¨ng Khoa~2100272191~106101"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001034050~1143133~41515698~5758989~75899~910788~1074810~-164022~~D­¬ng V¨n Tr¹ng ~~72929726~28849084~760000~43320642~6077477~107748~1292976~1299052~-6076~~Lª Tr­êng S¬n Vò ~~62374314~23523018~730000~38121296~5197860~19786~237432~560641~-323209~~Th©n V¨n Nhi~~64075619~24846509~950000~38279110~5339635~33964~407568~678313~-270745~~L÷ V¨n To¶n~~64208021~24574616~929728~38703677~5350668~35067~420804~661199~-240395~~NguyÔn Th¸i S¬n~~63604899~24599971~904999~38099929~5300408~30041~360492~645625~-285133~~NguyÔn V¨n D¹n~~66414280~27623839~820000~37970441~5534523~53452~641424~794197~-152773~~TrÇn Vò Toµn~~65627226~26375736~950000~38301490~5468936~46894~562728~828677~-265949~~Bµnh Kim B¶o Träng ~~68191443~2778145"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001033050663876~66388~796656~1011361~-214705~~NguyÔn Tr­êng HËn ~~61341513~24926722~1528681~34886110~5111793~11179~134148~394928~-260780~~NguyÔn ChÝ Trung ~~72847113~28932703~820000~43094410~6070593~107059~1284708~1315180~-30472~~D­¬ng Minh Hïng ~~76917108~28190671~999728~47726709~6409759~140976~1691712~1613361~78351~~TrÇn V¨n DÝnh ~~70154637~26355954~1100000~42698683~5846220~84622~1015464~1101254~-85790~~L©m Quèc C­êng ~~61760356~23995351~700000~37065005~5146696~14670~176040~562703~-386663~~TrÇn Minh Lu©n ~~73046432~28613153~920000~43513279~6087203~108720~1304640~1363263~-58623~~TrÞnh Thanh Tïng~~73625004~28834519~850000~43940485~6135417~113542~1362504~1391732~-29228~~D­¬ng Quan Huy~~69107871~26449040"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100103205070~1169618~~Lª Ph­íc S¬n~2100271977~95330862~37016008~3210000~55104854~7944239~294424~3533088~2918390~614698~~NguyÔn V¨n Lanh ~2100271945~90980634~38258336~1320000~51402298~7581720~258172~3098064~2522556~575508~~Lª V¨n HiÒn ~~81133500~33739740~860000~46533760~6761125~176113~2113356~1833537~279819~~NguyÔn Hoµng ~~75841890~28923160~800000~46118730~6320158~132016~1584192~1545702~38490~~Lª Minh D­¬ng ~~76185693~32676722~790000~42718971~6348808~134881~1618572~1424491~194081~~Vâ Thanh Phong ~~65280158~26045955~764278~38469925~5440013~44001~528012~781531~-253519~~Ph­¬ng TÊn H­ng ~~72693820~26464374~959278~45270168~6057818~105782~1269384~1330321~-60937~~C« V¨n Tù ~~67966509~27485350~1024047~39457112~5"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100103105027279031~844000~41602623~5810471~81047~972564~405089~567475~~TrÇn Th¸i Hßa~~67152957~26672678~800000~39680279~5596080~59608~715296~347606~367690~~TrÇn Minh Hïng~2100272106~87491635~34557763~890000~52043872~7290970~229097~2749164~1257694~1491470~~§ç V¨n Trung HiÕu~2100272120~103816149~40819696~1064000~61932453~8651346~365135~4381620~1772112~2609508~~Mai Khoa §¨ng~~73570629~28494649~1315000~43760980~6130886~113089~1357068~530167~826903~~§ç V¨n Th¾ng~~66110253~26063379~944000~39102874~5509188~50919~611028~330961~280067~~NguyÔn V¨n Phông~~66265403~24879519~1167200~40218684~5522117~52212~626544~348565~277979~~Phïng ThÕ Mü ~2100261753~120374932~46657129~4330000~69387803~10031244~503124~6037488~48678"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001038050M¹ch V¨n §¼ng~~68761518~24073526~1450000~43237992~5730127~73013~876156~807527~68629~~NguyÔn TÊn SÜ~~80291376~32584111~968000~46739265~6690948~169095~2029140~1359914~669226~~S¬n khone~~74561649~27658816~920000~45982833~6213471~121347~1456164~1153148~303016~~NguyÔn TriÖu ¢n Khoa~~78634163~24201832~1122013~53310318~6552847~155285~1863420~1178798~684622~~Ng« Thanh Dòng~~73215799~25867026~810000~46538773~6101317~110132~1321584~1030277~291307~~NguyÔn H¶i Phong~~70825072~27990756~790000~42044316~5902089~90209~1082508~934222~148286~~Th¹ch Thanh S¬n~~74985570~27389545~920000~46676025~6248798~124880~1498560~1371241~127319~~TrÞnh Minh Hoµng~~69903488~26536346~968000~42399142~5825291~82529~990348~861802~1"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100103705009~700000~36004855~5294155~29416~352992~424489~-71497~~D­¬ng Quèc B¶o~~67285745~24368517~1239816~41677412~5607145~60715~728580~745988~-17408~~NguyÔn Quèc Khang~~72057126~26903075~1220614~43933437~6004761~100476~1205712~971487~234225~~Tèng B¸ Tïng~2100272554~104767736~37512971~3450000~63804765~8730645~373065~4476780~3233968~1242812~~TrÇn ChÝ Minh~~70964207~24397973~1946377~44619857~5913684~91368~1096416~964537~131879~~Huúnh Duy C­¬ng~2100272561~68910977~26430746~856000~41624231~5742581~74258~891096~803120~87976~~Vâ §øc Trung~2100272579~76788591~27341549~1580000~47867042~6399049~139905~1678860~1232815~446045~~NguyÔn V¨n KÕt~~66055203~23634638~2565979~39854586~5504600~50460~605520~664810~-59290~~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001036050i Thanh Xu©n ~~62490392~24439430~800000~37250962~5207533~20753~249036~584913~-335877~~TrÇn V¨n Vò~~62168306~25281849~760000~36126457~5180692~18069~216828~541088~-324260~~NguyÔn Kh¸nh HuÊn~~71452789~28242391~800000~42410398~5954399~95440~1145280~1169360~-24080~~NguyÔn V¨n Khiªm~2100262186~112437118~40457449~5070614~66909055~9369760~436976~5243712~3600706~1643006~~Bµnh §øc B×nh~2100272498~96365097~38574404~990000~56800693~8030425~303043~3636516~2483875~1152641~~L©m Quèc C­êng ~2100272515~90780589~35907835~920000~53952754~7565049~256505~3078060~2138509~939551~~L­u TrÝ Dâng~2100272522~83737249~30065005~970000~52702244~6978104~197810~2373720~1884848~488872~~L©m Thanh Tïng~2100272547~63529864~268250"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010350502~920000~39489991~5682620~68262~819144~925430~-106286~~Ph¹m ChÝ T©m~~65952320~26980003~820000~38152317~5496027~49603~595236~785707~-190471~~DiÖp TiÕn Ng÷~~62128672~26450653~700000~34978019~5177389~17739~212868~518357~-305489~~C« V¨n B×nh ~2100349366~82488856~34086012~700000~47702844~6874071~187407~2248884~1974601~274283~~Huúnh Hång H¶i~2100272018~89140239~35891067~1050000~52199172~7428353~242835~2914020~2483222~430798~~L©m V¨n Nhanh ~2100349373~87036180~34104789~900000~52031391~7253015~225302~2703624~2348340~355284~~NguyÔn Quèc Kh¶i~~63366252~24816857~920000~37629395~5280521~28052~336624~672544~-335920~~TrÇn Trung §«ng ~~62203450~24915068~890000~36398382~5183621~18362~220344~541333~-320989~~Bï"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001042050 Thanh~~76960255~29013202~2000000~45947053~6413355~141336~1696026~1824357~-128331~~T¹ V¨n Hµo~~70925199~26362979~853910~43708310~5910433~91043~1092520~1389792~-297272~~Th¹ch V¨n NhÆn~~67420802~26869889~700000~39850913~5618400~61840~742080~1072758~-330678~~Lª V¨n Dòng~~78652593~32393118~820000~45439475~6554383~155438~1865260~2035073~-169813~~Th¹ch Ngäc Th¸i~~83513050~33268939~700000~49544111~6959421~195942~2351305~2420042~-68737~~Kim Ngäc Toµn~~71139625~27259616~884201~42995808~5928302~92830~1113962~1388785~-274823~~S¬n Th¸i Giang~~69769003~27235543~1123444~41410016~5814084~81408~976901~1186254~-209353~~NguyÔn V¨n Phong~~72867124~29930497~1050000~41886627~6072260~107226~1286712~1449998~-163286~"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001041050h¹m Hoµng Nguyªn ~~118121820~45013510~3400000~69708310~9843485~484349~5812182~5322341~489841~~DiÖp Rªnh~~104198818~40514989~1091301~62592528~8683235~368324~4419882~4088253~331629~~Kim TuyÒn~~61542927~22968047~911910~37662970~5128577~12858~154292~581448~-427156~~Vâ Minh Lu©n~~61690904~24396928~800000~36493976~5140909~14091~169091~607900~-438809~~Cï Thanh T©m~~81431385~31416351~1478000~48537034~6785949~178595~2143139~2332086~-188947~~NguyÔn Ngäc Ven~~65240046~26140122~908482~38191442~5436671~43667~524005~886899~-362894~~Kim Ngäc Thi~~68037676~27163730~923444~39950502~5669806~66981~803767~1082610~-278843~~Ng« Quèc Kû~~86898095~37902316~950000~48045779~7241508~224151~2689810~2660423~29387~~Lª Quèc"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   0020080010010400508~~Tr­¬ng V¨n Qu©n~~65737078~25205216~790000~39741862~5478090~47809~573708~634430~-60722~~Lª Minh TrÝ ~~69703453~26205835~800000~42697618~5808621~80862~970344~875826~94518~~T¹ H÷u HiÕu ~~61946566~23765048~800000~37381518~5162214~16221~194652~479714~-285062~~NguyÔn TÊn Thi ~~61021870~22814108~900000~37307762~5085156~8516~102192~422604~-320412~~TrÞnh Minh TuÊn ~~63724368~24087260~1020000~38617108~5310364~31036~372432~584334~-211902~~Lª Hång Nhùt ~~61939725~23210546~700000~38029179~5161644~16164~193968~471874~-277906~~NguyÔn Thanh L©m ~~63552291~24050166~920000~38582125~5296024~29602~355224~572378~-217154~~Ph¹m ThÕ Toµn~~116726275~45712635~4791444~66222196~9727190~472719~5672628~5251855~420773~~P"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100103905028546~~NguyÔn Hång Th¾ng ~~66735623~23496209~1164614~42074800~5561302~56130~673560~705791~-32231~~NguyÔn V¨n ót Em ~~63666500~23324861~800000~39541639~5305542~30554~366648~513572~-146924~~NguyÔn C«ng Khai~~61888855~21452820~800000~39636035~5157405~15741~188892~428587~-239695~~La Thµnh Lîi~~65649544~25266470~830000~39553074~5470795~47080~564960~642227~-77267~~NguyÔn V¨n ót~~66073453~25386486~820000~39866967~5506121~50612~607344~656041~-48697~~Lª Thµnh Ch­¬ng~~63743238~22571147~1154614~40017477~5311937~31194~374328~535561~-161233~~Lý Quèc An~~67248731~22223851~993210~44031670~5604061~60406~724872~609520~115352~~Tr­¬ng H÷u Tho¹i~~65979172~25042665~890000~40046507~5498264~49826~597912~655220~-5730"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001046050913~92825412~36021670~2199642~54604100~7735451~273545~3282540~2975198~307342~~NguyÔn V¨n HiÒn~2100271938~89854470~32354653~1461267~56038550~7487872~248787~2985444~2666174~319270~~TrÇn Minh Trung~2100271952~86376100~34337559~1774600~50263941~7198008~219801~2637612~2170327~467285~~Lª Hoµng Ph­¬ng~2100262122~93044260~34179053~1314243~57550964~7753688~275369~3304428~2900906~403522~~L©m Thanh Liªm~2100271984~74236942~27076759~1501983~45658200~6186412~118641~1423692~1520216~-96524~~Ng« L­¬ng Hång Léc~2100272000~83468072~31636190~1631827~50200055~6955673~195567~2346804~2361931~-15127~~Th¹ch Ngäc Ny~2100271409~83285403~29902942~1386227~51996234~6940450~194045~2328540~2241423~87117~~TrÇn S« B« Tra~2100"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100104505002063~2424756~2345640~79116~~Ng« Minh Hïng ~~66950284~27746973~790000~38413311~5579190~57919~695028~1044392~-349364~~D­¬ng V¨n Tµi ~~65261142~25590674~800000~38870468~5438429~43843~526115~887094~-360979~~§inh Thanh TrÝ ~~66980651~27735045~700000~38545606~5581721~58172~698065~1075930~-377865~~TrÇn TiÕn Dòng ~~63209991~22440760~990000~39779231~5267499~26750~320999~783017~-462018~~TrÇn Hoµng §¹t~2100262108~127567677~45891905~7115667~74560105~10630640~563064~6756768~5204623~1552145~~NguyÔn Phóc Léc~~82689606~34462542~1291203~46935861~6890800~189080~2268960~1867334~401626~~NguyÔn V¨n Phong~2100271906~104191474~42700879~1672251~59818344~8682623~368262~4419144~3320550~1098594~~TrÇm Thanh Long~2100271"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001044050~930396~-369692~~NguyÔn Thµnh Trung~~72054506~31219500~700000~40135006~6004542~100454~1205450~1394425~-188975~~Huúnh C«ng B»ng~~63273272~23445412~1005586~38822274~5272773~27277~327328~700260~-372932~~Th¹ch Ngäc Minh~~70593298~30028192~1070000~39495106~5882775~88278~1059330~1234491~-175161~~NguyÔn §¨ng Phong~~68685293~27783393~700000~40201900~5723774~72377~868529~1082385~-213856~~Høa V¨n Vinh~~70342460~29302987~700000~40339473~5861872~86187~1034246~1242226~-207980~~S¬n Th¸i Vinh~~71211044~30386958~820000~40004086~5934254~93425~1121105~1316623~-195518~~TrÇn ChÝ Hïng~~71304647~29867741~940000~40496906~5942054~94205~1130465~1322501~-192036~~Lª Tr­êng An~~84247564~34652972~920000~48674592~7020630~2"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001043050~L©m Hoµng Nam~~79988440~32479082~820000~46689358~6665703~166570~1998844~2101172~-102328~~S¬n Thµnh C«ng~~73759419~31552619~820000~41386800~6146618~114662~1375942~1553578~-177636~~L©m Quang H­ng~~78906776~32899458~790000~45217318~6575565~157557~1890678~2052905~-162227~~NguyÔn T©n Thanh~~70345407~26497956~989286~42858165~5862117~86212~1034540~1342174~-307634~~NguyÔn ChÝ T©m~~64681582~25350100~1079755~38251727~5390132~39013~468158~773484~-305326~~Ph­¬ng Kim Ngäc~~77719127~30976626~850000~45892501~6476594~147659~1771913~1927892~-155979~~Th¹ch Ngäc Quan~~77624923~31987809~950000~44687114~6468744~146874~1762493~2052553~-290060~~Bïi Thanh §iÒn ~~65607038~24553885~730000~40323153~5467253~46725~560704"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100105005035529~1398693~38678580~5384400~38440~461280~743706~-282426~~NguyÔn Thanh Ph­¬ng ~~62112594~22661403~1576243~37874948~5176049~17605~211260~672920~-461660~~Vâ Hoµng Vò ~~63824245~24520627~1504259~37799359~5318687~31869~382428~776182~-393754~~Tiªu Tr­êng Th¹nh ~~61740347~22788994~1311165~37640188~5145029~14503~174036~637490~-463454~~Ph¹m V¨n Ph­íc~~65433041~24519923~1402328~39510790~5452753~45275~543300~823202~-279902~~NguyÔn Thanh Long~~73643131~26586417~1504443~45552271~6136928~113693~1364316~1592778~-228462~~NguyÔn V¨n TriÕt ~2100271920~79374154~31408141~1286227~46679786~6614513~161451~1937412~1977893~-40481~</S><S>~~0~0~0~0~0~0~0~0~0~</S></S04-1>"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001049050~42660144~5929566~92957~1115484~1345516~-230032~~NguyÔn Thanh T©m (M)~~74841614~28432863~1396179~45012572~6236801~123680~1484160~1615814~-131654~~Danh Pholla~~73305682~26801974~1506243~44997465~6108807~110881~1330572~1542323~-211751~~NguyÔn T©y Nam ~~64741177~21628108~2567667~40545402~5395098~39510~474120~774236~-300116~~Lª NguyÔn Th­¬ng ~~61144872~20981985~1446179~38716708~5095406~9541~114492~568111~-453619~~Ph¹m ThÕ Nh©n~~63010788~23433020~2336414~37241354~5250899~25090~301080~437157~-136077~~TrÇn V¨n LÑ~~62334611~23187772~2226842~36919997~5194551~19455~233460~394632~-161172~~D­¬ng Hoµng S¬n~~60333754~22685391~1428668~36219695~5027813~2781~33372~261587~-228215~~NguyÔn Minh ViÔn~~64612802~245"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   00200800100104805051~5755425~75542~906504~989063~-82559~~Hå Minh HuÊn~~86320926~29436335~4317651~52566940~7193410~219341~2632092~2270632~361460~~Mai Léc B¸ch~~68910259~24798920~2263892~41847447~5742522~74252~891024~1235265~-344241~~TrÇn Thanh Th¸i ~~61223231~24990045~1307176~34926010~5101936~10194~122328~351934~-229606~~L÷ Minh Tïng~~68753120~25812243~1470415~41470462~5729427~72943~875316~1283742~-408426~~NguyÔn Trung HËu~~85286688~32611038~2599981~50075669~7107224~210722~2528664~2299383~229281~~TrÇn V¨n RiÕp~~64919104~23612434~2830707~38475963~5409925~40993~491916~688356~-196440~~L©m V¨n ót Lín~~74255890~27450254~1226179~45579457~6187991~118799~1425588~1594317~-168729~~NguyÔn V¨n SÊm~~71154788~27126481~1368163"
'    Barcode_Scaned str1
'    str1 = "aa200172300103987   002008001001047050272025~83703947~31509512~1496179~50698256~6975329~197533~2370396~2342381~28015~~S¬n Chia Phi Rïm~~68600447~25209113~2694238~40697096~5716704~71670~860040~1125901~-265861~~TrÇn Träng Tr­êng~2100272057~76002787~27624887~1728163~46649737~6333566~133357~1600284~1668364~-68080~~Hµ S¬n L©m~2100272071~75189222~29178684~1412195~44598343~6265768~126577~1518924~1557668~-38744~~NguyÔn Thµnh TËp~2100272089~71887572~26046646~1441923~44399003~5990631~99063~1188756~1450347~-261591~~L©m TÊn §¹t~~88746276~33321463~1939717~53485096~7395523~239552~2874624~2586873~287751~~NguyÔn Quèc C­êng~~72195774~27110429~2279155~42806190~6016314~101631~1219572~1379954~-160382~~Huúnh Ngäc Vinh~~69065100~27373092~1636357~400556"
'    Barcode_Scaned str1
'
    


'KHBS 01/GTGT
'str1 = "bs131012300103987   07200800100100100201/0114/06/2006<S01><S>0~0~0~0~0~0~0~0~0~0~0~0~0~500000~0~1421824~0~0~1421824~0~0~1245924~0~175900~0~0~0~0~0~1421824~921824~0~0~0</S></S01>"
'Barcode_Scaned str1
'str2 = "bs131012300103987   072008001001002002<SKHBS><S>ThuÕ GTGT HHDV b¸n ra chÞu thuÕ 5%~31~ThuÕ GTGT HHDV b¸n ra chÞu thuÕ 5%                                                                  L_23                31~1754076~3000000~1245924~ThuÕ GTGT HHDV b¸n ra chÞu thuÕ 10%~33~ThuÕ GTGT HHDV b¸n ra chÞu thuÕ 10%                                                                 L_24                33~5824100~6000000~175900</S><S>Tæng sè thuÕ GTGT ®­îc khÊu trõ kú nµy~23~Tæng sè thuÕ GTGT ®­îc khÊu trõ kú nµy                                                              L_17                23~0~500000~-500000</S><S>8~3687~tai lieu dinh kem~Hµ ThÕ Ph­¬ng~28/08/2008~921824</S></SKHBS>"
'Barcode_Scaned str2






'KHBS 03/GTGT
'str1 = "tt130042300100601   02200800100100100201/0114/06/2006<S01><S>1000000~200000~3000000~400000~5000000~600000~10000000~20000000~2000000~19000000~1000000~1800000~2800000</S></S01>"
'Barcode_Scaned str1
'str2 = "tt130042300100601   022008001001002002<SKHBS><S>ThuÕ GTGT ph¶i nép (thuÕ suÊt 5%)~20~ThuÕ GTGT ph¶i nép (thuÕ suÊt 5%)                                                                   M_14                20~100000~1000000~900000</S><S>ThuÕ GTGT ph¶i nép (thuÕ suÊt 10%)~21~ThuÕ GTGT ph¶i nép (thuÕ suÊt 10%)                                                                  P_14                21~1900000~1800000~-100000</S><S>61~24400~~20/05/2008~800000</S></SKHBS>"
'Barcode_Scaned str2

'KHBS 01/TAIN
'str1 = "bs131062300100601   02200800300400100301/0114/06/2006<S01><S></S><S>0101~~1000~2000~2.0000~0~0~010201~~2000~4000~2.0000~0~0~010101~~0~0~0~0~0~010203~~0~0~0~0~0</S><S>~08/03/2008</S></S01>"
'Barcode_Scaned str1
'str2 = "bs131062300100601   022008003004002003<SKHBS><S>Kho¸ng s¶n kim lo¹i ®en~10~Kho¸ng s¶n kim lo¹i ®en                                                                                                                                                                                 0101      Kg        2         0~10000000~10040000~40000~Vµng sa kho¸ng~10~Vµ"
'Barcode_Scaned str2
'str3 = "bs131062300100601   022008003004003003ng sa kho¸ng                                                                                                                                                                                          010201    Kg        2         0~4000000~4160000~160000</S><S>~~~0~0~0</S><S>67~6700~~26/05/2008~200000</S></SKHBS>"
'Barcode_Scaned str3





'aa130014200324111   03200800800800100501/0114/06/2006<S01><S>~10000000~11133333333~557777778~11133333333~557777778~0~0~0~0~0~0~557777778~200000~16834444549~1217222223~0~16834444549~1217222223~2156666775~5011111109~250555556~9666666665~966666667~10000000~10000000~0~213212121~16844444549~1014010102~1003810102~0~0~0</S></S01>
'aa130014200324111   032008008008002005<S01_3><S>01/2008~19/04/2008~10000000~0</S></S01_3>#
'aa130014200324111   032008008008003005<S01_4A><S>56311111110~200000000~55555555555~555555555~0~0~0.00~555555555~0</S></S01_4A>#
'aa130014200324111   032008008008004005<S01_4B><S>2008~22246666666~22222222~2222222~22222222222~211111111~212121~0.10~22222222222~22222222~211112~22011110</S></S01_4B>#
'aa130014200324111   032008008008005005<S01_5><S>1345666~01/01/2008~DTYUII~RDGHHHNN~213212121</S></S01_5>#


'02  GTGT
'   str1 = "aa130022300103987   02200800300300100201/0114/06/2006<S01><S>ten du an dau tu~1000000~10000000~1000000~6000000~500000~4000000~500000~600000~700000~800000~900000~800000~1800000~100000~110000~1590000</S></S01>"
'   Barcode_Scaned str1
'   str2 = "aa130022300103987   022008003003002002<S01_2><S>Ky hieu hoa don~1111111~01/01/2008~ha the phuong~~mat hang~10000000~10.00~1000000~ghi chu</S><S>10000000~1000000</S></S01_2>"
'   Barcode_Scaned str2

'03 GTGT

'    str1 = "aa130042600105632   02200800100100100101/0114/06/2006<S01><S>1000000~200000~3000000~400000~5000000~600000~10000000~20000000~2000000~19000000~100000~1900000~2000000</S></S01>"
'    Barcode_Scaned str1

' 04/GTGT
'    str1 = "aa131072300103987   00200700100100100101/0114/06/2006<S01><S>100000000~100000000~10000000~10000000~1100000000~100000000~1090000000~90000000~54500000~9000000~100000~100000~54400000~8900000~63300000~0</S></S01>"
'    Barcode_Scaned str1


'01B TNDN
        
'     str1 = "aa200122400181361   02200900500500100201/0114/06/2006<S01><S>100042245~100000000~42245~10.000~10.000~~5.000~1000211~1000000~211~0~1000211~~08/03/2008</S></S01>"
'     Barcode_Scaned str1
'     str2 = "aa200122400181361   022009005005002002<S01-1><S>ha the phuong~02/01/2007~ha the phuong~100000~10000~100~1000~ha the phuong1~02/01/2007~ha the phuong2~20000~20000~500~2000</S></S01-1>"
'     Barcode_Scaned str2

'0AB TNDN
'str1 = "aa250110101724672   01200902602700100201/0114/06/2006<S01><S>1000000~100000~1120000~1211~1244426~24.000~5~185842~NguyÔn ThÞ Nga~24/03/08..</S></S01>"
'Barcode_Scaned str1
'str1 = "aa250110101724672   012009026027002002<S01-1><S>1~01/12/2007~söegtÎ~2q3r43~2000~20000000~5000000</S></S01-1>"
'Barcode_Scaned str1



'03 TNDN
        
'     str1 = "aa250030101724672   00200900100100100101/0114/06/200601/01/200931/12/2009<S03><S>100000000~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~100000000~100"
'     Barcode_Scaned str1
'     str2 = "aa250030101724672   002009001001001001000000~0~0~0~0~100000000~100000000~0~25000000~25000000~0~0~0~0~0~0~0~25000000~~~~~~~~Ha~06/05/2010</S></S03>"
'     Barcode_Scaned str2

'05/TNDN
'    str1 = "aa130142600105632   02200800100100100101/0114/06/2006<S01><S>ha the phuong~2300103987~01/01/2008~01/02/2008~100000~10.00~10000~test1~ha the phuong 1~2300103987~01/01/2008~01/01/2008~200000~12.00~24000~test2</S><S>htphuong~22/03/2008</S></S01>"
'    Barcode_Scaned str1


'01/TAIN
'    str1 = "aa131062300103987   01200900200200100101/0114/06/2006<S01><S></S><S>0101~Kg~10000.00~0,00~0.0000~1000~0~010201~Kg~2000.00~0,00~0.0000~2000~0~010101~Kg~3000.00~0,00~0.0000~3000~0</S><S>htphuong~08/03/2008</S></S01>"
'    Barcode_Scaned str1
'02/TAIN
'    str1 = "aa130092600105632   02200800100100100101/0114/06/2006<S02><S></S><S>010203~Kg~100.00~0,00~0.00~100~0~010202~Kg~200.00~0,00~0.00~100~0</S><S>~15/03/2008</S></S02>"
'    Barcode_Scaned str1

'03/TAIN
'    str1 = "aa1310862300103987   00200900200200100101/0114/06/2006<S03><S></S><S>0101~Kg~10000.00~0,00~0.0000~1000~0~010201~Kg~2000.00~0,00~0.0000~2000~0~010101~Kg~3000.00~0,00~0.0000~3000~0</S><S>htphuong~08/03/2008</S></S03>"
'    Barcode_Scaned str1
'
'     str1 = "aa130082300103987   00200800100100100101/0114/06/2006<S03><S></S><S>0101a~Kg~1000.00~10.00~2.0000~0~0~0101b~Kg~2000.00~10.00~5.0000~0~0</S><S>~12/06/2008</S></S03>"
'     Barcode_Scaned str1

'01/TTDB
    str1 = "aa250050101724672   04201000200200100101/0101/01/1900<S01><S>~11111~7663.00~0~0~3448</S><S>10200~LÝt~2.00~11111~7663.00~45.0~0~0~3448</S><S>150000~115385.00~0~0~34616</S><S>20102~phong~6.00~150000~115385.00~30.0~0~0~34616</S><S>11714443</S><S>10200~LÝt~6.00~565555</S><S>10102~Bao~2.00~8560000</S><S>10103~Chai~2.00~2588888</S><S>11875554~123048.00~0~0~38064</S></S01>"
    Barcode_Scaned str1
'
    
'     str1 = "aa130052600105632   02200800400400100301/0101/01/1900<S01><S>~10000~6061.00~0~0~3940</S><S>10101~Bao~100.00~10000~6061.00~65~0~0~3940</S><S>10000~9091.00~0~0~909</S><S>20400~~1000.00~10000~9091.00~10~0~0~909</S><S>0</S><S>10600~C¸i~0.00~0</S><S>~~0.00~0</S><S>~~0.00~0</S><S>20000~15152.00~0~0~4849</S></S01>"
'     Barcode_Scaned str1
'     str2 = "aa130052600105632   022008004004002003<S01-1><S>568568~0101~01/01/2007~6223623~20300~2362562~10000~1000~10000000~46346~34634~01/01/2008~325235~10102~234626~1000~235325~235325000</S><S>Kinh doanh gi¶i trÝ cã ®Æt c­îc~10000~10000000~Thuèc l¸ ®iÕu~1000~235325000</S></S01-1>"
'     Barcode_Scaned str2
'     str3 = "aa130052600105632   022008004004003003<S01-2><S>~~~~0.00~0~0</S><S>~0~~0.00~0</S></S01-2>"
'     Barcode_Scaned str3

    

'01A/TNCN
'    str1 = "aa130152600105632   02200800100100100101/0101/01/1900<S01><S>0100000~VCB~1000~500~100~50~1000000~5000000~6000000~7000000~htphuong~23/03/2008</S></S01>"
'    Barcode_Scaned str1
'01_1/TNCN
'    str1 = "aa130222600105632   04200700100100100101/0101/01/1900<S01><S>000001~VP~100000~20000~1000~5000~100000~10000~10000~20000~htphuong~23/03/2008</S></S01>"
'    Barcode_Scaned str1

'02/TNCN
'    str1 = "aa130162600105632   02200800300300100101/0101/01/1900<S02><S>4~1150000000~300000000~850000000~115000000~30000000~85000000~575000~0~114425000~29850000~84575000</S><S>1~123~123~~123~100000000~10000000~123~01/01/2008~2~321~321~~321~200000000~20000000~321~01/01/2008</S><S>1~456~456~~456~500000000~50000000~654~01/01/2008~2~789~789~~789~350000000~35000000~987~01/01/2008</S><S>~12/03/2008</S></S02>"
'    Barcode_Scaned str1

'02_1/TNCN
'    str1 = "aa130232600105632   04200700100100100101/0101/01/1900<S02><S>4~100000000~30000000~70000000~10000000~3000000~7000000~50000~0~9950000~2985000~6965000</S><S>0~ha the phuong~123~~test~10000000~1000000~011111~01/01/2007~2~ha the phuong1~345~~test1~20000000~2000000~022222~01/02/2008</S><S>1~ha the phuong3~567~~test3~30000000~3000000~033333~01/03/2008~2~ha the phuong4~789~~test4~40000000~4000000~044444~01/03/2008</S><S>htphuong~23/03/2008</S></S02>"
'    Barcode_Scaned str1

'04/TNCN
'    str1 = "aa130172600105632   00200700400500100301/0101/01/190001/01/200731/12/2007<S04><S>6~2~4~2~1~1~14400000~12000000~2400000~8865774~968987~7896787~44329~4845~39484~4~1~3~8419885~7536756~883129~818767~57457~761310~4094~287~3807~1686706~455145~1231561~65182478~65156216~26262~325912~325781~131~24506591~19991901~4514"
'    Barcode_Scaned str1
'    str1 = "aa130172600105632   002007004005002003690~74867019~66182660~8684359~374335~330913~43422~74492684~65851747~8640937</S><S>542646~~7536756~57457~74567~01/01/2001</S><S>6453654~~764~745756~745756~01/01/2001~785768765~~786576~7685~76878~01/01/2001~85875~~95789~7869~7813543~01/01/2001</S><S>~21/03/2008</S></S04>"
'    Barcode_Scaned str1
'    str1 = "aa130172600105632   002007004005003003<S04-1><S>6357~~12000000~1000000~10000000~1000000~1000000~0~968987~87976~881011~</S><S>4563456~~2400000~200000~200000~2000000~200000~9~7896787~99789~7796998~</S></S04-1>"
'    Barcode_Scaned str1

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
    Dim I As Long
    Dim varBuff As Variant
    Dim lByte() As Byte
        
On Error GoTo ErrHandle
    Select Case MSComm1.CommEvent
        Case comEvReceive                                       ' Received RThreshold # of chars.
            varBuff = MSComm1.Input
            lByte = varBuff
            For I = 0 To UBound(lByte)
                If Chr$(lByte(I)) <> "#" Then
                    strTemp = strTemp & Chr$(lByte(I))
                Else
                    Barcode_Scaned strTemp
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
    Dim strPrefix As String, strBarcodeCount As String, strData As String
    Dim idToKhai As String
    
    
    If Left$(strBarcode, 2) = "bs" Then
        LoaiTk = "TKBS"
    Else
        LoaiTk = ""
    End If
    strBarcode = TrimString(strBarcode)
    strBarcode = TAX_Utilities_Svr_New.Convert(strBarcode, TCVN, UNICODE)
    
    If Left$(strBarcode, 1) <> "0" Then
        'Version 1.2.0 and later
        If Val(Left$(strBarcode, 3)) > Val(Replace$(APP_VERSION, ".", "")) Then
        'Version tai doanh nghiep lon hon tai co quan thue APP_VERSION
            DisplayMessage "0074", msOKOnly, miCriticalError
            Exit Sub
        ElseIf Val(Left$(strBarcode, 3)) < 200 Then ' Truong hop to khai thue TNCN duoc in bang phien ban 1.3.1 se khong con hieu luc theo luat thue TNCN moi nam 2009
            If Val(Mid$(strBarcode, 4, 2)) = 15 Or Val(Mid$(strBarcode, 4, 2)) = 16 Or Val(Mid$(strBarcode, 4, 2)) = 22 Or Val(Mid$(strBarcode, 4, 2)) = 23 Then
                DisplayMessage "0105", msOKOnly, miCriticalError
                Exit Sub
            End If
        End If
        
        strPrefix = Left$(strBarcode, 36)
        strBarcodeCount = Right$(strPrefix, 6)
        strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 6)
        
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
            If (Trim(idToKhai) = "53" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "37" And Val(Mid(strPrefix, 21, 4)) > 2009) _
                Or (Trim(idToKhai) = "54" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "38" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0115", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' To khai 01/TAIN, 02/TAIN, 03/TAIN bat dau thu thang 2 se ko nhan ND 50 2010 CP” doi voi to 01/TAIN và 02/TAIN va to khai co nien do 2010
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
        
            
            ReDim Preserve arrStrElements(intBarcodeCount)
            arrStrElements(intBarcodeNo) = strBarcode
            
            ' hlnam Edit
            ' Lay them trong truong hop ko quet het ma vach ma muon hien thi luon
            ReDim Preserve arrBCBuffer(intBarcodeCount)
            arrBCBuffer(intBarcodeNo) = strPrefix & strBarcodeCount & strBarcode
            
            
            If IsCompleteData(strData) Then
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
    Dim arrStrValue() As String
    
    lElementsNo = GetElementsNo(xmlSectionTemplate.childNodes(0))
    'Get array of data units
    arrStrValue = Split(xmlSectionData.Text, "~")
    If GetAttribute(xmlSectionTemplate, "Dynamic") = "0" Then
        'Static data
        If UBound(arrStrValue) + 1 > lElementsNo Then
            blnValidData = False
            'DisplayMessage "0070", msOKOnly, miCriticalError
            Exit Sub
        End If
    Else
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

Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, I As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For I = xmlNodeListCell.length - 1 To 0 Step -1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(I), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(I), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(I), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(I), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
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

On Error GoTo ErrHandle
    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
        TAX_Utilities_Svr_New.Month = Left$(strValue, 2)
        TAX_Utilities_Svr_New.ThreeMonths = ""
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = 1 Then
        TAX_Utilities_Svr_New.ThreeMonths = Left$(strValue, 2)
        TAX_Utilities_Svr_New.Month = ""
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
    
On Error GoTo ErrHandle
    
    TAX_Utilities_Svr_New.Month = ""
    TAX_Utilities_Svr_New.ThreeMonths = ""
    TAX_Utilities_Svr_New.Year = ""
    TAX_Utilities_Svr_New.FinanceStartDate = ""
    
'    If Left$(strData, 1) = "0" Then
'        strTaxReportVersion = "1.1.0"
'        lblVersion.caption = "1.1.0"
''**********************************
    
    If Left$(strData, 3) = "120" Then
        lblVersion.caption = "1.2.0"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    ElseIf Left$(strData, 3) = "130" Then
    'Version 1.3.0
        'Get version of application
        lblVersion.caption = "1.3.0"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    ElseIf Left$(strData, 3) = "131" Then
    'Version 1.3.1
        'Get version of application
        lblVersion.caption = "1.3.1"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    ElseIf Left$(strData, 3) = "200" Then
    'Version 2.0.0
        'Get version of application
        lblVersion.caption = "2.0.0"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    ElseIf Left$(strData, 3) = "210" Then
    'Version 2.1.0
        'Get version of application
        lblVersion.caption = "2.1.0"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    ElseIf Left$(strData, 3) = "250" Then
    'Version 2.1.0
        'Get version of application
        lblVersion.caption = "2.5.0"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    ElseIf Left$(strData, 3) = "252" Then
        'Version 2.1.0
        'Get version of application
        lblVersion.caption = "2.5.2"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    Else
        'Version 2.5.3
        'Get version of application
        lblVersion.caption = "2.5.3"
        strTaxReportVersion = Left$(strData, 3)
        strData = Mid$(strData, 4)
    End If
    
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
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then
        dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Svr_New.ThreeMonths), CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)
        dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
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
    
    Dim lSheet As Long, I As Long, j As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .Sheet = lSheet
            If .SheetVisible = True Then
                For I = 1 To .MaxRows
                    .Row = I
                    If .RowHeight(I) > 10 And .RowHeight(I) < 15 Then .RowHeight(I) = 14
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

    Dim rsHeaderData As ADODB.Recordset
    Dim arrStrHeaderData() As String
    Dim rs As ADODB.Recordset, strSQL As String
    Dim blnConnected As Boolean
    Dim strPhongXuLy As String
    
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
    
    clsDAO.Disconnect
    
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
    
    If CLng(Replace$(strTaxReportVersion, ".", "")) < CLng(Replace$(APP_VERSION, ".", "")) Then
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
        Dim I As Integer
        I = Len(GetAttribute(xmlNode, "ID"))
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
                arrStrData(intIndex) = Mid$(strBarcodeData, intLoc1, intLoc2 + I + 3)
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

Function GenerateSQL_Details(xmlDomData As MSXML.DOMDocument, strSQL_DTL As String, vHdrID As Variant, lPos As Long) As String
    Dim xmlListSection As MSXML.IXMLDOMNodeList
    Dim xmlNodeSection As MSXML.IXMLDOMNode
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim lPosition As Long, strCondition As String
    Dim I As Long, j As Long, strLoaiDL As String
    
On Error GoTo ErrHandle
    iRowID = 0
    Set xmlListSection = xmlDomData.getElementsByTagName("Section")
    For Each xmlNodeSection In xmlListSection
        If Trim(xmlNodeSection.Attributes.getNamedItem("Dynamic").nodeValue) = "1" Then
            For I = 0 To xmlNodeSection.childNodes.length - 1
                iRowID = iRowID + 1
                For j = 0 To xmlNodeSection.childNodes(I).childNodes.length - 1
                    Set xmlAttribute = xmlDomData.createAttribute("RowID")
                    xmlAttribute.Value = iRowID
                    Set xmlNode = xmlNodeSection.childNodes(I).childNodes(j).Attributes.setNamedItem(xmlAttribute)
                    Set xmlAttribute = Nothing
                Next
            Next
        End If
    Next
    
    strLoaiDL = Trim(TAX_Utilities_Svr_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue) & Trim(TAX_Utilities_Svr_New.NodeValidity.childNodes(lPos).Attributes.getNamedItem("ID").nodeValue)
    Set xmlList = xmlDomData.getElementsByTagName("Cell")
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
    Set xmlDomData = Nothing
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
    
    Dim I As Long
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
        
        For I = 12 To .MaxRows
            .Sheet = mHeaderSheet
            .Col = .ColLetterToNumber("B")
            .Row = I
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
        For I = 12 To .MaxRows
            .Sheet = mHeaderSheet
            .Col = 2
            .Row = I
            vFormulaFunc = .Formula
            If Trim(.Text) <> "" Then
                .GetText .ColLetterToNumber("B"), I, vFunction
                .GetText .ColLetterToNumber("E"), I, vMsg
                .GetText .ColLetterToNumber("S"), I, vWarning
                .GetText .ColLetterToNumber("T"), I, vOrder
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
            For I = 2 To cOrder.Count
                X = Val(Left(cOrder(I), InStr(cOrder(I), "[]")))
                If min >= X Then
                    min = X
                    strCell = Right(cOrder(I), Len(cOrder(I)) - InStr(cOrder(I), "[]") - 1)
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
    Dim I As Long
    
On Error GoTo ErrHandle
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For I = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, I)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, I)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For I = 1 To .SheetCount
            .Sheet = I
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
                lSheet = I
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
    Dim lCol As Long, lRow As Long, I As Long
    
On Error GoTo ErrHandle
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For I = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, I)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, I)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For I = 1 To .SheetCount
            .Sheet = I
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
    Dim I As Long, lCol As Long, lRow As Long
    Dim xmlNodeListIni As MSXML.IXMLDOMNodeList
    Dim xmlNodeIni As MSXML.IXMLDOMNode
    
    For I = 0 To fpSpread1.SheetCount - 2
        ReDim Preserve xmlDocumentInit(I)
        Set xmlDocumentInit(I) = New MSXML.DOMDocument
        xmlDocumentInit(I).Load GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(I), "Ini"))
        Set xmlNodeListIni = xmlDocumentInit(I).getElementsByTagName("Cell")
        For Each xmlNodeIni In xmlNodeListIni
            fpSpread1.Sheet = I + 1
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
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
   Dim rs As ADODB.Recordset
   Dim MATKHAI As Variant
   Dim KYLBO As Variant
   Dim NGNOP As Variant
   Dim NGNHAP As Variant
   Dim KYKKHAI As Variant
   Dim maDTNT As Variant
   
   Dim DAGHI As Variant
   Dim LOAITIEN As Variant
   Dim MAVACH As Variant
   Dim SHTEP As Variant
   Dim HANNOP2 As Variant
   Dim THUETKY As Variant
   Dim THUETKY2 As Variant
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim CTHUC As Variant
   Dim BSUNG As Variant
   Dim LANBS As Variant
   Dim TRICH_YEU As String
   
   Dim CHKGIAHAN As Variant
   Dim bln  As Boolean
   'dhdang them bien
   'Dim DHS_MA As Variant
   Dim PHONG_XL As Variant
   Dim PHONG_XL_X As Variant
   Dim PHONG_XL_Y As Variant
   Dim SO_HOSO As Variant
   Dim NGAY_XL As Date
   Dim NGAY_HEN As Date
   Dim NGAY_NHAN As Date
   Dim ID_TK As Variant
   Dim MST As Variant
   
   Dim GHICHU_U As Variant
   Dim DIA_CHI_U As Variant
   Dim NGUOI_NOP_U As Variant
   Dim NGUOI_NOP As Variant
   Dim GHICHU As Variant
   Dim DIA_CHI As Variant
   'Dim NGUOI_NOP As Variant
   Dim strSQL As String
   Dim LOAI_HS As String
   Dim HTHUC_NOP As String
   Dim TRANG_THAI As String
   Dim SO_HOSO_BSUNG As String
   'Dim SO_TEP As String
    
    'sSQLCol = "DHS_MA, SO_HOSO_NHAN, TIN,TEN,DIA_CHI, NGAY_NHAN,NGUOI_NOP,NGAY_NHAP,NGUOI_NHAP,HAN_XULY,PHONG_XULY,PHONG_XULY_HIENTAI,GHI_CHU,TTHAI_HOSO,HTHUC_NOP,GUI_BD"
    

    With fpSpread1
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 4, maDTNT
        If Trim(maDTNT) = vbNullString Then
            maDTNT = "''"
        Else
            maDTNT = "'" & maDTNT & "'"
        End If
      
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
             If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        'Dia chi
        '.GetText .ColLetterToNumber("G"), 5, NGUOI_NOP
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            'NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = DateSerial(Int(Mid$(NGNOP, 7, 4)), Int(Mid$(NGNOP, 4, 2)), Int(Mid$(NGNOP, 1, 2)))
        End If
        
        .GetText .ColLetterToNumber("M"), 12, NGNHAP
        'NGNHAP = Date
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            'NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
            
        If (Trim(TAX_Utilities_Svr_New.Month) <> vbNullString Or Trim(TAX_Utilities_Svr_New.Month) <> "") And (Trim(TAX_Utilities_Svr_New.ThreeMonths) = vbNullString Or Trim(TAX_Utilities_Svr_New.ThreeMonths) = "") Then
            KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
            Tinhkykekkhaithang (Mid$(KYKKHAI, 2, 7))
        ElseIf (Trim(TAX_Utilities_Svr_New.Month) = vbNullString Or Trim(TAX_Utilities_Svr_New.Month) = "") And (Trim(TAX_Utilities_Svr_New.ThreeMonths) <> vbNullString Or Trim(TAX_Utilities_Svr_New.ThreeMonths) <> "") Then
            KYKKHAI = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
            Tinhkykekkhaiquy (Mid$(KYKKHAI, 2, 7))
        Else
            KYKK_TU_NGAY = "01/01/" & TAX_Utilities_Svr_New.Year
            KYKK_DEN_NGAY = "12/31/" & TAX_Utilities_Svr_New.Year
        End If
        'TAX_Utilities_Svr_New.ThreeMonths
        
        
       'NGNHAN = Date
        NGAY_NHAN = GetNgayNhap
        
'        If Trim(NGAY_NHAN) = vbNullString Then
'            NGAY_NHAN = "CTOD('')"
'        Else
'            'NGAY_NHAN = ToDate(Trim(NGAY_NHAN), DDMMYYYY)
'            NGAY_NHAN = "CTOD('" & format(NGAY_NHAN, "mm/dd/yyyy") & "')"
'        End If

        
        
        'MST
         .GetText .ColLetterToNumber("G"), 4, MST
         'USE
         USER = strFile(1) & "_NTKCC"
         .GetText .ColLetterToNumber("G"), 6, DIA_CHI
         DIA_CHI = TAX_Utilities_Svr_New.Convert(Trim(DIA_CHI), UNICODE, TCVN)
         'Ghi chu
         .GetText .ColLetterToNumber("M"), 14, GHICHU
         GHICHU = TAX_Utilities_Svr_New.Convert(Trim(GHICHU), UNICODE, TCVN)
         ID_TK = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")
         'LOAI_HS = changeToLoaiToKhaiQHS(ID_TK)
         DHS_MA = changeToKhaiQHS(ID_TK)
         SO_HOSO = SinhSoHoSo(DHS_MA)
         NGAY_XL = NGAY_NHAN
         .GetText .ColLetterToNumber("G"), 5, NGUOI_NOP
         NGUOI_NOP = TAX_Utilities_Svr_New.Convert(Trim(NGUOI_NOP), UNICODE, TCVN)
         NGAY_HEN = NGAY_XL
        'dhdang xu ly lay ma phong xu ly tren Form
        'ngay 05-08-2010
        If Not objTaxBusiness Is Nothing Then
        'Get Params
            PHONG_XL_X = objTaxBusiness.PHONG_XU_LY_X1
            PHONG_XL_Y = objTaxBusiness.PHONG_XU_LY_Y1
        End If
        If PHONG_XL_X <> "" And PHONG_XL_Y <> "" Then
            .GetText .ColLetterToNumber(PHONG_XL_X), PHONG_XL_Y, PHONG_XL
            PHONG_XL = Mid(PHONG_XL, InStr(1, PHONG_XL, "{") + 1, InStr(1, PHONG_XL, "}") - InStr(1, PHONG_XL, "{") - 1)
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
                 .GetText .ColLetterToNumber("O"), 2, BSUNG
                 If Trim(BSUNG) = "[X]" Then
                    TRANG_THAI = "02"
                    strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set HAN_XULY = '" & format(NGAY_XL, "mm/dd/yyyy") & "' where ID = '" & rs(0) & "'"
                    bln = clsDAO.ExecuteDLL(strSQL)
                 Else
                    TRANG_THAI = "03"
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
         sSQLVal = DHS_MA & ",'" & SO_HOSO & "','" & MST & "','" & NGUOI_NOP & "','" & DIA_CHI & "','" & KYKK_TU_NGAY & "','" & KYKK_DEN_NGAY & "','" & format(NGAY_NHAN, "mm/dd/yyyy") & "','" & NGUOI_NOP & "','" & _
            format(NGAY_NHAN, "mm/dd/yyyy") & "','" & USER & "','" & format(NGAY_XL, "mm/dd/yyyy") & "','" & format(NGAY_XL, "mm/dd/yyyy") & "','" & PHONG_XL & "','" & PHONG_XL & "','" & GHICHU & "','" & format(NGNOP, "mm/dd/yyyy") & "','" & _
           TRANG_THAI & "','" & F & "','" & HTHUC_NOP & "','" & SO_TEP & "','" & SO_HOSO_BSUNG & "','" & TRICH_YEU & "'"
       
        sSQL = "INSERT INTO QHSCC.dbo.QHS_SO_HOSO" & _
                "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
        'bln = clsDAO.ExecuteDLL(sSQL)
        
    End With
     
     
    Prepare_QLT = sSQL
    'clsDAO.Disconnect
End Function
Private Function SinhSoHoSo(DHS_MA) As String
    Dim rs As ADODB.Recordset
    Dim s, s1 As String
    Dim I As Integer
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
        I = InStrRev(s, "/")
        s1 = Right(s, Len(s) - I)
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
                 DHS_MA = "17"
            Case "05"
                 DHS_MA = "81"
            Case "38"
                 DHS_MA = "271"
            Case "54"
                 DHS_MA = "25"
            Case "09"
                 DHS_MA = "177"
            Case "02"
                 DHS_MA = "30"
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
                 DHS_MA = "16"
            Case "36"
                 DHS_MA = "350"
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
            Case Else
                MsgBox ("To khai khong ton tai")
        End Select
changeToKhaiQHS = DHS_MA
End Function
Private Function changeToLoaiToKhaiQHS(strMaToKhai) As String
Dim DHS_MA As String
Dim strSQL As String
    On Error Resume Next
    
         Select Case strMaToKhai
            Case "37"
                 DHS_MA = "200514"
            Case "37"
                 DHS_MA = "200514"
            Case "53"
                 DHS_MA = "200503"
            Case "11"
                 DHS_MA = "200201"
            Case "01"
                 DHS_MA = "200101"
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
            Case Else
                MsgBox ("To khai khong ton tai")
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
        KYKK_DEN_NGAY = Mid$(s2, 4, 2) + "/" + Mid$(s2, 1, 2) + "/" + Mid$(s2, 7, 4)
End Sub
Private Sub Insert_QHS()

On Error GoTo ErrHandle

    Dim strSQL As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String, strSQL_KHBS As String
    Dim HdrID As Variant, strDate() As String, dDate As Date
    Dim rs As ADODB.Recordset, I As Long
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
     menuId = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")
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
            If SO_TEP = "50" Then
            
            'Sinh so hieu tep
            
             s = format(NGAY_HIENTAI, "YYMM")
             s = s + DHS_MA
                
             If clsDAO.Connected = False Then
                    Me.MousePointer = vbHourglass
                    frmSystem.MousePointer = vbHourglass
                    clsDAO.CreateConnectionStringSQL spathQHSCC
                    clsDAO.Connect
                    frmSystem.MousePointer = vbDefault
                    Me.MousePointer = vbDefault
            End If
             strSQL = "Select top 1 SO_HIEU, NGAY_TAO from QHSCC.dbo.QHS_TEP_HOSO where SO_HIEU like '" & s & "%' order by ID DESC "
             Set rs = clsDAO.Execute(strSQL)
                
                If rs Is Nothing Then
                    s = s + "-1"
                Else
                    If Left$(rs(0), 4) <> format(NGAY_HIENTAI, "YYMM") Then
                        s = s + "-1"
                    Else
                        I = Right$(rs(0), Len(rs(0)) - InStr(1, rs(0), "-"))
                        I = I + 1
                        s = s & "-" & I
                    End If
                End If
                
                TEP_ID = s
            If clsDAO.Connected = False Then
                    Me.MousePointer = vbHourglass
                    frmSystem.MousePointer = vbHourglass
                    clsDAO.CreateConnectionStringSQL spathQHSCC
                    clsDAO.Connect
                    frmSystem.MousePointer = vbDefault
                    Me.MousePointer = vbDefault
            End If
            'Update QHS_SO_HOSO
            strSQL = "Update QHSCC.dbo.QHS_SO_HOSO set SO_HIEU_TEP = '" & s & "' where SO_HIEU_TEP = '' and DHS_MA = '" + DHS_MA + "' and HTHUC_NOP = '02' and NGUOI_NHAP = '" + USER + "'"
            bln = clsDAO.ExecuteDLL(strSQL)
            ' insert QHS_TEP_HOSO
            strSQL = "insert into QHSCC.dbo.QHS_TEP_HOSO (SO_HIEU, DHS_MA, KYKK_TU_NGAY, KYKK_DEN_NGAY, NGAY_TAO, SO_HOSO, TTHAI, NGUOI_TAO) values ('" & s & "', '" & DHS_MA & "', " & KYKK_TU_NGAY & ", " & KYKK_DEN_NGAY & ", '" & format(NGAY_HIENTAI, "mm/dd/yyyy") & "', '" & SO_TEP & "', '', '" & USER & "')"
            bln = clsDAO.ExecuteDLL(strSQL)
           End If
          
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
    Dim I As Integer
    
    soPl = TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 2
    For I = 1 To soPl
        If TAX_Utilities_Svr_New.NodeValidity.childNodes(I).Attributes.getNamedItem("Active").nodeValue = 1 Then
            str = str & "[" & TAX_Utilities_Svr_New.NodeValidity.childNodes(I).Attributes.getNamedItem("Caption").nodeValue & "];"
        End If
    Next
    If str <> "" Then
        str = "Phu Luc :" & str
    End If
    TinhPhuLucTk = str
End Function



