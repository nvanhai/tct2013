VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6750
      Left            =   0
      TabIndex        =   3
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
         Left            =   30
         TabIndex        =   2
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
         SpreadDesigner  =   "frmInterfaces.frx":0000
      End
      Begin MSForms.Label lblExit 
         Height          =   945
         Left            =   2760
         TabIndex        =   11
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
         TabIndex        =   8
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
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   6990
      Width           =   11535
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   30
         Width           =   1335
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   375
         Left            =   7440
         TabIndex        =   18
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "X„a"
         Size            =   "2293;661"
         Accelerator     =   88
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdViewNow 
         Height          =   375
         Left            =   6090
         TabIndex        =   17
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Xem tÍ khai"
         Size            =   "2302;661"
         Accelerator     =   72
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblBarcode 
         Height          =   255
         Left            =   180
         TabIndex        =   16
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
      Begin MSForms.Label lblWarning 
         Height          =   255
         Left            =   9120
         TabIndex        =   14
         Top             =   150
         Visible         =   0   'False
         Width           =   2325
         ForeColor       =   255
         VariousPropertyBits=   8388627
         Size            =   "4101;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVersion 
         Height          =   255
         Left            =   8610
         TabIndex        =   13
         Top             =   150
         Width           =   405
         VariousPropertyBits=   8388627
         Size            =   "714;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblLabelVersion 
         Height          =   255
         Left            =   4380
         TabIndex        =   12
         Top             =   150
         Width           =   4125
         VariousPropertyBits=   8388627
         Size            =   "7276;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblFile 
         Height          =   255
         Left            =   180
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   150
         Width           =   1785
         VariousPropertyBits=   8388627
         Size            =   "3149;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   8790
         TabIndex        =   0
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Ghi lπi"
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
         Caption         =   "Tho∏t"
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
      TabIndex        =   5
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

Private xmlDocumentInit()       As MSXML.DOMDocument
Private arrStrElements()        As String               ' array of barcode string or file name string
Private mHeaderSheet            As Integer              ' Save value of Header sheet (last sheet)
Private blnReceiveByBarcode     As Boolean                    ' Check whether form is loaded
Private objTaxBusiness          As Object               ' private business object (cls001, cls002, cls003, ...)
Private strTaxReportInfo        As String               ' Info about current tax report

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
Private verToKhai As Byte                               ' Luu cac kieu ma vach cho cac version ke khai khac nhau
Private maxBarCode As Long       ' Su dung trong truong hop to khai 04/TNCN

Private checkSoCT As Integer
Private isSheetTk As Boolean
Private checkTT As Integer

Private strMaPhongQuanLy As String
Private strTenPhongQuanLy As String

Private isTonTaiAC As Boolean ' Su dung de kiem tra xem bao cao ac co phai thay the hay khong?
Private isTKTonTai As Boolean

Private strMaSoThue, strMaDaiLyThue As String
Private isToKhaiCT As Boolean

Private isTKDA30 As Boolean  ' kiem tra QLT da co tk theo mau cu chua

Private isTKLanPS As Boolean
Private ngayPS As String

Private isToKhaiPsDaNhanTN As Boolean  ' Kiem tra cac to khai phat sinh da nhan trong ngay
' xu ly cho to khai 08, 08A/TNCN
Private isTKThang As Boolean
Private TuNgay As String
Private DenNgay As String


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
    If Not TAX_Utilities_Srv_New.Data(0) Is Nothing Then
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

    If Not TAX_Utilities_Srv_New.Data(0) Is Nothing Then
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

    Dim strSQL As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String
    Dim HdrID As Variant, strDate() As String, dDate As Date
    Dim rs As New ADODB.Recordset, i As Long
    Dim qBoSung As Variant
    Dim msgRs As MsgBoxResult
    Dim idToKhai As Integer
    
    Dim mTemp As Integer
    
    Dim dsTK_DLT As String
    'dntai them bien de luu ngay dau nam tai chinh va ngay cuoi nam tai chinh
    Dim dNgayDauNamTC As Date
    Dim dNgayCuoiNamTC As Date
    Dim varDate1 As String
    Dim varDate2 As String
    '***************************
    'Date:23/11/2005
    If TAX_Utilities_Srv_New.Data(0) Is Nothing Then Exit Sub
    '***************************
    
    blnSaveSuccess = False
    
    CallFinish
    
    
    '***************************
    'Date:02/01/2006
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepared4 dNgayDauKy
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    '***************************
     If CheckValidData = False Then
        MessageBox "0046", msOKOnly, miWarning
        Exit Sub
    End If
       
    
    ' Kiem tra trang thai 02 -> khong cho ghi, trang thai 03 canh bao van cho ghi
    If checkTT = 2 Then
        'MessageBox "0108", msOKOnly, miCriticalError
        mTemp = MessageBox("0108", msYesNo, miQuestion)
        If mTemp = mrNo Then
            Exit Sub
        End If
    End If
    ' end
    
    dsTK_DLT = "~1~2~3~4~5~6~11~12~46~47~48~49~15~16~50~51~36~70~71~72~73~74~75~80~81~82~77~86~87~89~42~43~17~59~41~76~"
    ' Kiem tra neu MDL thue khac thi canh bao
    idToKhai = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    'If IdToKhai = 1 Or IdToKhai = 2 Or IdToKhai = 4 Or IdToKhai = 11 Or IdToKhai = 12 Or IdToKhai = 46 Or IdToKhai = 47 Or IdToKhai = 48 Or IdToKhai = 49 Or IdToKhai = 15 Or IdToKhai = 16 Or IdToKhai = 50 Or IdToKhai = 51 _
    '    Or IdToKhai = 36 Or IdToKhai = 70 Or IdToKhai = 6 Or IdToKhai = 5 Then
    If InStr(1, dsTK_DLT, "~" & idToKhai & "~", vbTextCompare) > 0 Then
        If isMaDLT(strMaSoThue, strMaDaiLyThue) = False Then
            mTemp = MessageBox("0115", msYesNo, miQuestion)
            If mTemp = mrNo Then
                Exit Sub
            End If
        End If
        
        ' Kiem tra to khai BS'
        If verToKhai = 2 Then
            If isToKhaiCT = False Then
                 MessageBox "0116", msOKOnly, miWarning
                 Exit Sub
            End If
        End If
        
        ' Kiem tra xem ben QLT da co to khai theo mau cu chua
        If isTKDA30 = True Then
            MessageBox "0114", msOKOnly, miWarning
            Exit Sub
        End If
        ' end
    End If
    ' End
    ' Cac to khai PIT se khong nhan to khai co ky ke khai < thang 7 hoac quy 3
    If TAX_Utilities_Srv_New.isCheckPIT = True Then
        If idToKhai = 46 Or idToKhai = 48 Or idToKhai = 15 Or idToKhai = 50 Or idToKhai = 36 Then
            If TAX_Utilities_Srv_New.Year < 2011 Or (TAX_Utilities_Srv_New.Year = 2011 And TAX_Utilities_Srv_New.Month < 7) Then
                MessageBox "0118", msOKOnly, miWarning
                Exit Sub
            End If
        End If
        If idToKhai = 47 Or idToKhai = 49 Or idToKhai = 16 Or idToKhai = 51 Or (idToKhai = 74 And isTKThang = False) Or (idToKhai = 75 And isTKThang = False) Then
            If TAX_Utilities_Srv_New.Year < 2011 Or (TAX_Utilities_Srv_New.Year = 2011 And TAX_Utilities_Srv_New.ThreeMonths < 3) Then
                    MessageBox "0119", msOKOnly, miWarning
                Exit Sub
            End If
        End If
            
        If ((idToKhai = 74 Or idToKhai = 75) And isTKThang = True) Then
              Dim arrNgay() As String
              arrNgay = Split(TuNgay, "/")
              
          
              If Val(arrNgay(1)) < 2011 Or (Val(arrNgay(1)) = 2011 And Val(arrNgay(0)) < 7) Then
                  MessageBox "0118", msOKOnly, miWarning
                  Exit Sub
              End If
          End If
    End If
    ' end
    

    ' Tam thoi bat len message de thong bao truong hop la to khai Bo sung
    ' Truong hop to khai la version 1.3.0
    If verToKhai = 0 Then
        With fpSpread1
        
            .EventEnabled(EventAllEvents) = False
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("B"), 17, qBoSung
            ' qBoSung : neu ky lap bo lon hon ky ke khai thi set =0 nguoc lai set 1
            If qBoSung = 0 And (TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue <> "104") Then
                ' Huy bo thi quay lai man hinh quet to khai
                msgRs = MessageBox("0084", msYesNoCancel, miQuestion, 1)
                If msgRs = mrCancel Then
                    If Not TAX_Utilities_Srv_New.Data(0) Is Nothing Then
                        If Not objTaxBusiness Is Nothing Then
                            objTaxBusiness.Prepared4 dNgayDauKy
                            objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
                        End If
                        StartReceiveForm
                    End If
                    Exit Sub
                ' Neu  ghi Thay the thi set lai trang thai cua to khai la 1 va ghi binh thuong
                ElseIf msgRs = mrYes Then
                    verToKhai = 1
                ' Neu ghi Bo sung thi phai set lai tinh trang cua to khai la 2 va phai yeu cau quet phu luc KHBS
                ElseIf msgRs = mrNo Then
'                    If TAX_Utilities_Srv_New.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue = 0 Then
'                        MessageBox "0086", msOKOnly, miInformation
'                        Exit Sub
'                    Else
                        verToKhai = 2
'                    End If
                End If
            End If
            .EventEnabled(EventAllEvents) = True
        End With
    ElseIf verToKhai = 2 Then
        ' Kiem tra neu la to khai TNCN moi thi ko phai quet KHBS
        ' IdToKhai = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
        If (TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue <> "104") Then
            With fpSpread1
'                ' Neu ghi Bo sung thi phai set lai tinh trang cua to khai la 2 va phai yeu cau quet phu luc KHBS
'                If TAX_Utilities_Srv_New.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue = 0 Then
'                    MessageBox "0086", msOKOnly, miInformation
'                    Exit Sub
'                Else
                    verToKhai = 2
'                End If
            End With
        End If
        ' 04-01-2011
        ' Kiem tra neu la to khai TNCN thi set lai  trang thai bo sung thanh thay the (2->1)
        If (TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "104") Then
            verToKhai = 1
        End If
    End If
    
    ' Lay lai ID cua to khai de biet la to khai co duoc gia han thue hay khong
    idToKhai = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    ' Truong hop to khai TNCN mau 02 va 07/TNCN hien tai tu 01->05/2009 cho phep gia han thue
    ' ngoai thoi gian nay phai thong bao khong duoc gia han
    ' Do voi to khai 02, 07/TNCN thang
    If Val(idToKhai) = 15 Or Val(idToKhai) = 36 Then
        Dim varTemp
        ' Lay thong tin ve gia han nop thue TNCN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 36
            varTemp = .Value
        End With
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
        If Val(TAX_Utilities_Srv_New.Year) <> 2009 Then
            If Val(varTemp) = 1 Then
                MessageBox "0090", msOKOnly, miInformation
                Exit Sub
            End If
'        Else
'            If Val(TAX_Utilities_Srv_New.Month) > 5 Then
'                ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu lon hon thang 5 nam 2009 thi thong bao khong duoc gia han nop thue
'                If Val(varTemp) = 1 Then
'                    MessageBox "0090", msOKOnly, miInformation
'                    Exit Sub
'                End If
'            End If
        End If
    End If
    
    ' Do voi to khai 02/TNCN quy
    If Val(idToKhai) = 37 Then
        ' Lay thong tin ve gia han nop thue TNCN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 36
            varTemp = .Value
        End With
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
        If Val(TAX_Utilities_Srv_New.Year) <> 2009 Then
            If Val(varTemp) = 1 Then
                MessageBox "0090", msOKOnly, miInformation
                Exit Sub
            End If
'        Else
'            If Val(TAX_Utilities_Srv_New.ThreeMonths) > 1 Then
'                ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu lon hon thang 5 nam 2009 thi thong bao khong duoc gia han nop thue
'                If Val(varTemp) = 1 Then
'                    MessageBox "0091", msOKOnly, miInformation
'                    Exit Sub
'                End If
'            End If
        End If
    End If
    
    ' Truong hop to khai TNDN va quyet toan hien tai 2009 cho phep gia han thue
    ' ngoai thoi gian nay phai thong bao khong duoc gia han
    ' Do voi to khai 01A, 01B/TNDN thang
    If Val(idToKhai) = 11 Or Val(idToKhai) = 12 Then
        ' Lay thong tin ve gia han nop thue TNDN
        'dntai 06/03/2012 set lai vi tri cell check gia han
        If Val(idToKhai) = 11 Then
            With fpSpread1
                .Sheet = 1
                .Col = .ColLetterToNumber("E")
                .Row = 37
                varTemp = .Value
            End With
        ElseIf Val(idToKhai) = 12 Then
            With fpSpread1
                .Sheet = 1
                .Col = .ColLetterToNumber("E")
                .Row = 38
                varTemp = .Value
            End With
        End If
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
        ' yeu cau mo tat ca cac ky
'        If Val(TAX_Utilities_Srv_New.Year) <> 2009 And Val(TAX_Utilities_Srv_New.Year) <> 2010 And Val(TAX_Utilities_Srv_New.Year) <> 2011 Then
'            If Val(varTemp) = 1 Then
'                MessageBox "0092", msOKOnly, miInformation
'                Exit Sub
'            End If
'        End If
        ' Kiem tra quy truoc da co gia han canh bao NSD check vao gia han nop thue
        ' nvhai
        If Val(TAX_Utilities_Srv_New.Year) = 2009 Or Val(TAX_Utilities_Srv_New.Year) = 2010 Or Val(TAX_Utilities_Srv_New.Year) = 2011 Then
            If Val(varTemp) = 0 Then
                'lay ngay dau nam tc va ngay ket thuc nam tc de loc de lieu

                dNgayDauNamTC = GetNgayDauNam(TAX_Utilities_Srv_New.Year, iThangTaiChinh, iNgayTaiChinh)
                'set lai dinh dang de so sanh
                varDate1 = "'" & DatePart("D", dNgayDauNamTC) & "-" & MonthName(DatePart("M", dNgayDauNamTC), True) & "-" & DatePart("YYYY", dNgayDauNamTC) & "'"
                dNgayCuoiNamTC = NgayCuoiNamTaiChinh(TAX_Utilities_Srv_New.Year, iThangTaiChinh, iNgayTaiChinh)
                'set lai dinh dang de so sanh
                varDate2 = "'" & DatePart("D", dNgayCuoiNamTC) & "-" & MonthName(DatePart("M", dNgayCuoiNamTC), True) & "-" & DatePart("YYYY", dNgayCuoiNamTC) & "'"
                'connect to database QLT
                If Not clsDAO.Connected Then
                    clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
                    clsDAO.Connect
                End If
                ' SQL check du lieu
                strSQL = "select ID from RCV_TKHAI_HDR "
                strSQL = strSQL & " where TIN='" & strMST & "' "
                strSQL = strSQL & "and KYKK_TU_NGAY >= " & varDate1 & " and KYKK_DEN_NGAY < " & varDate2
                strSQL = strSQL & " and CO_GHAN='Y' "
                strSQL = strSQL & " and LOAI_TKHAI='" & changeMaToKhai(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue) & "' "
                Set rs = clsDAO.Execute(strSQL)
            
                If (Not rs Is Nothing) And rs.Fields.Count > 0 Then
                    If Year(dNgayDauKy) < 2011 Or (Year(dNgayDauKy) = 2011 And DatePart("Q", dNgayDauKy) < 4) Then
                    If MessageBox("0098", msYesNo, miQuestion) = mrYes Then
'                        With fpSpread1
'                            .Sheet = 1
'                            .Col = .ColLetterToNumber("E")
'                            .Row = 17
'                            .Value = 1
'                        End With
                        Exit Sub
                        End If
                    End If
                    'DisplayMessage "0098", msOKOnly, miInformation
                    'Exit Sub
                End If
                Set rs = Nothing
            End If
        End If
        
        ' end nvhai
        
    End If
    ' Do voi to khai 05/TNDN thang
    If Val(idToKhai) = 14 Then
        ' Lay thong tin ve gia han nop thue TNDN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 19
            varTemp = .Value
        End With
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
'        If Val(TAX_Utilities_Srv_New.Year) <> 2009 Then
'            If Val(varTemp) = 1 Then
'                MessageBox "0092", msOKOnly, miInformation
'                Exit Sub
'            End If
'        End If
    End If
    ' Do voi to khai 03/TNDN thang
    If Val(idToKhai) = 3 Then
        ' Lay thong tin ve gia han nop thue TNDN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 37
            varTemp = .Value
        End With
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
        ' mo gia han cho cac ky ke khai
'        If Val(TAX_Utilities_Srv_New.Year) <> 2009 And Val(TAX_Utilities_Srv_New.Year) <> 2010 And Val(TAX_Utilities_Srv_New.Year) <> 2011 Then
'            If Val(varTemp) = 1 Then
'                MessageBox "0092", msOKOnly, miInformation
'                Exit Sub
'            End If
'        End If
        
        ' Kiem tra to khai quy da co gia han canh bao NSD check vao gia han nop thue tren to khai quyet toan nam
        ' nvhai
        If Val(TAX_Utilities_Srv_New.Year) = 2009 Or Val(TAX_Utilities_Srv_New.Year) = 2010 Or Val(TAX_Utilities_Srv_New.Year) = 2011 Then
            If Val(varTemp) = 0 Then
                'lay ngay dau nam tc va ngay ket thuc nam tc de loc de lieu
                dNgayDauNamTC = DateSerial(TAX_Utilities_Srv_New.Year, iThangTaiChinh, iNgayTaiChinh)
                'set lai dinh dang de so sanh
                varDate1 = "'" & DatePart("D", dNgayDauNamTC) & "-" & MonthName(DatePart("M", dNgayDauNamTC), True) & "-" & DatePart("YYYY", dNgayDauNamTC) & "'"
                dNgayCuoiNamTC = NgayCuoiNamTaiChinh(TAX_Utilities_Srv_New.Year, iThangTaiChinh, iNgayTaiChinh)
                'set lai dinh dang de so sanh
                varDate2 = "'" & DatePart("D", dNgayCuoiNamTC) & "-" & MonthName(DatePart("M", dNgayCuoiNamTC), True) & "-" & DatePart("YYYY", dNgayCuoiNamTC) & "'"
                'connect to database QLT
                If Not clsDAO.Connected Then
                    clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
                    clsDAO.Connect
                End If
                ' SQL check du lieu
                strSQL = "select ID from RCV_TKHAI_HDR "
                strSQL = strSQL & " where TIN='" & strMST & "' "
                strSQL = strSQL & "and KYKK_TU_NGAY >= " & varDate1 & " and KYKK_DEN_NGAY < " & varDate2
                strSQL = strSQL & " and CO_GHAN='Y' "
                strSQL = strSQL & " and ( LOAI_TKHAI='01A_TNDN' Or LOAI_TKHAI='01B_TNDN' OR LOAI_TKHAI='01A_TNDN11' Or LOAI_TKHAI='01B_TNDN11') "
                Set rs = clsDAO.Execute(strSQL)
            
                If (Not rs Is Nothing) And rs.Fields.Count > 0 Then
                    If Year(dNgayDauKy) <= 2011 Then
                    If MessageBox("0099", msYesNo, miQuestion) = mrYes Then
'                        With fpSpread1
'                            .Sheet = 1
'                            .Col = .ColLetterToNumber("E")
'                            .Row = 17
'                            .Value = 1
'                        End With
                        Exit Sub
                        End If
                    End If
                    'DisplayMessage "0098", msOKOnly, miInformation
                    'Exit Sub
                End If
                Set rs = Nothing
            End If
        End If
        
        ' end nvhai
        
        
        
    End If
    
    ' Kiem tra gia han to khai 01/GTGT
    ' yeu cau mo gia han tat ca cac ky
'    If Val(idToKhai) = 1 Then
'         ' Lay thong tin ve gia han nop thue GTGT
'        With fpSpread1
'            .Sheet = 1
'            .Col = .ColLetterToNumber("E")
'            .Row = 38
'            varTemp = .Value
'        End With
'        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2012 thi thong bao khong duoc gia han nop thue
'        If Val(TAX_Utilities_Srv_New.Year) = 2012 And (Val(TAX_Utilities_Srv_New.Month) = 4 Or Val(TAX_Utilities_Srv_New.Month) = 5 Or Val(TAX_Utilities_Srv_New.Month) = 6) Then
'        Else
'            If Val(varTemp) = 1 Then
'                MessageBox "0128", msOKOnly, miInformation
'                Exit Sub
'            End If
'        End If
'    End If
    
    
    If clsDAO.Connected = False Then
        Me.MousePointer = vbHourglass
        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
        frmSystem.MousePointer = vbDefault
        Me.MousePointer = vbDefault
    End If
    '*************************************************************
    'Date: 25/02/2006
    'Kiem tra khoa so
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlKhoaSo")
    Set rs = clsDAO.Execute(strSQL)
    
    If (Not rs Is Nothing) And rs.Fields.Count > 0 Then
        If objTaxBusiness.KiemTraKhoaSo(rs.Fields(0)) Then
            DisplayMessage "0070", msOKOnly, miInformation
            Exit Sub
        End If
    End If
    Set rs = Nothing
    '*************************************************************
    
    ' xu ly nhan cac mau an chi co canh bao khi quet trung
    If Val(idToKhai) <= 68 And Val(idToKhai) >= 64 Then
        ' xu ly an chi
        If isTonTaiAC = True Then
            mResult = MessageBox("0047", msYesNo, miQuestion)
            If mResult = mrNo Then
                Exit Sub
            End If
        End If
    Else
    strSQL = "select ID, DA_NHAN from RCV_TKHAI_HDR "
    strSQL = strSQL & " where TIN='" & strMST & "' "
    strSQL = strSQL & " and LOAI_TKHAI='" & changeMaToKhai(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue) & "' "
    
    'Ngay dau ky ke khai va ngay cuoi ky ke khai
    dDate = dNgayDauKy
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
        dDate = DateAdd("m", 1, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
        strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
        dDate = DateAdd("m", 3, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
        strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
        dDate = DateAdd("m", 12, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    End If
        Dim flgBCTC As Boolean
    clsDAO.BeginTrans
    Set rs = clsDAO.Execute(strSQL)
    
    ' nvhai
    ' Xu ly cho nhan BCTC in bang HTKK 2.1.0
    
    If rs.Fields.Count > 0 Then
        ' nvhai
        ' Neu la ID cua cac BCTC in bang HTKK 2.1.0
        ' begin
        If (Val(idToKhai) = 24 Or Val(idToKhai) = 25 Or Val(idToKhai) = 26 Or Val(idToKhai) = 27 Or Val(idToKhai) = 28 Or Val(idToKhai) = 29 _
            Or Val(idToKhai) = 30 Or Val(idToKhai) = 31 Or Val(idToKhai) = 32 Or Val(idToKhai) = 33 Or Val(idToKhai) = 34 Or Val(idToKhai) = 35 _
            Or Val(idToKhai) = 55 Or Val(idToKhai) = 56 Or Val(idToKhai) = 57 Or Val(idToKhai) = 58) Then
            flgBCTC = True
            If verToKhai = 0 Then ' Trong truong hop to khai thay the nhung ke khai ko su dung KHBS de ke khai ma su dung chuc nang ke khai goc
                mResult = MessageBox("0047", msYesNo, miQuestion)
                If mResult = mrYes Then ' Neu dong y ghi la to khai thay the thi phai dat lai trang thai = 1
                    verToKhai = 1
                    If UCase(rs.Fields("DA_NHAN").Value) = "E" Then
                        clsDAO.ExecuteQuery "delete from RCV_TKHAI_HDR where ID='" & rs(0).Value & "'"
                        clsDAO.ExecuteQuery "delete from RCV_TKHAI_DTL where HDR_ID='" & rs(0).Value & "'"
                    End If
                    'clsDAO.Execute "delete from RCV_TKHAI_HDR where ID='" & rs(0).Value & "'"
                    
                Else
                    '********************
                    clsDAO.CommitTrans
                    '********************
                    Exit Sub
                End If
            ElseIf verToKhai = 2 Then
                If UCase(rs.Fields("DA_NHAN").Value) = "E" Then
                    clsDAO.ExecuteQuery "delete from RCV_TKHAI_HDR where ID='" & rs(0).Value & "'"
                    clsDAO.ExecuteQuery "delete from RCV_TKHAI_DTL where HDR_ID='" & rs(0).Value & "'"
                End If
            End If
        Else
        ' end
            If verToKhai = 0 And isTKTonTai = True Then ' Trong truong hop to khai thay the nhung ke khai ko su dung KHBS de ke khai ma su dung chuc nang ke khai goc
                mResult = MessageBox("0047", msYesNo, miQuestion)
                If mResult = mrYes Then ' Neu dong y ghi la to khai thay the thi phai dat lai trang thai = 1
                    verToKhai = 1
                    If UCase(rs.Fields("DA_NHAN").Value) = "E" Then
                        clsDAO.ExecuteQuery "delete from RCV_TKHAI_HDR where ID='" & rs(0).Value & "'"
                        clsDAO.ExecuteQuery "delete from RCV_TKHAI_DTL where HDR_ID='" & rs(0).Value & "'"
                    End If
                    'clsDAO.Execute "delete from RCV_TKHAI_HDR where ID='" & rs(0).Value & "'"
                Else
                    '********************
                    clsDAO.CommitTrans
                    '********************
                    Exit Sub
                End If
            ElseIf verToKhai = 2 Then
                If UCase(rs.Fields("DA_NHAN").Value) = "E" Then
                    clsDAO.ExecuteQuery "delete from RCV_TKHAI_HDR where ID='" & rs(0).Value & "'"
                    clsDAO.ExecuteQuery "delete from RCV_TKHAI_DTL where HDR_ID='" & rs(0).Value & "'"
                End If
            End If
        End If
    End If
        ' su loi ghi du lieu hdr va dtl tren transaction khac nhau
        clsDAO.CommitTrans
        ' end xu ly to khai binh thuong
    End If
    
    ' xu ly cho 2 to khai 08, 08A/TNCN
    If idToKhai = "74" Or idToKhai = "75" Then
        If verToKhai = 0 And isTKTonTai = True Then ' Trong truong hop to khai thay the nhung ke khai ko su dung KHBS de ke khai ma su dung chuc nang ke khai goc
                mResult = MessageBox("0047", msYesNo, miQuestion)
                If mResult = mrYes Then ' Neu dong y ghi la to khai thay the thi phai dat lai trang thai = 1
                    verToKhai = 1
                End If
        End If
    End If
    
    
    Set rs = Nothing
    

    If idToKhai = 2 Or idToKhai = 4 Or idToKhai = 46 Or idToKhai = 47 Or idToKhai = 48 Or idToKhai = 49 Or idToKhai = 15 Or idToKhai = 16 Or idToKhai = 50 Or idToKhai = 51 _
    Or idToKhai = 36 Or idToKhai = 6 Or idToKhai = 72 Or idToKhai = 87 Or idToKhai = 86 Or idToKhai = 77 Or idToKhai = 71 Or idToKhai = 74 Or idToKhai = 89 Or idToKhai = 42 Or idToKhai = 43 Or idToKhai = 17 Or idToKhai = 59 Or idToKhai = 41 Or idToKhai = 76 Then
        strSQL_HDR = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlHdrTT28").nodeValue)
    ElseIf idToKhai = 1 Or idToKhai = 11 Or idToKhai = 12 Or idToKhai = 5 Or idToKhai = 70 Or idToKhai = 80 Or idToKhai = 81 Or idToKhai = 82 Or idToKhai = 3 Or idToKhai = 73 Then
        strSQL_HDR = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlHdrTT28_NNKD").nodeValue)
    Else
        strSQL_HDR = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlHdr").nodeValue)
    End If
    ' xu ly de ghi cac mau an chi
    If Val(idToKhai) = 66 Or Val(idToKhai) = 68 Or Val(idToKhai) = 67 Or Val(idToKhai) = 64 Or Val(idToKhai) = 65 Then
        strSQL_DTL = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlDtl_AC").nodeValue)
    Else
    strSQL_DTL = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlDtl").nodeValue)
    End If
    Set rs = clsDAO.Execute("select RCV_XLTK_HDR_SEQ.NEXTVAL from dual")
    HdrID = rs(0).Value
    

    
    For i = 0 To TAX_Utilities_Srv_New.NodeValidity.childNodes.length - 1
        clsDAO.BeginTrans
        If Val(TAX_Utilities_Srv_New.NodeValidity.childNodes(i).Attributes.getNamedItem("Active").nodeValue) > 0 Then
            If i = 0 Then
                ' Kiem tra xem chkQuetBangKe neu = true thi day la quet them bang ke
                If (frmSystem.chkQuetBangKe.Value = True) Then
                    clsDAO.Execute objTaxBusiness.GenerateSQL_Header(TAX_Utilities_Srv_New.Data(i), strSQL_HDR, HdrID, verToKhai, dNgayDauKy, True)
                Else ' Day la quet to khai thuan tuy
                    clsDAO.Execute objTaxBusiness.GenerateSQL_Header(TAX_Utilities_Srv_New.Data(i), strSQL_HDR, HdrID, verToKhai, dNgayDauKy)
                End If
                ' HDR va DTL ghi tren cung 1 transaction
                'clsDAO.CommitTrans
            End If
                GenerateSQL_Details TAX_Utilities_Srv_New.Data(i), strSQL_DTL, HdrID, i
        End If
        clsDAO.CommitTrans
    Next

    '***************************
    ' Clear data
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

    SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    MessageBox "0049", msOKOnly, miCriticalError
    
    On Error GoTo ExitErr
    'Rollback
    clsDAO.RollbackTrans
    With clsDAO
        .BeginTrans
        .Execute "delete from RCV_TKHAI_DTL where HDR_ID ='" & HdrID & "'"
        .Execute "delete from RCV_TKHAI_HDR where ID='" & HdrID & "'"
        .CommitTrans
    End With
    
    Set rs = Nothing
    blnSaveSuccess = True
    Exit Sub
ExitErr:
    Set rs = Nothing
    SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    MessageBox "0049", msOKOnly, miCriticalError
    blnSaveSuccess = True
End Sub

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
    
    Dim strLoaiTK As String
    ' Phien ban 1.3.1, Danh dau ma vach cua to khai Bo sung la TT (verToKhai = 2)
'    If verToKhai = 2 Then
'        MessageBox "0085", msOKOnly, miInformation
'        cmdViewNow.Enabled = False
'        Exit Sub
'    End If
    'strBarcode = TAX_Utilities_Srv_New.Convert(strBarcode, UNICODE, TCVN)
    If verToKhai = 2 Then
        strLoaiTK = "bs"
    Else
        strLoaiTK = "aa"
    End If
    
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
            arrBCNew(counter) = strLoaiTK & strPrefix & strBarcodeCount & strBarcode
        End If
    Next
    ' Neu chua quet to khai ma co yeu cau hien thi thi thong bao phai quet to khai
    If chkToKhai = False Then
        DisplayMessage "0083", msOKOnly, miCriticalError
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
   Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String, str9 As String, str10 As String
   Dim str11 As String, str12 As String, str13 As String, str14 As String, str15 As String, str16 As String, str17 As String, str18 As String, str19 As String, str20 As String
   Dim str21 As String, str22 As String, str23 As String, str24 As String, str25 As String, str26 As String, str27 As String, str28 As String, str29 As String, str30 As String
   Dim str31 As String, str32 As String, str33 As String, str34 As String, str35 As String, str36 As String, str37 As String, str38 As String, str39 As String, str40 As String
   Dim str41 As String, str42 As String, str43 As String, str44 As String, str45 As String, str46 As String, str47 As String, str48 As String, str49 As String, str50 As String
   Dim str51 As String, str52 As String, str53 As String
  

' str1 = "aa256684400114520   04201000300300100101/0401/01/2009<S01><S>~~01/10/2010~31/12/2010</S><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT2/200~AA/12T~34~0000012~0000045~~~~~0~0~0~~0~~0~~0000012~0000045~34~0</S><S>~Nguy‘n V®n An~27/06/2011</S></S01>"
'
''str1 = "aa256504400114520   05201100200200100101/0401/01/2010<S01><S>40000000~2000000~30000000~30000~0~0~0~0~0~0</S><S>Nguy‘n V®n An~27/06/2011~1~~</S></S01>"
'     Barcode_Scaned str1

'str1 = "aa300462300100633   06201100100100100101/0101/01/2010<S01><S>0201021271001</S><S>10000000~2000000~1000000</S><S>ui87o97o~adsasd~asdasdad~19/07/2011~1~~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa300472300100633   01201100100100100101/0101/01/2010<S01><S></S><S>3000000~2000000~1000000</S><S>ui87o97o~adsasd~asdasdad~19/07/2011~1~~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa300482300100633   06201100200200100101/0101/01/2010<S01><S>0201021271001</S><S>20000000~10000000~5000000</S><S>ui87o97o~adsasd~asdasdad~19/07/2011~1~~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa300492300100633   01201100200200100101/0101/01/2010<S01><S>0201021271001</S><S>30000000~1000000~200000</S><S>ui87o97o~adsasd~asdasdad~19/07/2011~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa300152300100633   06201100300300100101/0101/01/2010<S01><S>0201021271001</S><S>1000~330~1000000000~200000000~50000000~10000000~20000000~0~0~2000000~0</S><S>ui87o97o~13/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa300162300100633   01201100100100100101/0101/01/2010<S01><S>0201021271001</S><S>1000000000~200000000~400000000~200000000~200000000~0~30000000~20000000~500000~3000000~4000000</S><S>ui87o97o~06/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str1
'str1 = "aa300152300100633   06201100400400100101/0101/01/2010<S01><S>0201021271001</S><S>20000000~10000000~40000000~15000000~1200000~4000000~200000~13000~250000~20000~2600</S><S>ui87o97o~20/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa300162300100633   03201100100100100101/0101/01/2010<S01><S>0201021271001</S><S>4000~2000~120000000~30000000~4000000~200000~100000~500000~200000~10000~100000</S><S>ui87o97o~20/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa300502300100633   06201100200200100101/0101/01/2010<S01><S>0201021271001</S><S>300000000~15000000~2000000~2000~4000000~200000~1000000~100000~20000000~100000</S><S>ui87o97o~20/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa300512300100633   03201100100100100101/0101/01/2010<S01><S>0201021271001</S><S>30000000~1500000~400000000~400000~500000000~25000000~60000000~6000000~7000000000~700000</S><S>ui87o97o~20/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str1

'str1 = "aa300363602400346   06201100400400100101/0101/01/2010<S07><S>0201021271001</S><S>~10000000000~118000000~4000000~64000000~40000000~10000000~9882000000~3448850000~20000000~3448850~3445401150~30000000~20%~6000000~10000~1~100</S><S>ui87o97o~20/07/2011~test~asdasdad~1~~</S></S07>"
'Barcode_Scaned str1

'str1 = "aa300462300100633   07201100700700100101/0101/01/2010<S01><S>6868686868</S><S>200000000~10000000~5000000</S><S>ui87o97o~~asdasdad~28/07/2011~1~~</S></S01>"
'Barcode_Scaned str1
' 01A_TNCN_BH
'str2 = "aa300463600278732   07201100200300100101/0101/01/2010<S01><S>6868686868</S><S>1000000~200000~10000</S><S>~test~asdasdad~08/08/2011~1~~</S></S01>"
'Barcode_Scaned str2
' 01B_TNCN_BH
'str2 = "aa300473600278732   03201100400400100101/0101/01/2010<S01><S>6868686868</S><S>10000000~2000000~100000</S><S>~test~asdasdad~08/08/2011~1~~</S></S01>"
'Barcode_Scaned str2
'01A_TNCN_XS
'str2 = "aa300483600278732   07201100200200100101/0101/01/2010<S01><S>6868686868</S><S>100000000~10000000~2000000</S><S>~test~asdasdad~08/08/2011~1~~</S></S01>"
'Barcode_Scaned str2
'01B_TNCN_XS
'str2 = "aa300493600278732   03201100400400100101/0101/01/2010<S01><S>6868686868</S><S>1000000000~100000000~2000000</S><S>~test~asdasdad~08/08/2011~1~~</S></S01>"
'Barcode_Scaned str2
'02A_TNCN
'str2 = "aa300153600278732   07201100100100100101/0101/01/2010<S01><S>6868686868</S><S>100000000~10000000~20000000~2000000~10000000~10000000~2000000~1000000~200000~200000~200000</S><S>~08/08/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str2
'02B_TNCN
'str2 = "aa300163600278732   03201100300300100101/0101/01/2010<S01><S>6868686868</S><S>10000000000~1000000000~100000000~10000000~100000000~20000000~2000000~330000~200000~200000~66000</S><S>~08/08/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str2
'03A_TNCN
'str2 = "aa300503600278732   07201100200200100101/0101/01/2010<S01><S>6868686868</S><S>10000000~500000~2000000~2000~1000000~50000~100000~10000~10000000~10000</S><S>~08/08/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str2
'03B_TNCN
'str2 = "aa300513600278732   03201100200200100101/0101/01/2010<S01><S>6868686868</S><S>30000000~1500000~400000000~400000~500000000~25000000~60000000~6000000~7000000000~700000</S><S>ui87o97o~20/07/2011~test~asdasdad~1~~</S></S01>"
'Barcode_Scaned str2
'07_TNCN
'str2 = "aa316013600247325   02201300200200100101/0114/06/2006<S01><S></S><S>~35329039~0~100000000~1000000~0~0~100000000~0~0~100000000~0~0~0~100000000~99000000~0~0~0~63670961~0~63670961~0~0~0</S><S>~~~01/03/2013~1~~~1701~x~03</S></S01>"
'Barcode_Scaned str2

'str2 = "aa316113600247325   04201200500500100101/0114/06/2006<S01><S></S><S>100000000~0~100000000~0~0~100000000~0~0~100000000~25~0~25000000~x</S><S></S><S>~~~27/02/2013~1~0~~1052~01</S></S01>"
'Barcode_Scaned str2
'str2 = "aa316033600247325   00201201001000100301/0114/06/200601/01/201231/12/2012<S03><S></S><S>10000000~0~0~0~0~0~0~0~0~0~0~0~10000000~10000000~0~10000000~0~100000~9900000~0~"
'Barcode_Scaned str2
'str2 = "aa316033600247325   0020120100100020039900000~2475000~0~0~0~2475000~0~2475000~2475000~0</S><S></S><S>fgdfg~fgdfgdfg</S><S>x</S><S>~~~27/02/2013~1~1~0~1052~</S></S03>"
'Barcode_Scaned str2
'str2 = "aa316033600247325   002012010010003003<S03-2A><S>2007~0~0~0~0~2008~0~0~0~0~2009~0~0~0~0~2010~0~0~0~0~2011~150000~50000~100000~0~150000~50000~100000~0</S></S03-2A>"
'Barcode_Scaned str2
'str2 = "aa316013600247325   02201300200200100101/0114/06/2006<S01><S></S><S>~35329039~0~100000000~1000000~0~0~100000000~0~0~100000000~0~0~0~100000000~99000000~0~0~0~63670961~0~63670961~0~0~0</S><S>~~~01/03/2013~1~~~1701~x~03</S></S01>"
'Barcode_Scaned str2
'str2 = "aa316113600247325   04201200700700100101/0114/06/2006<S01><S></S><S>100000000~0~100000000~0~0~100000000~0~0~100000000~25~0~25000000~x</S><S></S><S>~~~27/02/2013~1~0~~1052~02</S></S01>"
'Barcode_Scaned str2
'str2 = "aa316123600247325   04201200900900100101/0114/06/2006<S01><S></S><S>~0~0~0~0~0~~0~0~0~0~0~0~x</S><S>~27/02/2013~~~1~~1052~03</S></S01>"
'Barcode_Scaned str2
'str2 = "aa316033600247325   00201201401400100301/0114/06/200601/01/201231/12/2012<S03><S></S><S>10000000~0~0~0~0~0~0~0~0~0~0~0~10000000~10000000~0~10000000~0~100000~9900000~0~9"
'Barcode_Scaned str2
'str2 = "aa316033600247325   002012014014002003900000~2475000~0~0~0~2475000~0~2475000~2475000~0</S><S></S><S>fgdfg~fgdfgdfg</S><S>x</S><S>~~~27/02/2013~1~1~0~1052~02</S></S03>"
'Barcode_Scaned str2
'str2 = "aa316033600247325   002012014014003003<S03-2A><S>2007~0~0~0~0~2008~0~0~0~0~2009~0~0~0~0~2010~0~0~0~0~2011~150000~50000~100000~0~150000~50000~100000~0</S></S03-2A>"
'Barcode_Scaned str2

str2 = "aa316681400633697   02201101301400100101/0101/01/2009<S01><S>~~01/04/2011~30/06/2011</S><S>BC01-2L~01GTKT-2LN-0"
str2 = str2 & "1~XQ/2008T~48~0036603~0036650~~~0036603~0036650~48~0~0~~0~~48~36603-36650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~31~0036670~00"
str2 = str2 & "36700~~~0036670~0036700~31~0~0~~0~~31~36670-36700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~17~0036734~0036750~~~0036734~0036750~"
str2 = str2 & "17~0~0~~0~~17~36734-36750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~33~0036768~0036800~~~0036768~0036800~33~0~0~~0~~33~36768-3680"
str2 = str2 & "0~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0036851~0036900~~~0036851~0036900~50~0~0~~0~~50~36851-36900~~~0~0~BC01-2L~01GTKT-2"
str2 = str2 & "LN-01~XQ/2008T~47~0036904~0036950~~~0036904~0036950~47~0~0~~0~~47~36904-36950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~43~003695"
str2 = str2 & "8~0037000~~~0036958~0037000~43~0~0~~0~~43~36958-37000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0037001~0037050~~~0037001~0037"
str2 = str2 & "050~50~0~0~~0~~50~37001-37050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~450~0037051~0037500~~~0037051~0037500~450~0~0~~0~~450~370"
str2 = str2 & "51-37500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~450~0044051~0044500~~~0044051~0044500~450~0~0~~0~~450~44051-44500~~~0~0~BC01-2"
str2 = str2 & "L~01GTKT-2LN-01~XQ/2008T~35~0037916~0037950~~~0037916~0037950~35~0~0~~0~~35~37916-37950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T"
str2 = str2 & "~50~0043851~0043900~~~0043851~0043900~50~0~0~~0~~50~43851-43900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~47~0038004~0038050~~~00"
str2 = str2 & "38004~0038050~47~0~0~~0~~47~38004-38050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~48~0038053~0038100~~~0038053~0038100~48~0~0~~0~"
str2 = str2 & "~48~38053-38100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~42~0037959~0038000~~~0037959~0038000~42~0~0~~0~~42~37959-38000~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XQ/2008T~48~0038103~0038150~~~0038103~0038150~48~0~0~~0~~48~38103-38150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2"
str2 = str2 & "008T~23~0038478~0038500~~~0038478~0038500~23~0~0~~0~~23~38478-38500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~36~0037765~0037800~"
str2 = str2 & "~~0037765~0037800~36~0~0~~0~~36~37765-37800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~9~0043942~0043950~~~0043942~0043950~9~0~0~~"
str2 = str2 & "0~~9~43942-43950~~~0~0~BC01"
str2 = str2 & ""
str2 = str2 & "-2L~01GTKT-2LN-01~XQ/2008T~16~0043585~0043600~~~0043585~0043600~16~0~0~~0~~16~43585-43600~~~0~0~BC01-2L~01GTKT-"
str2 = str2 & "2LN-01~XQ/2008T~27~0043674~0043700~~~0043674~0043700~27~0~0~~0~~27~43674-43700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~18~00437"
str2 = str2 & "83~0043800~~~0043783~0043800~18~0~0~~0~~18~43783-43800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0043951~0044000~~~0043951~004"
str2 = str2 & "4000~50~0~0~~0~~50~43951-44000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0044001~0044050~~~0044001~0044050~50~0~0~~0~~50~44001"
str2 = str2 & "-44050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~21~0039580~0039600~~~0039580~0039600~21~0~0~~0~~21~39580-39600~~~0~0~BC01-2L~01G"
str2 = str2 & "TKT-2LN-01~XQ/2008T~50~0039601~0039650~~~0039601~0039650~50~0~0~~0~~50~39601-39650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0"
str2 = str2 & "039651~0039700~~~0039651~0039700~50~0~0~~0~~50~39651-39700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0039701~0039750~~~0039701"
str2 = str2 & "~0039750~50~0~0~~0~~50~39701-39750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~49~0039752~0039800~~~0039752~0039800~49~0~0~~0~~49~3"
str2 = str2 & "9752-39800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~45~0040106~0040150~~~0040106~0040150~45~0~0~~0~~45~40106-40150~~~0~0~BC01-2L"
str2 = str2 & "~01GTKT-2LN-01~XQ/2008T~50~0040151~0040200~~~0040151~0040200~50~0~0~~0~~50~40151-40200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~"
str2 = str2 & "50~0040201~0040250~~~0040201~0040250~50~0~0~~0~~50~40201-40250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~6~0040045~0040050~~~0040"
str2 = str2 & "045~0040050~6~0~0~~0~~6~40045-40050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~36~0040315~0040350~~~0040315~0040350~36~0~0~~0~~36~"
str2 = str2 & "40315-40350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~43~0040258~0040300~~~0040258~0040300~43~0~0~~0~~43~40258-40300~~~0~0~BC01-2"
str2 = str2 & "L~01GTKT-2LN-01~XQ/2008T~30~0040071~0040100~~~0040071~0040100~30~0~0~~0~~30~40071-40100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T"
str2 = str2 & "~32~0039969~0040000~~~0039969~0040000~32~0~0~~0~~32~39969-40000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~150~0040351~0040500~~~0"
str2 = str2 & "040351~0040500~150~0~0~~0~~150~40351-40500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~35~0035566~0035600~~~0035566~0035"
str2 = str2 & ""
str2 = str2 & "600~35~0~0~~0~~35~35566-35600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~5~0035746~0035750~~~0035746~0035750~5~0~0~~0"
str2 = str2 & "~~5~35746-35750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~46~0035855~0035900~~~0035855~0035900~46~0~0~~0~~46~35855-35900~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XQ/2008T~48~0035903~0035950~~~0035903~0035950~48~0~0~~0~~48~35903-35950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2"
str2 = str2 & "008T~49~0035952~0036000~~~0035952~0036000~49~0~0~~0~~49~35952-36000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~42~0036009~0036050~"
str2 = str2 & "~~0036009~0036050~42~0~0~~0~~42~36009-36050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~28~0036073~0036100~~~0036073~0036100~28~0~0"
str2 = str2 & "~~0~~28~36073-36100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~32~0036119~0036150~~~0036119~0036150~32~0~0~~0~~32~36119-36150~~~0~"
str2 = str2 & "0~BC01-2L~01GTKT-2LN-01~XQ/2008T~35~0036166~0036200~~~0036166~0036200~35~0~0~~0~~35~36166-36200~~~0~0~BC01-2L~01GTKT-2LN-01~"
str2 = str2 & "XQ/2008T~43~0036208~0036250~~~0036208~0036250~43~0~0~~0~~43~36208-36250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~38~0036263~0036"
str2 = str2 & "300~~~0036263~0036300~38~0~0~~0~~38~36263-36300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~200~0036301~0036500~~~0036301~0036500~2"
str2 = str2 & "00~0~0~~0~~200~36301-36500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0040501~0040550~~~0040501~0040550~50~0~0~~0~~50~40501-405"
str2 = str2 & "50~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~950~0040551~0041500~~~0040551~0041500~950~0~0~~0~~950~40551-41500~~~0~0~BC01-2L~01GT"
str2 = str2 & "KT-2LN-01~XQ/2008T~40~0041511~0041550~~~0041511~0041550~40~0~0~~0~~40~41511-41550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~4~004"
str2 = str2 & "1597~0041600~~~0041597~0041600~4~0~0~~0~~4~41597-41600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~33~0041668~0041700~~~0041668~004"
str2 = str2 & "1700~33~0~0~~0~~33~41668-41700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~47~0041704~0041750~~~0041704~0041750~47~0~0~~0~~47~41704"
str2 = str2 & "-41750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~3~0041798~0041800~~~0041798~0041800~3~0~0~~0~~3~41798-41800~~~0~0~BC01-2L~01GTKT"
str2 = str2 & "-2LN-01~XQ/2008T~50~0041801~0041850~~~0041801~0041850~50~0~0~~0~~50~41801-41850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/20"
str2 = str2 & ""
str2 = str2 & "08T~650~0041851~0042500~~~0041851~0042500~650~0~0~~0~~650~41851-42500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~1300"
str2 = str2 & "0~0044501~0057500~~~0044501~0057500~13000~0~0~~0~~13000~44501-57500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~23~0038678~0038700~"
str2 = str2 & "~~0038678~0038700~23~0~0~~0~~23~38678-38700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~100~0038701~0038800~~~0038701~0038800~100~0"
str2 = str2 & "~0~~0~~100~38701-38800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~2~0038549~0038550~~~0038549~0038550~2~0~0~~0~~2~38549-38550~~~0~"
str2 = str2 & "0~BC01-2L~01GTKT-2LN-01~XQ/2008T~3~0038898~0038900~~~0038898~0038900~3~0~0~~0~~3~38898-38900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/"
str2 = str2 & "2008T~12~0039039~0039050~~~0039039~0039050~12~0~0~~0~~12~39039-39050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~17~0039084~0039100"
str2 = str2 & "~~~0039084~0039100~17~0~0~~0~~17~39084-39100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~400~0039101~0039500~~~0039101~0039500~400~"
str2 = str2 & "0~0~~0~~400~39101-39500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~1000~0042501~0043500~~~0042501~0043500~1000~0~0~~0~~1000~42501-"
str2 = str2 & "43500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~27~0035074~0035100~~~0035074~0035100~27~0~0~~0~~27~35074-35100~~~0~0~BC01-2L~01GT"
str2 = str2 & "KT-2LN-01~XQ/2008T~300~0035101~0035400~~~0035101~0035400~300~0~0~~0~~300~35101-35400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~28"
str2 = str2 & "~0035423~0035450~~~0035423~0035450~28~0~0~~0~~28~35423-35450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~50~0035451~0035500~~~00354"
str2 = str2 & "51~0035500~50~0~0~~0~~50~35451-35500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~9~0092642~0092650~~~0092642~0092650~9~0~0~~0~~9~92"
str2 = str2 & "642-92650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~4~0116547~0116550~~~0116547~0116550~4~0~0~~0~~4~116547-116550~~~0~0~BC01-2L~0"
str2 = str2 & "1GTKT-2LN-01~XQ/2008T~31~0116570~0116600~~~0116570~0116600~31~0~0~~0~~31~116570-116600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~"
str2 = str2 & "36~0115965~0116000~~~0115965~0116000~36~0~0~~0~~36~115965-116000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~26~0147725~0147750~~~0"
str2 = str2 & "147725~0147750~26~0~0~~0~~26~147725-147750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~23~0105678~0105700~~~0105678~0105"
str2 = str2 & ""
str2 = str2 & "700~23~0~0~~0~~23~105678-105700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~14~0105887~0105900~~~0105887~0105900~14~0~"
str2 = str2 & "0~~0~~14~105887-105900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~37~0106264~0106300~~~0106264~0106300~37~0~0~~0~~37~106264-106300"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~29~0106222~0106250~~~0106222~0106250~29~0~0~~0~~29~106222-106250~~~0~0~BC01-2L~01GTKT-"
str2 = str2 & "2LN-01~XQ/2008T~27~0106324~0106350~~~0106324~0106350~27~0~0~~0~~27~106324-106350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~49~010"
str2 = str2 & "6152~0106200~~~0106152~0106200~49~0~0~~0~~49~106152-106200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~31~0106370~0106400~~~0106370"
str2 = str2 & "~0106400~31~0~0~~0~~31~106370-106400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~15~0097486~0097500~~~0097486~0097500~15~0~0~~0~~15"
str2 = str2 & "~97486-97500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~32~0144119~0144150~~~0144119~0144150~32~0~0~~0~~32~144119-144150~~~0~0~BC0"
str2 = str2 & "1-2L~01GTKT-2LN-01~XQ/2008T~42~0096009~0096050~~~0096009~0096050~42~0~0~~0~~42~96009-96050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/20"
str2 = str2 & "08T~20~0096131~0096150~~~0096131~0096150~20~0~0~~0~~20~96131-96150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~14~0098337~0098350~~"
str2 = str2 & "~0098337~0098350~14~0~0~~0~~14~98337-98350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~10~0098441~0098450~~~0098441~0098450~10~0~0~"
str2 = str2 & "~0~~10~98441-98450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~2~0097849~0097850~~~0097849~0097850~2~0~0~~0~~2~97849-97850~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XQ/2008T~18~0091933~0091950~~~0091933~0091950~18~0~0~~0~~18~91933-91950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2"
str2 = str2 & "008T~20~0108681~0108700~~~0108681~0108700~20~0~0~~0~~20~108681-108700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~43~0107908~010795"
str2 = str2 & "0~~~0107908~0107950~43~0~0~~0~~43~107908-107950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~43~0107858~0107900~~~0107858~0107900~43"
str2 = str2 & "~0~0~~0~~43~107858-107900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~14~0103537~0103550~~~0103537~0103550~14~0~0~~0~~14~103537-103"
str2 = str2 & "550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~35~0103766~0103800~~~0103766~0103800~35~0~0~~0~~35~103766-103800~~~0~0~B"
str2 = str2 & ""
str2 = str2 & "C01-2L~01GTKT-2LN-01~XQ/2008T~43~0103808~0103850~~~0103808~0103850~43~0~0~~0~~43~103808-103850~~~0~0~BC01-2L~01"
str2 = str2 & "GTKT-2LN-01~XQ/2008T~21~0103930~0103950~~~0103930~0103950~21~0~0~~0~~21~103930-103950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~3"
str2 = str2 & "9~0104062~0104100~~~0104062~0104100~39~0~0~~0~~39~104062-104100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~18~0104133~0104150~~~01"
str2 = str2 & "04133~0104150~18~0~0~~0~~18~104133-104150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~23~0104328~0104350~~~0104328~0104350~23~0~0~~"
str2 = str2 & "0~~23~104328-104350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~14~0101137~0101150~~~0101137~0101150~14~0~0~~0~~14~101137-101150~~~"
str2 = str2 & "0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~33~0100668~0100700~~~0100668~0100700~33~0~0~~0~~33~100668-100700~~~0~0~BC01-2L~01GTKT-2LN"
str2 = str2 & "-01~XQ/2008T~39~0100562~0100600~~~0100562~0100600~39~0~0~~0~~39~100562-100600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~12~010118"
str2 = str2 & "9~0101200~~~0101189~0101200~12~0~0~~0~~12~101189-101200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~2~0101649~0101650~~~0101649~010"
str2 = str2 & "1650~2~0~0~~0~~2~101649-101650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~25~0152326~0152350~~~0152326~0152350~25~0~0~~0~~25~15232"
str2 = str2 & "6-152350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~8~0101493~0101500~~~0101493~0101500~8~0~0~~0~~8~101493-101500~~~0~0~BC01-2L~01"
str2 = str2 & "GTKT-2LN-01~XQ/2008T~32~0101269~0101300~~~0101269~0101300~32~0~0~~0~~32~101269-101300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~2"
str2 = str2 & "2~0101229~0101250~~~0101229~0101250~22~0~0~~0~~22~101229-101250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~4~0113547~0113550~~~011"
str2 = str2 & "3547~0113550~4~0~0~~0~~4~113547-113550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2008T~2~0113799~0113800~~~0113799~0113800~2~0~0~~0~~2~"
str2 = str2 & "113799-113800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~38~0049713~0049750~~~0049713~0049750~38~0~0~~0~~38~049713-049750~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XQ/2009T~28~0031473~0031500~~~0031473~0031500~28~0~0~~0~~28~031473-031500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ"
str2 = str2 & "/2009T~42~0007709~0007750~~~0007709~0007750~42~0~0~~0~~42~007709-007750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~2~00"
str2 = str2 & ""
str2 = str2 & "17649~0017650~~~0017649~0017650~2~0~0~~0~~2~017649-017650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~2~0044799~004480"
str2 = str2 & "0~~~0044799~0044800~2~0~0~~0~~2~044799-044800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~12~0062989~0063000~~~0062989~0063000~12~0"
str2 = str2 & "~0~~0~~12~062989-063000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~19~0059032~0059050~~~0059032~0059050~19~0~0~~0~~19~059032-05905"
str2 = str2 & "0~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~26~0061325~0061350~~~0061325~0061350~26~0~0~~0~~26~061325-061350~~~0~0~BC01-2L~01GTKT"
str2 = str2 & "-2LN-01~XQ/2009T~17~0061384~0061400~~~0061384~0061400~17~0~0~~0~~17~061384-061400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~50~00"
str2 = str2 & "61401~0061450~~~0061401~0061450~50~0~0~~0~~50~061401-061450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~7~0019994~0020000~~~0019994"
str2 = str2 & "~0020000~7~0~0~~0~~7~019994-020000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~28~0050523~0050550~~~0050523~0050550~28~0~0~~0~~28~0"
str2 = str2 & "50523-050550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~18~0050583~0050600~~~0050583~0050600~18~0~0~~0~~18~050583-050600~~~0~0~BC0"
str2 = str2 & "1-2L~01GTKT-2LN-01~XQ/2009T~34~0012217~0012250~~~0012217~0012250~34~0~0~~0~~34~012217-012250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/"
str2 = str2 & "2009T~28~0050273~0050300~~~0050273~0050300~28~0~0~~0~~28~050273-050300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~10~0050691~00507"
str2 = str2 & "00~~~0050691~0050700~10~0~0~~0~~10~050691-050700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~18~0010233~0010250~~~0010233~0010250~1"
str2 = str2 & "8~0~0~~0~~18~010233-010250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~26~0040925~0040950~~~0040925~0040950~26~0~0~~0~~26~040925-04"
str2 = str2 & "0950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~9~0047992~0048000~~~0047992~0048000~9~0~0~~0~~9~047992-048000~~~0~0~BC01-2L~01GTKT"
str2 = str2 & "-2LN-01~XQ/2009T~9~0043742~0043750~~~0043742~0043750~9~0~0~~0~~9~043742-043750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~2~002884"
str2 = str2 & "9~0028850~~~0028849~0028850~2~0~0~~0~~2~028849-028850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~21~0055880~0055900~~~0055880~0055"
str2 = str2 & "900~21~0~0~~0~~21~055880-055900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~27~0055124~0055150~~~0055124~0055150~27~0~0~"
str2 = str2 & ""
str2 = str2 & "~0~~27~055124-055150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~16~0055335~0055350~~~0055335~0055350~16~0~0~~0~~16~05"
str2 = str2 & "5335-055350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~38~0055663~0055700~~~0055663~0055700~38~0~0~~0~~38~055663-055700~~~0~0~BC01"
str2 = str2 & "-2L~01GTKT-2LN-01~XQ/2009T~14~0055937~0055950~~~0055937~0055950~14~0~0~~0~~14~055937-055950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2"
str2 = str2 & "009T~10~0064091~0064100~~~0064091~0064100~10~0~0~~0~~10~064091-064100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~35~0046966~004700"
str2 = str2 & "0~~~0046966~0047000~35~0~0~~0~~35~046966-047000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~14~0047087~0047100~~~0047087~0047100~14"
str2 = str2 & "~0~0~~0~~14~047087-047100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~29~0011872~0011900~~~0011872~0011900~29~0~0~~0~~29~011872-011"
str2 = str2 & "900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~10~0033091~0033100~~~0033091~0033100~10~0~0~~0~~10~033091-033100~~~0~0~BC01-2L~01GT"
str2 = str2 & "KT-2LN-01~XQ/2009T~22~0063979~0064000~~~0063979~0064000~22~0~0~~0~~22~063979-064000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~6~0"
str2 = str2 & "033895~0033900~~~0033895~0033900~6~0~0~~0~~6~033895-033900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~23~0063528~0063550~~~0063528"
str2 = str2 & "~0063550~23~0~0~~0~~23~063528-063550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~42~0064459~0064500~~~0064459~0064500~42~0~0~~0~~42"
str2 = str2 & "~064459-064500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~2~0064349~0064350~~~0064349~0064350~2~0~0~~0~~2~064349-064350~~~0~0~BC01"
str2 = str2 & "-2L~01GTKT-2LN-01~XQ/2009T~7~0054544~0054550~~~0054544~0054550~7~0~0~~0~~7~054544-054550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009"
str2 = str2 & "T~19~0054882~0054900~~~0054882~0054900~19~0~0~~0~~19~054882-054900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~18~0054033~0054050~~"
str2 = str2 & "~0054033~0054050~18~0~0~~0~~18~054033-054050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~10~0054241~0054250~~~0054241~0054250~10~0~"
str2 = str2 & "0~~0~~10~054241-054250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~13~0057788~0057800~~~0057788~0057800~13~0~0~~0~~13~057788-057800"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~10~0060091~0060100~~~0060091~0060100~10~0~0~~0~~10~060091-060100~~~0~0~BC01"
str2 = str2 & ""
str2 = str2 & "-2L~01GTKT-2LN-01~XQ/2009T~31~0060470~0060500~~~0060470~0060500~31~0~0~~0~~31~060470-060500~~~0~0~BC01-2L~01GTK"
str2 = str2 & "T-2LN-01~XQ/2009T~12~0021289~0021300~~~0021289~0021300~12~0~0~~0~~12~021289-021300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~11~0"
str2 = str2 & "036840~0036850~~~0036840~0036850~11~0~0~~0~~11~036840-036850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~24~0025577~0025600~~~00255"
str2 = str2 & "77~0025600~24~0~0~~0~~24~025577-025600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~15~0039086~0039100~~~0039086~0039100~15~0~0~~0~~"
str2 = str2 & "15~039086-039100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~9~0039842~0039850~~~0039842~0039850~9~0~0~~0~~9~039842-039850~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XQ/2009T~13~0025238~0025250~~~0025238~0025250~13~0~0~~0~~13~025238-025250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ"
str2 = str2 & "/2009T~5~0039246~0039250~~~0039246~0039250~5~0~0~~0~~5~039246-039250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~3~0056898~0056900~"
str2 = str2 & "~~0056898~0056900~3~0~0~~0~~3~056898-056900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~17~0053084~0053100~~~0053084~0053100~17~0~0"
str2 = str2 & "~~0~~17~053084-053100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~34~0056667~0056700~~~0056667~0056700~34~0~0~~0~~34~056667-056700~"
str2 = str2 & "~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~25~0032226~0032250~~~0032226~0032250~25~0~0~~0~~25~032226-032250~~~0~0~BC01-2L~01GTKT-2"
str2 = str2 & "LN-01~XQ/2009T~25~0032276~0032300~~~0032276~0032300~25~0~0~~0~~25~032276-032300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2009T~23~0032"
str2 = str2 & "528~0032550~~~0032528~0032550~23~0~0~~0~~23~032528-032550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~31~0018970~0019000~~~0018970~"
str2 = str2 & "0019000~31~0~0~~0~~31~18970-19000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~9~0006642~0006650~~~0006642~0006650~9~0~0~~0~~9~6642-"
str2 = str2 & "6650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~38~0006813~0006850~~~0006813~0006850~38~0~0~~0~~38~6813-6850~~~0~0~BC01-2L~01GTKT-"
str2 = str2 & "2LN-01~XQ/2010T~1~0006950~0006950~~~0006950~0006950~1~0~0~~0~~1~6950-6950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~20~0022331~00"
str2 = str2 & "22350~~~0022331~0022350~20~0~0~~0~~20~22331-22350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~22~0027729~0027750~~~00277"
str2 = str2 & ""
str2 = str2 & "29~0027750~22~0~0~~0~~22~27729-27750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~16~0027785~0027800~~~0027785~0027800~"
str2 = str2 & "16~0~0~~0~~16~27785-27800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~11~0027890~0027900~~~0027890~0027900~11~0~0~~0~~11~27890-2790"
str2 = str2 & "0~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~6~0027995~0028000~~~0027995~0028000~6~0~0~~0~~6~27995-28000~~~0~0~BC01-2L~01GTKT-2LN-"
str2 = str2 & "01~XQ/2010T~41~0024960~0025000~~~0024960~0025000~41~0~0~~0~~41~24960-25000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~49~0025002~0"
str2 = str2 & "025050~~~0025002~0025050~49~0~0~~0~~49~25002-25050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~28~0025073~0025100~~~0025073~0025100"
str2 = str2 & "~28~0~0~~0~~28~25073-25100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~92~0025159~0025250~~~0025159~0025250~92~0~0~~0~~92~25159-252"
str2 = str2 & "50~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~32~0013719~0013750~~~0013719~0013750~32~0~0~~0~~32~13719-13750~~~0~0~BC01-2L~01GTKT-"
str2 = str2 & "2LN-01~XQ/2010T~28~0013773~0013800~~~0013773~0013800~28~0~0~~0~~28~13773-13800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~26~00144"
str2 = str2 & "75~0014500~~~0014475~0014500~26~0~0~~0~~26~14475-14500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~30~0021871~0021900~~~0021871~002"
str2 = str2 & "1900~30~0~0~~0~~30~21871-21900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~20~0022181~0022200~~~0022181~0022200~20~0~0~~0~~20~22181"
str2 = str2 & "-22200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~21~0023530~0023550~~~0023530~0023550~21~0~0~~0~~21~23530-23550~~~0~0~BC01-2L~01G"
str2 = str2 & "TKT-2LN-01~XQ/2010T~15~0023936~0023950~~~0023936~0023950~15~0~0~~0~~15~23936-23950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~39~0"
str2 = str2 & "023962~0024000~~~0023962~0024000~39~0~0~~0~~39~23962-24000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~42~0024009~0024050~~~0024009"
str2 = str2 & "~0024050~42~0~0~~0~~42~24009-24050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~38~0024163~0024200~~~0024163~0024200~38~0~0~~0~~38~2"
str2 = str2 & "4163-24200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~44~0024207~0024250~~~0024207~0024250~44~0~0~~0~~44~24207-24250~~~0~0~BC01-2L"
str2 = str2 & "~01GTKT-2LN-01~XQ/2010T~31~0028270~0028300~~~0028270~0028300~31~0~0~~0~~31~28270-28300~~~0~0~BC01-2L~01GTKT-2LN-0"
str2 = str2 & ""
str2 = str2 & "1~XQ/2010T~30~0014821~0014850~~~0014821~0014850~30~0~0~~0~~30~14821-14850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~"
str2 = str2 & "50~0014851~0014900~~~0014851~0014900~50~0~0~~0~~50~14851-14900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~49~0028552~0028600~~~002"
str2 = str2 & "8552~0028600~49~0~0~~0~~49~28552-28600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0028601~0028650~~~0028601~0028650~50~0~0~~0~~"
str2 = str2 & "50~28601-28650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0028651~0028700~~~0028651~0028700~50~0~0~~0~~50~28651-28700~~~0~0~BC0"
str2 = str2 & "1-2L~01GTKT-2LN-01~XQ/2010T~50~0028701~0028750~~~0028701~0028750~50~0~0~~0~~50~28701-28750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/20"
str2 = str2 & "10T~50~0004801~0004850~~~0004801~0004850~50~0~0~~0~~50~4801-4850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0004851~0004900~~~0"
str2 = str2 & "004851~0004900~50~0~0~~0~~50~4851-4900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0004901~0004950~~~0004901~0004950~50~0~0~~0~~"
str2 = str2 & "50~4901-4950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~26~0014925~0014950~~~0014925~0014950~26~0~0~~0~~26~14925-14950~~~0~0~BC01-"
str2 = str2 & "2L~01GTKT-2LN-01~XQ/2010T~44~0014957~0015000~~~0014957~0015000~44~0~0~~0~~44~14957-15000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010"
str2 = str2 & "T~40~0010861~0010900~~~0010861~0010900~40~0~0~~0~~40~10861-10900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~30~0020421~0020450~~~0"
str2 = str2 & "020421~0020450~30~0~0~~0~~30~20421-20450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~43~0010758~0010800~~~0010758~0010800~43~0~0~~0"
str2 = str2 & "~~43~10758-10800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~43~0010658~0010700~~~0010658~0010700~43~0~0~~0~~43~10658-10700~~~0~0~B"
str2 = str2 & "C01-2L~01GTKT-2LN-01~XQ/2010T~50~0010701~0010750~~~0010701~0010750~50~0~0~~0~~50~10701-10750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/"
str2 = str2 & "2010T~20~0024481~0024500~~~0024481~0024500~20~0~0~~0~~20~24481-24500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~21~0010330~0010350"
str2 = str2 & "~~~0010330~0010350~21~0~0~~0~~21~10330-10350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~27~0019674~0019700~~~0019674~0019700~27~0~"
str2 = str2 & "0~~0~~27~19674-19700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~48~0024703~0024750~~~0024703~0024750~48~0~0~~0~~48~2470"
str2 = str2 & ""
str2 = str2 & "3-24750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~22~0001379~0001400~~~0001379~0001400~22~0~0~~0~~22~1379-1400~~~0~0"
str2 = str2 & "~BC01-2L~01GTKT-2LN-01~XQ/2010T~4~0001847~0001850~~~0001847~0001850~4~0~0~~0~~4~1847-1850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/201"
str2 = str2 & "0T~38~0024663~0024700~~~0024663~0024700~38~0~0~~0~~38~24663-24700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~40~0010361~0010400~~~"
str2 = str2 & "0010361~0010400~40~0~0~~0~~40~10361-10400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~43~0001908~0001950~~~0001908~0001950~43~0~0~~"
str2 = str2 & "0~~43~1908-1950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~35~0029366~0029400~~~0029366~0029400~35~0~0~~0~~35~29366-29400~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XQ/2010T~37~0029414~0029450~~~0029414~0029450~37~0~0~~0~~37~29414-29450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2"
str2 = str2 & "010T~14~0029487~0029500~~~0029487~0029500~14~0~0~~0~~14~29487-29500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~5~0029546~0029550~~"
str2 = str2 & "~0029546~0029550~5~0~0~~0~~5~29546-29550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~39~0029562~0029600~~~0029562~0029600~39~0~0~~0"
str2 = str2 & "~~39~29562-29600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0029601~0029650~~~0029601~0029650~50~0~0~~0~~50~29601-29650~~~0~0~B"
str2 = str2 & "C01-2L~01GTKT-2LN-01~XQ/2010T~19~0029682~0029700~~~0029682~0029700~19~0~0~~0~~19~29682-29700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/"
str2 = str2 & "2010T~16~0029735~0029750~~~0029735~0029750~16~0~0~~0~~16~29735-29750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~39~0017462~0017500"
str2 = str2 & "~~~0017462~0017500~39~0~0~~0~~39~17462-17500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~18~0021033~0021050~~~0021033~0021050~18~0~"
str2 = str2 & "0~~0~~18~21033-21050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~46~0021405~0021450~~~0021405~0021450~46~0~0~~0~~46~21405-21450~~~0"
str2 = str2 & "~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~18~0017083~0017100~~~0017083~0017100~18~0~0~~0~~18~17083-17100~~~0~0~BC01-2L~01GTKT-2LN-01"
str2 = str2 & "~XQ/2010T~3~0017248~0017250~~~0017248~0017250~3~0~0~~0~~3~17248-17250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~23~0021278~002130"
str2 = str2 & "0~~~0021278~0021300~23~0~0~~0~~23~21278-21300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~35~0021616~0021650~~~0021616~0"
str2 = str2 & ""
str2 = str2 & "021650~35~0~0~~0~~35~21616-21650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~10~0017291~0017300~~~0017291~0017300~10~0"
str2 = str2 & "~0~~0~~10~17291-17300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~43~0021558~0021600~~~0021558~0021600~43~0~0~~0~~43~21558-21600~~~"
str2 = str2 & "0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~23~0017328~0017350~~~0017328~0017350~23~0~0~~0~~23~17328-17350~~~0~0~BC01-2L~01GTKT-2LN-0"
str2 = str2 & "1~XQ/2010T~38~0021113~0021150~~~0021113~0021150~38~0~0~~0~~38~21113-21150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~43~0021058~00"
str2 = str2 & "21100~~~0021058~0021100~43~0~0~~0~~43~21058-21100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~36~0021515~0021550~~~0021515~0021550~"
str2 = str2 & "36~0~0~~0~~36~21515-21550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~28~0021173~0021200~~~0021173~0021200~28~0~0~~0~~28~21173-2120"
str2 = str2 & "0~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~47~0021704~0021750~~~0021704~0021750~47~0~0~~0~~47~21704-21750~~~0~0~BC01-2L~01GTKT-2"
str2 = str2 & "LN-01~XQ/2010T~3~0018398~0018400~~~0018398~0018400~3~0~0~~0~~3~18398-18400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~18~0030283~0"
str2 = str2 & "030300~~~0030283~0030300~18~0~0~~0~~18~30283-30300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~32~0030419~0030450~~~0030419~0030450"
str2 = str2 & "~32~0~0~~0~~32~30419-30450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~5~0030496~0030500~~~0030496~0030500~5~0~0~~0~~5~30496-30500~"
str2 = str2 & "~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~1~0030550~0030550~~~0030550~0030550~1~0~0~~0~~1~30550-30550~~~0~0~BC01-2L~01GTKT-2LN-01"
str2 = str2 & "~XQ/2010T~21~0030680~0030700~~~0030680~0030700~21~0~0~~0~~21~30680-30700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~35~0030716~003"
str2 = str2 & "0750~~~0030716~0030750~35~0~0~~0~~35~30716-30750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~47~0031104~0031150~~~0031104~0031150~4"
str2 = str2 & "7~0~0~~0~~47~31104-31150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~10~0031241~0031250~~~0031241~0031250~10~0~0~~0~~10~31241-31250"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~1~0031350~0031350~~~0031350~0031350~1~0~0~~0~~1~031350-031350~~~0~0~BC01-2L~01GTKT-2LN"
str2 = str2 & "-01~XQ/2010T~48~0031353~0031400~~~0031353~0031400~48~0~0~~0~~48~31353-31400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~"
str2 = str2 & ""
str2 = str2 & "15~0031436~0031450~~~0031436~0031450~15~0~0~~0~~15~31436-31450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~16~0031485~"
str2 = str2 & "0031500~~~0031485~0031500~16~0~0~~0~~16~31485-31500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~250~0031501~0031750~~~0031501~00317"
str2 = str2 & "50~250~0~0~~0~~250~31501-31750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~19~0032082~0032100~~~0032082~0032100~19~0~0~~0~~19~32082"
str2 = str2 & "-32100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~9~0032192~0032200~~~0032192~0032200~9~0~0~~0~~9~32192-32200~~~0~0~BC01-2L~01GTKT"
str2 = str2 & "-2LN-01~XQ/2010T~7~0032244~0032250~~~0032244~0032250~7~0~0~~0~~7~32244-32250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~13~0032388"
str2 = str2 & "~0032400~~~0032388~0032400~13~0~0~~0~~13~32388-32400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~42~0032409~0032450~~~0032409~00324"
str2 = str2 & "50~42~0~0~~0~~42~32409-32450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~32~0032469~0032500~~~0032469~0032500~32~0~0~~0~~32~32469-3"
str2 = str2 & "2500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~16~0032535~0032550~~~0032535~0032550~16~0~0~~0~~16~32535-32550~~~0~0~BC01-2L~01GTK"
str2 = str2 & "T-2LN-01~XQ/2010T~45~0032656~0032700~~~0032656~0032700~45~0~0~~0~~45~32656-32700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~35~003"
str2 = str2 & "2716~0032750~~~0032716~0032750~35~0~0~~0~~35~32716-32750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~19~0032782~0032800~~~0032782~0"
str2 = str2 & "032800~19~0~0~~0~~19~32782-32800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~10~0032841~0032850~~~0032841~0032850~10~0~0~~0~~10~328"
str2 = str2 & "41-32850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~8~0032993~0033000~~~0032993~0033000~8~0~0~~0~~8~32993-33000~~~0~0~BC01-2L~01GT"
str2 = str2 & "KT-2LN-01~XQ/2010T~18~0033133~0033150~~~0033133~0033150~18~0~0~~0~~18~33133-33150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~5~003"
str2 = str2 & "3246~0033250~~~0033246~0033250~5~0~0~~0~~5~33246-33250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~41~0033260~0033300~~~0033260~003"
str2 = str2 & "3300~41~0~0~~0~~41~33260-33300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~25~0033326~0033350~~~0033326~0033350~25~0~0~~0~~25~33326"
str2 = str2 & "-33350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~6~0033395~0033400~~~0033395~0033400~6~0~0~~0~~6~33395-33400~~~0~0~BC0"
str2 = str2 & ""
str2 = str2 & "1-2L~01GTKT-2LN-01~XQ/2010T~15~0033486~0033500~~~0033486~0033500~15~0~0~~0~~15~33486-33500~~~0~0~BC01-2L~01GTKT"
str2 = str2 & "-2LN-01~XQ/2010T~12~0033889~0033900~~~0033889~0033900~12~0~0~~0~~12~33889-33900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~5~00339"
str2 = str2 & "46~0033950~~~0033946~0033950~5~0~0~~0~~5~33946-33950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~49~0033952~0034000~~~0033952~00340"
str2 = str2 & "00~49~0~0~~0~~49~33952-34000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~40~0034011~0034050~~~0034011~0034050~40~0~0~~0~~40~34011-3"
str2 = str2 & "4050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~28~0034123~0034150~~~0034123~0034150~28~0~0~~0~~28~34123-34150~~~0~0~BC01-2L~01GTK"
str2 = str2 & "T-2LN-01~XQ/2010T~46~0034155~0034200~~~0034155~0034200~46~0~0~~0~~46~34155-34200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~18~003"
str2 = str2 & "4233~0034250~~~0034233~0034250~18~0~0~~0~~18~34233-34250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~250~0034251~0034500~~~0034251~"
str2 = str2 & "0034500~250~0~0~~0~~250~34251-34500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~32~0034569~0034600~~~0034569~0034600~32~0~0~~0~~32~"
str2 = str2 & "34569-34600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~14~0034637~0034650~~~0034637~0034650~14~0~0~~0~~14~34637-34650~~~0~0~BC01-2"
str2 = str2 & "L~01GTKT-2LN-01~XQ/2010T~44~0034657~0034700~~~0034657~0034700~44~0~0~~0~~44~34657-34700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T"
str2 = str2 & "~35~0034716~0034750~~~0034716~0034750~35~0~0~~0~~35~34716-34750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~17~0034834~0034850~~~00"
str2 = str2 & "34834~0034850~17~0~0~~0~~17~34834-34850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~3~0034898~0034900~~~0034898~0034900~3~0~0~~0~~3"
str2 = str2 & "~34898-34900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~31~0034920~0034950~~~0034920~0034950~31~0~0~~0~~31~34920-34950~~~0~0~BC01-"
str2 = str2 & "2L~01GTKT-2LN-01~XQ/2010T~46~0034955~0035000~~~0034955~0035000~46~0~0~~0~~46~34955-35000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010"
str2 = str2 & "T~25~0027476~0027500~~~0027476~0027500~25~0~0~~0~~25~27476-27500~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~46~0027305~0027350~~~0"
str2 = str2 & "027305~0027350~46~0~0~~0~~46~27305-27350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~41~0017860~0017900~~~0017860~001790"
str2 = str2 & ""
str2 = str2 & "0~41~0~0~~0~~41~17860-17900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0017901~0017950~~~0017901~0017950~50~0~0~~0"
str2 = str2 & "~~50~17901-17950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~26~0018325~0018350~~~0018325~0018350~26~0~0~~0~~26~18325-18350~~~0~0~B"
str2 = str2 & "C01-2L~01GTKT-2LN-01~XQ/2010T~44~0020907~0020950~~~0020907~0020950~44~0~0~~0~~44~20907-20950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/"
str2 = str2 & "2010T~44~0020507~0020550~~~0020507~0020550~44~0~0~~0~~44~20507-20550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~43~0017508~0017550"
str2 = str2 & "~~~0017508~0017550~43~0~0~~0~~43~17508-17550~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~39~0018012~0018050~~~0018012~0018050~39~0~"
str2 = str2 & "0~~0~~39~18012-18050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~49~0017552~0017600~~~0017552~0017600~49~0~0~~0~~49~17552-17600~~~0"
str2 = str2 & "~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~32~0017719~0017750~~~0017719~0017750~32~0~0~~0~~32~17719-17750~~~0~0~BC01-2L~01GTKT-2LN-01"
str2 = str2 & "~XQ/2010T~44~0020757~0020800~~~0020757~0020800~44~0~0~~0~~44~20757-20800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~34~0017817~001"
str2 = str2 & "7850~~~0017817~0017850~34~0~0~~0~~34~17817-17850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~39~0017612~0017650~~~0017612~0017650~3"
str2 = str2 & "9~0~0~~0~~39~17612-17650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0017651~0017700~~~0017651~0017700~50~0~0~~0~~50~17651-17700"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~34~0017767~0017800~~~0017767~0017800~34~0~0~~0~~34~17767-17800~~~0~0~BC01-2L~01GTKT-2L"
str2 = str2 & "N-01~XQ/2010T~49~0018052~0018100~~~0018052~0018100~49~0~0~~0~~49~18052-18100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~41~0017960"
str2 = str2 & "~0018000~~~0017960~0018000~41~0~0~~0~~41~17960-18000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~42~0020609~0020650~~~0020609~00206"
str2 = str2 & "50~42~0~0~~0~~42~20609-20650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~44~0020557~0020600~~~0020557~0020600~44~0~0~~0~~44~20557-2"
str2 = str2 & "0600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~15~0002786~0002800~~~0002786~0002800~15~0~0~~0~~15~2786-2800~~~0~0~BC01-2L~01GTKT-"
str2 = str2 & "2LN-01~XQ/2010T~1~0002950~0002950~~~0002950~0002950~1~0~0~~0~~1~002950-002950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010"
str2 = str2 & ""
str2 = str2 & "T~16~0003285~0003300~~~0003285~0003300~16~0~0~~0~~16~3285-3300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~38~0003313~"
str2 = str2 & "0003350~~~0003313~0003350~38~0~0~~0~~38~3313-3350~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~1~0003450~0003450~~~0003450~0003450~1"
str2 = str2 & "~0~0~~0~~1~003450-003450~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~42~0003459~0003500~~~0003459~0003500~42~0~0~~0~~42~3459-3500~~"
str2 = str2 & "~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~49~0003602~0003650~~~0003602~0003650~49~0~0~~0~~49~3602-3650~~~0~0~BC01-2L~01GTKT-2LN-01"
str2 = str2 & "~XQ/2010T~11~0003690~0003700~~~0003690~0003700~11~0~0~~0~~11~3690-3700~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~23~0003728~00037"
str2 = str2 & "50~~~0003728~0003750~23~0~0~~0~~23~3728-3750~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0003751~0003800~~~0003751~0003800~50~0~"
str2 = str2 & "0~~0~~50~3751-3800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~48~0003803~0003850~~~0003803~0003850~48~0~0~~0~~48~3803-3850~~~0~0~B"
str2 = str2 & "C01-2L~01GTKT-2LN-01~XQ/2010T~49~0003852~0003900~~~0003852~0003900~49~0~0~~0~~49~3852-3900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/20"
str2 = str2 & "10T~50~0003901~0003950~~~0003901~0003950~50~0~0~~0~~50~3901-3950~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0003951~0004000~~~0"
str2 = str2 & "003951~0004000~50~0~0~~0~~50~3951-4000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~41~0002260~0002300~~~0002260~0002300~41~0~0~~0~~"
str2 = str2 & "41~2260-2300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~30~0003021~0003050~~~0003021~0003050~30~0~0~~0~~30~3021-3050~~~0~0~BC01-2L"
str2 = str2 & "~01GTKT-2LN-01~XQ/2010T~22~0029029~0029050~~~0029029~0029050~22~0~0~~0~~22~29029-29050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~"
str2 = str2 & "28~0029223~0029250~~~0029223~0029250~28~0~0~~0~~28~29223-29250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~50~0000551~0000600~~~000"
str2 = str2 & "0551~0000600~50~0~0~~0~~50~551-600~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~15~0000636~0000650~~~0000636~0000650~15~0~0~~0~~15~6"
str2 = str2 & "36-650~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~40~0001061~0001100~~~0001061~0001100~40~0~0~~0~~40~1061-1100~~~0~0~BC01-2L~01GTK"
str2 = str2 & "T-2LN-01~XQ/2010T~39~0012362~0012400~~~0012362~0012400~39~0~0~~0~~39~12362-12400~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2"
str2 = str2 & ""
str2 = str2 & "010T~22~0025879~0025900~~~0025879~0025900~22~0~0~~0~~22~25879-25900~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~12~002"
str2 = str2 & "6039~0026050~~~0026039~0026050~12~0~0~~0~~12~26039-26050~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~34~0026067~0026100~~~0026067~0"
str2 = str2 & "026100~34~0~0~~0~~34~26067-26100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~5~0029796~0029800~~~0029796~0029800~5~0~0~~0~~5~29796-"
str2 = str2 & "29800~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~40~0029961~0030000~~~0029961~0030000~40~0~0~~0~~40~29961-30000~~~0~0~BC01-2L~01GT"
str2 = str2 & "KT-2LN-01~XQ/2010T~26~0008275~0008300~~~0008275~0008300~26~0~0~~0~~26~8275-8300~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~44~0008"
str2 = str2 & "807~0008850~~~0008807~0008850~44~0~0~~0~~44~8807-8850~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~9~0008992~0009000~~~0008992~00090"
str2 = str2 & "00~9~0~0~~0~~9~8992-9000~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~48~0013253~0013300~~~0013253~0013300~48~0~0~~0~~48~13253-13300"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~2~0012649~0012650~~~0012649~0012650~2~0~0~~0~~2~12649-12650~~~0~0~BC01-2L~01GTKT-2LN-0"
str2 = str2 & "1~XQ/2010T~18~0027183~0027200~~~0027183~0027200~18~0~0~~0~~18~27183-27200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~24~0027077~00"
str2 = str2 & "27100~~~0027077~0027100~24~0~0~~0~~24~27077-27100~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~18~0011133~0011150~~~0011133~0011150~"
str2 = str2 & "18~0~0~~0~~18~11133-11150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~3~0011398~0011400~~~0011398~0011400~3~0~0~~0~~3~11398-11400~~"
str2 = str2 & "~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~16~0011435~0011450~~~0011435~0011450~16~0~0~~0~~16~11435-11450~~~0~0~BC01-2L~01GTKT-2LN-"
str2 = str2 & "01~XQ/2010T~38~0019113~0019150~~~0019113~0019150~38~0~0~~0~~38~19113-19150~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~7~0019194~00"
str2 = str2 & "19200~~~0019194~0019200~7~0~0~~0~~7~19194-19200~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~45~0019206~0019250~~~0019206~0019250~45"
str2 = str2 & "~0~0~~0~~45~19206-19250~~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~45~0019256~0019300~~~0019256~0019300~45~0~0~~0~~45~19256-19300~"
str2 = str2 & "~~0~0~BC01-2L~01GTKT-2LN-01~XQ/2010T~34~0019317~0019350~~~0019317~0019350~34~0~0~~0~~34~19317-19350~~~0~0~BC01-2L"
str2 = str2 & ""
str2 = str2 & "~01GTKT-2LN-01~XQ/2010T~13~0023188~0023200~~~0023188~0023200~13~0~0~~0~~13~23188-23200~~~0~0~BC01-2L~01GTKT-2LN"
str2 = str2 & "-01~XQ/2010T~8~0028143~0028150~~~0028143~0028150~8~0~0~~0~~8~28143-28150~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~14~0012337~001"
str2 = str2 & "2350~~~0012337~0012350~14~0~0~~0~~14~12337-12350~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~46~0012355~0012400~~~0012355~0012400~4"
str2 = str2 & "6~0~0~~0~~46~12355-12400~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~24~0012427~0012450~~~0012427~0012450~24~0~0~~0~~24~12427-12450"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~24~0012477~0012500~~~0012477~0012500~24~0~0~~0~~24~12477-12500~~~0~0~BC01-2L~01GTKT-2L"
str2 = str2 & "N-01~XR/2009T~1500~0012501~0014000~~~0012501~0014000~1500~0~0~~0~~1500~12501-14000~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~19~0"
str2 = str2 & "014132~0014150~~~0014132~0014150~19~0~0~~0~~19~014132-014150~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~28~0014173~0014200~~~00141"
str2 = str2 & "73~0014200~28~0~0~~0~~28~014173-014200~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~31~0014270~0014300~~~0014270~0014300~31~0~0~~0~~"
str2 = str2 & "31~14270-14300~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~25~0014326~0014350~~~0014326~0014350~25~0~0~~0~~25~14326-14350~~~0~0~BC0"
str2 = str2 & "1-2L~01GTKT-2LN-01~XR/2009T~46~0014355~0014400~~~0014355~0014400~46~0~0~~0~~46~14355-14400~~~0~0~BC01-2L~01GTKT-2LN-01~XR/20"
str2 = str2 & "09T~38~0014413~0014450~~~0014413~0014450~38~0~0~~0~~38~14413-14450~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~48~0014453~0014500~~"
str2 = str2 & "~0014453~0014500~48~0~0~~0~~48~14453-14500~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~37~0014514~0014550~~~0014514~0014550~37~0~0~"
str2 = str2 & "~0~~37~14514-14550~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~1450~0014551~0016000~~~0014551~0016000~1450~0~0~~0~~1450~14551-16000"
str2 = str2 & "~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~12~0016089~0016100~~~0016089~0016100~12~0~0~~0~~12~16089-16100~~~0~0~BC01-2L~01GTKT-2L"
str2 = str2 & "N-01~XR/2009T~9~0016142~0016150~~~0016142~0016150~9~0~0~~0~~9~16142-16150~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~24~0016177~00"
str2 = str2 & "16200~~~0016177~0016200~24~0~0~~0~~24~16177-16200~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~30~0016221~0016250~~~00162"
str2 = str2 & ""
str2 = str2 & "21~0016250~30~0~0~~0~~30~16221-16250~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~48~0016253~0016300~~~0016253~0016300~"
str2 = str2 & "48~0~0~~0~~48~16253-16300~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~42~0016309~0016350~~~0016309~0016350~42~0~0~~0~~42~16309-1635"
str2 = str2 & "0~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~650~0016351~0017000~~~0016351~0017000~650~0~0~~0~~650~16351-17000~~~0~0~BC01-2L~01GTK"
str2 = str2 & "T-2LN-01~XR/2009T~44~0017007~0017050~~~0017007~0017050~44~0~0~~0~~44~17007-17050~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~45~001"
str2 = str2 & "7056~0017100~~~0017056~0017100~45~0~0~~0~~45~17056-17100~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~44~0017107~0017150~~~0017107~0"
str2 = str2 & "017150~44~0~0~~0~~44~17107-17150~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~6~0017195~0017200~~~0017195~0017200~6~0~0~~0~~6~17195-"
str2 = str2 & "17200~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~38~0017213~0017250~~~0017213~0017250~38~0~0~~0~~38~17213-17250~~~0~0~BC01-2L~01GT"
str2 = str2 & "KT-2LN-01~XR/2009T~50~0017251~0017300~~~0017251~0017300~50~0~0~~0~~50~17251-17300~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~350~0"
str2 = str2 & "017301~0017650~~~0017301~0017650~350~0~0~~0~~350~17301-17650~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~350~0017651~0018000~~~0017"
str2 = str2 & "651~0018000~350~0~0~~0~~350~17651-18000~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~22~0018029~0018050~~~0018029~0018050~22~0~0~~0~"
str2 = str2 & "~22~18029-18050~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~17~0018084~0018100~~~0018084~0018100~17~0~0~~0~~17~18084-18100~~~0~0~BC"
str2 = str2 & "01-2L~01GTKT-2LN-01~XR/2009T~26~0018125~0018150~~~0018125~0018150~26~0~0~~0~~26~18125-18150~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2"
str2 = str2 & "009T~44~0018157~0018200~~~0018157~0018200~44~0~0~~0~~44~18157-18200~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~43~0018208~0018250~"
str2 = str2 & "~~0018208~0018250~43~0~0~~0~~43~18208-18250~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~1750~0018251~0020000~~~0018251~0020000~1750"
str2 = str2 & "~0~0~~0~~1750~18251-20000~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~257~0001744~0002000~~~0001744~0002000~257~0~0~~0~~257~1744-20"
str2 = str2 & "00~~~0~0~BC01-2L~01GTKT-2LN-01~XR/2009T~848~0003153~0004000~~~0003153~0004000~848~0~0~~0~~848~3153-4000~~~0~0~BC0"
str2 = str2 & ""
str2 = str2 & "1-2L~01GTKT-2LN-01~XR/2009T~144~0011857~0012000~~~0011857~0012000~144~0~0~~0~~144~11857-12000~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000001~0000050~~~0000001~0000050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0000051~0000100~~~0000051~0000100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01G"
str2 = str2 & "TKT2/001~AA/11P~50~0000101~0000150~~~0000101~0000150~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/1"
str2 = str2 & "1P~50~0000151~0000200~~~0000151~0000200~50~49~1~0000200~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0000201~0000250~~~0000201~0000216~16~16~0~~0~~0~~0000217~0000250~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~5"
str2 = str2 & "0~0000251~0000300~~~0000251~0000300~50~49~1~0000259~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000"
str2 = str2 & "301~0000350~~~0000301~0000350~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000351~0000400~~"
str2 = str2 & "~0000351~0000396~46~46~0~~0~~0~~0000397~0000400~4~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000401~0000450~"
str2 = str2 & "~~0000401~0000450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000451~0000500~~~0000451~000"
str2 = str2 & "0500~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000501~0000550~~~0000501~0000515~15~15~0~"
str2 = str2 & "~0~~0~~0000516~0000550~35~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000551~0000600~~~0000551~0000579~29~29~"
str2 = str2 & "0~~0~~0~~0000580~0000600~21~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000601~0000650~~~0000601~0000611~11~1"
str2 = str2 & "1~0~~0~~0~~0000612~0000650~39~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000651~0000700~~~0000651~0000651~1~"
str2 = str2 & "1~0~~0~~0~~0000652~0000700~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000701~0000750~~~0000701~0000750~50"
str2 = str2 & "~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000751~0000800~~~0000751~0000800~50~5"
str2 = str2 & ""
str2 = str2 & "0~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000801~0000850~~~0000801~0000834~34~3"
str2 = str2 & "4~0~~0~~0~~0000835~0000850~16~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000851~0000900~~~0000851~0000900~50"
str2 = str2 & "~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000901~0000950~~~0000901~0000946~46~44~2~905;915"
str2 = str2 & "~0~~0~~0000947~0000950~4~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0000951~0001000~~~0000951~0000992~42~42~0"
str2 = str2 & "~~0~~0~~0000993~0001000~8~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001001~0001050~~~0001001~0001050~50~50~"
str2 = str2 & "0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001051~0001100~~~~~0~0~0~~0~~0~~0001051~0001100~50~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001101~0001150~~~0001101~0001150~50~49~1~1111~0~~0~~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001151~0001200~~~0001151~0001200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng - BC012L~01GTKT2/001~AA/11P~50~0001201~0001250~~~0001201~0001250~50~49~1~1222~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC0"
str2 = str2 & "12L~01GTKT2/001~AA/11P~50~0001251~0001300~~~0001251~0001300~50~48~2~1265;1293~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~"
str2 = str2 & "01GTKT2/001~AA/11P~50~0001301~0001350~~~0001301~0001350~50~49~1~1344~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/0"
str2 = str2 & "01~AA/11P~50~0001351~0001400~~~0001351~0001400~50~49~1~1370~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P"
str2 = str2 & "~50~0001401~0001450~~~0001401~0001450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001451~0"
str2 = str2 & "001500~~~0001451~0001500~50~48~2~1484;1485~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001501~00015"
str2 = str2 & "50~~~0001501~0001550~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001551~0001600~~~0001551~"
str2 = str2 & "0001600~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001601~0001650~~~0001601~00"
str2 = str2 & ""
str2 = str2 & "01650~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001651~0001700~~~0001651~00"
str2 = str2 & "01700~50~49~1~1699~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001701~0001750~~~0001701~0001750~50~"
str2 = str2 & "50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001751~0001800~~~0001751~0001800~50~50~0~~0~~0~~~"
str2 = str2 & "~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001801~0001850~~~0001801~0001850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0001851~0001900~~~0001851~0001900~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng - BC012L~01GTKT2/001~AA/11P~50~0001901~0001950~~~0001901~0001950~50~49~1~1944~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC0"
str2 = str2 & "12L~01GTKT2/001~AA/11P~50~0001951~0002000~~~0001951~0001963~13~13~0~~0~~0~~0001964~0002000~37~0~H„a Æ¨n gi∏ trﬁ gia t®ng - B"
str2 = str2 & "C012L~01GTKT2/001~AA/11P~50~0002001~0002050~~~0002001~0002050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2"
str2 = str2 & "/001~AA/11P~50~0002051~0002100~~~0002051~0002088~38~38~0~~0~~0~~0002089~0002100~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTK"
str2 = str2 & "T2/001~AA/11P~50~0002101~0002150~~~0002101~0002123~23~23~0~~0~~0~~0002124~0002150~27~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01G"
str2 = str2 & "TKT2/001~AA/11P~50~0002151~0002200~~~0002151~0002200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/1"
str2 = str2 & "1P~50~0002201~0002250~~~0002201~0002250~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0002251"
str2 = str2 & "~0002300~~~0002251~0002286~36~36~0~~0~~0~~0002287~0002300~14~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00023"
str2 = str2 & "01~0002350~~~0002301~0002304~4~4~0~~0~~0~~0002305~0002350~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00023"
str2 = str2 & "51~0002400~~~0002351~0002353~3~3~0~~0~~0~~0002354~0002400~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00024"
str2 = str2 & "01~0002450~~~0002401~0002430~30~29~1~0002401~0~~0~~0002431~0002450~20~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2"
str2 = str2 & ""
str2 = str2 & "/001~AA/11P~50~0002451~0002500~~~0002451~0002490~40~39~1~0002451~0~~0~~0002491~0002500~10~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0002501~0002550~~~0002501~0002507~7~7~0~~0~~0~~0002508~0002550~43~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0002551~0002600~~~0002551~0002553~3~3~0~~0~~0~~0002554~0002600~47~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0002601~0002650~~~0002601~0002605~5~5~0~~0~~0~~0002606~0002650~45~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0002651~0002700~~~0002651~0002659~9~8~1~0002657~0~~0~~0002660~0002700~41~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0002701~0002750~~~0002701~0002719~19~19~0~~0~~0~~0002720~0002750~31~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0002751~0002800~~~0002751~0002762~12~12~0~~0~~0~~0002763~0002800~38~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0002801~0002850~~~0002801~0002822~22~22~0~~0~~0~~0002823~0002850~28~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0002851~0002900~~~0002851~0002859~9~9~0~~0~~0~~0002860~0002900~41~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0002901~0002950~~~~~0~0~0~~0~~0~~0002901~0002950~50~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0002951~0003000~~~0002951~0002968~18~17~1~0002952~0~~0~~0002969~0003000~32~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0003001~0003050~~~0003001~0003002~2~2~0~~0~~0~~0003003~0003050~48~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0003051~0003100~~~0003051~0003056~6~6~0~~0~~0~~0003057~0003100~44~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0003101~0003150~~~~~0~0~0~~0~~0~~0003101~0003150~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0003151~0003200~~~0003151~0003200~50~49~1~0003193~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC01"
str2 = str2 & "2L~01GTKT2/001~AA/11P~50~0003201~0003250~~~0003201~0003244~44~44~0~~0~~0~~0003245~0003250~6~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & ""
str2 = str2 & " t®ng - BC012L~01GTKT2/001~AA/11P~50~0003251~0003300~~~0003251~0003300~50~48~2~0003263;0003273~0~~0~~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0003301~0003350~~~~~0~0~0~~0~~0~~0003301~0003350~50~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng - BC012L~01GTKT2/001~AA/11P~50~0003351~0003400~~~0003351~0003368~18~18~0~~0~~0~~0003369~0003400~32~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0003401~0003450~~~~~0~0~0~~0~~0~~0003401~0003450~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC0"
str2 = str2 & "12L~01GTKT2/001~AA/11P~50~0003451~0003500~~~0003451~0003468~18~17~1~0003466~0~~0~~0003469~0003500~32~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng - BC012L~01GTKT2/001~AA/11P~300~0003501~0003800~~~~~0~0~0~~0~~0~~0003501~0003800~300~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L"
str2 = str2 & "~01GTKT2/001~AA/11P~50~0003801~0003850~~~0003801~0003850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~0003851~0003900~~~0003851~0003900~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~000"
str2 = str2 & "3901~0003950~~~0003901~0003950~50~47~3~3916;3928;3945~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00"
str2 = str2 & "03951~0004000~~~0003951~0004000~50~49~1~0003971~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004001~"
str2 = str2 & "0004050~~~0004001~0004050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004051~0004100~~~000"
str2 = str2 & "4051~0004100~50~49~1~0004082~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004101~0004150~~~0004101~0"
str2 = str2 & "004150~50~49~1~0004130~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004151~0004200~~~0004151~0004183"
str2 = str2 & "~33~32~1~0004174~0~~0~~0004184~0004200~17~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004201~0004250~~~000420"
str2 = str2 & "1~0004203~3~3~0~~0~~0~~0004204~0004250~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004251~0004300~~~000425"
str2 = str2 & "1~0004255~5~5~0~~0~~0~~0004256~0004300~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004301~00043"
str2 = str2 & ""
str2 = str2 & "50~~~~~0~0~0~~0~~0~~0004301~0004350~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004351~000440"
str2 = str2 & "0~~~0004351~0004352~2~2~0~~0~~0~~0004353~0004400~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004401~000445"
str2 = str2 & "0~~~0004401~0004402~2~2~0~~0~~0~~0004403~0004450~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004451~000450"
str2 = str2 & "0~~~0004451~0004451~1~1~0~~0~~0~~0004452~0004500~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004501~000455"
str2 = str2 & "0~~~0004501~0004505~5~5~0~~0~~0~~0004506~0004550~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004551~000460"
str2 = str2 & "0~~~0004551~0004554~4~4~0~~0~~0~~0004555~0004600~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004601~000465"
str2 = str2 & "0~~~0004601~0004603~3~3~0~~0~~0~~0004604~0004650~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004651~000470"
str2 = str2 & "0~~~0004651~0004659~9~9~0~~0~~0~~0004660~0004700~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004701~000475"
str2 = str2 & "0~~~0004701~0004701~1~1~0~~0~~0~~0004702~0004750~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004751~000480"
str2 = str2 & "0~~~0004751~0004768~18~18~0~~0~~0~~0004769~0004800~32~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004801~0004"
str2 = str2 & "850~~~0004801~0004837~37~37~0~~0~~0~~0004838~0004850~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004851~00"
str2 = str2 & "04900~~~0004851~0004859~9~9~0~~0~~0~~0004860~0004900~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004901~00"
str2 = str2 & "04950~~~0004901~0004910~10~10~0~~0~~0~~0004911~0004950~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0004951~"
str2 = str2 & "0005000~~~~~0~0~0~~0~~0~~0004951~0005000~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0005001~0005050~~~0005"
str2 = str2 & "001~0005050~50~49~1~0005018~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0005051~0005100~~~0005051~00"
str2 = str2 & "05065~15~15~0~~0~~0~~0005066~0005100~35~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0005101~0005150"
str2 = str2 & ""
str2 = str2 & "~~~0005101~0005113~13~13~0~~0~~0~~0005114~0005150~37~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0005151~0005200~~~0005151~0005164~14~14~0~~0~~0~~0005165~0005200~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~5"
str2 = str2 & "0~0005201~0005250~~~0005201~0005202~2~2~0~~0~~0~~0005203~0005250~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~5"
str2 = str2 & "0~0005251~0005300~~~0005251~0005256~6~6~0~~0~~0~~0005257~0005300~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~5"
str2 = str2 & "0~0005301~0005350~~~0005301~0005320~20~20~0~~0~~0~~0005321~0005350~30~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P"
str2 = str2 & "~50~0005351~0005400~~~0005351~0005372~22~22~0~~0~~0~~0005373~0005400~28~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/1"
str2 = str2 & "1P~50~0005401~0005450~~~0005401~0005420~20~20~0~~0~~0~~0005421~0005450~30~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA"
str2 = str2 & "/11P~50~0005451~0005500~~~0005451~0005459~9~9~0~~0~~0~~0005460~0005500~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA"
str2 = str2 & "/11P~50~0005501~0005550~~~0005501~0005514~14~14~0~~0~~0~~0005515~0005550~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~0005551~0005600~~~0005551~0005557~7~7~0~~0~~0~~0005558~0005600~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~0005601~0005650~~~0005601~0005608~8~8~0~~0~~0~~0005609~0005650~42~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~0005651~0005700~~~0005651~0005652~2~2~0~~0~~0~~0005653~0005700~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~0005701~0005750~~~0005701~0005702~2~2~0~~0~~0~~0005703~0005750~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~0005751~0005800~~~~~0~0~0~~0~~0~~0005751~0005800~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0005"
str2 = str2 & "801~0005850~~~0005801~0005802~2~2~0~~0~~0~~0005803~0005850~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0005"
str2 = str2 & "851~0005900~~~0005851~0005864~14~14~0~~0~~0~~0005865~0005900~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~A"
str2 = str2 & ""
str2 = str2 & "A/11P~50~0005901~0005950~~~0005901~0005912~12~12~0~~0~~0~~0005913~0005950~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012"
str2 = str2 & "L~01GTKT2/001~AA/11P~50~0005951~0006000~~~0005951~0006000~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001"
str2 = str2 & "~AA/11P~50~0006001~0006050~~~0006001~0006050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00"
str2 = str2 & "06051~0006100~~~0006051~0006100~50~46~4~0006054;6062;6089;6090~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/"
str2 = str2 & "11P~50~0006101~0006150~~~0006101~0006150~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~000615"
str2 = str2 & "1~0006200~~~0006151~0006153~3~3~0~~0~~0~~0006154~0006200~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~000620"
str2 = str2 & "1~0006250~~~0006201~0006207~7~7~0~~0~~0~~0006208~0006250~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~000625"
str2 = str2 & "1~0006300~~~0006251~0006300~50~49~1~0006258~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006301~0006"
str2 = str2 & "350~~~0006301~0006324~24~24~0~~0~~0~~0006325~0006350~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006351~00"
str2 = str2 & "06400~~~0006351~0006351~1~1~0~~0~~0~~0006352~0006400~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006401~00"
str2 = str2 & "06450~~~0006401~0006409~9~9~0~~0~~0~~0006410~0006450~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006451~00"
str2 = str2 & "06500~~~0006451~0006453~3~3~0~~0~~0~~0006454~0006500~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006501~00"
str2 = str2 & "06550~~~0006501~0006550~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006551~0006600~~~00065"
str2 = str2 & "51~0006571~21~21~0~~0~~0~~0006572~0006600~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006601~0006650~~~000"
str2 = str2 & "6601~0006650~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006651~0006700~~~0006651~0006672~"
str2 = str2 & "22~22~0~~0~~0~~0006673~0006700~28~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006701~0006750~~~000"
str2 = str2 & ""
str2 = str2 & "6701~0006750~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006751~0006800~~~000"
str2 = str2 & "6751~0006771~21~21~0~~0~~0~~0006772~0006800~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006801~0006850~~~0"
str2 = str2 & "006801~0006816~16~16~0~~0~~0~~0006817~0006850~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006851~0006900~~"
str2 = str2 & "~0006851~0006900~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006901~0006950~~~0006901~0006"
str2 = str2 & "950~50~49~1~0006942~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0006951~0007000~~~0006951~0007000~50"
str2 = str2 & "~49~1~0006999~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007001~0007050~~~0007001~0007012~12~12~0~"
str2 = str2 & "~0~~0~~0007013~0007050~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007051~0007100~~~0007051~0007054~4~4~0~"
str2 = str2 & "~0~~0~~0007055~0007100~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007101~0007150~~~0007101~0007106~6~6~0~"
str2 = str2 & "~0~~0~~0007107~0007150~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007151~0007200~~~0007151~0007154~4~4~0~"
str2 = str2 & "~0~~0~~0007155~0007200~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007201~0007250~~~0007201~0007214~14~14~"
str2 = str2 & "0~~0~~0~~0007215~0007250~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007251~0007300~~~0007251~0007300~50~5"
str2 = str2 & "0~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007301~0007350~~~0007301~0007337~37~37~0~~0~~0~~00"
str2 = str2 & "07338~0007350~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007351~0007400~~~0007351~0007356~6~6~0~~0~~0~~00"
str2 = str2 & "07357~0007400~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007401~0007450~~~0007401~0007404~4~4~0~~0~~0~~00"
str2 = str2 & "07405~0007450~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007451~0007500~~~0007451~0007456~6~6~0~~0~~0~~00"
str2 = str2 & "07457~0007500~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007501~0007550~~~0007501~0007550~50~4"
str2 = str2 & ""
str2 = str2 & "5~5~7509;7518;7527;7535;7543~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007551~000760"
str2 = str2 & "0~~~0007551~0007600~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007601~0007650~~~0007601~0"
str2 = str2 & "007650~50~49~1~0007635~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007651~0007700~~~0007651~0007700"
str2 = str2 & "~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007701~0007750~~~0007701~0007735~35~35~0~~0~~"
str2 = str2 & "0~~0007736~0007750~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007751~0007800~~~0007751~0007766~16~16~0~~0"
str2 = str2 & "~~0~~0007767~0007800~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007801~0007850~~~0007801~0007850~50~50~0~"
str2 = str2 & "~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007851~0007900~~~0007851~0007900~50~49~1~0007853~0~~0~"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007901~0007950~~~0007901~0007940~40~40~0~~0~~0~~0007941~0007"
str2 = str2 & "950~10~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0007951~0008000~~~0007951~0007964~14~13~1~0007951~0~~0~~000"
str2 = str2 & "7965~0008000~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008001~0008050~~~~~0~0~0~~0~~0~~0008001~0008050~5"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008051~0008100~~~0008051~0008082~32~32~0~~0~~0~~0008083~0008100"
str2 = str2 & "~18~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008101~0008150~~~0008101~0008150~50~50~0~~0~~0~~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008151~0008200~~~0008151~0008160~10~9~1~0008158~0~~0~~0008161~0008200~40~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008201~0008250~~~0008201~0008206~6~6~0~~0~~0~~0008207~0008250~44~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008251~0008300~~~0008251~0008256~6~6~0~~0~~0~~0008257~0008300~44~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008301~0008350~~~0008301~0008305~5~5~0~~0~~0~~0008306~0"
str2 = str2 & ""
str2 = str2 & "008350~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008351~0008400~~~0008351~0008358~8~7~1~000"
str2 = str2 & "8356~0~~0~~0008359~0008400~42~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012~01GTKT2/001~AA/11P~50~0008401~0008450~~~0008401~0008405~5~5"
str2 = str2 & "~0~~0~~0~~0008406~0008450~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008451~0008500~~~0008451~0008500~50~"
str2 = str2 & "50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008501~0008550~~~0008501~0008521~21~21~0~~0~~0~~0"
str2 = str2 & "008522~0008550~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008551~0008600~~~0008551~0008600~50~49~1~000856"
str2 = str2 & "7~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008601~0008650~~~0008601~0008650~50~50~0~~0~~0~~~~0~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008651~0008700~~~0008651~0008700~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0008701~0008750~~~0008701~0008750~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0008751~0008800~~~0008751~0008800~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GT"
str2 = str2 & "KT2/001~AA/11P~50~0008801~0008850~~~0008801~0008805~5~5~0~~0~~0~~0008806~0008850~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GT"
str2 = str2 & "KT2/001~AA/11P~50~0008851~0008900~~~0008851~0008900~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11"
str2 = str2 & "P~50~0008901~0008950~~~0008901~0008930~30~29~1~0008922~0~~0~~0008931~0008950~20~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/"
str2 = str2 & "001~AA/11P~50~0008951~0009000~~~0008951~0008998~48~47~1~0008961~0~~0~~0008999~0009000~2~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~"
str2 = str2 & "01GTKT2/001~AA/11P~50~0009001~0009050~~~0009001~0009050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~A"
str2 = str2 & "A/11P~50~0009051~0009100~~~0009051~0009078~28~27~1~0009056~0~~0~~0009079~0009100~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GT"
str2 = str2 & "KT2/001~AA/11P~50~0009101~0009150~~~0009101~0009118~18~18~0~~0~~0~~0009119~0009150~32~0~H„a Æ¨n gi∏ trﬁ gia t®ng "
str2 = str2 & ""
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0009151~0009200~~~0009151~0009200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0009201~0009250~~~0009201~0009226~26~26~0~~0~~0~~0009227~0009250~24~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g - BC012L~01GTKT2/001~AA/11P~50~0009251~0009300~~~0009251~0009268~18~18~0~~0~~0~~0009269~0009300~32~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng - BC012L~01GTKT2/001~AA/11P~50~0009301~0009350~~~0009301~0009339~39~39~0~~0~~0~~0009340~0009350~11~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0009351~0009400~~~0009351~0009365~15~15~0~~0~~0~~0009366~0009400~35~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009401~0009450~~~0009401~0009439~39~37~2~009416;9431~0~~0~~0009440~0009450~11~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009451~0009500~~~0009451~0009456~6~6~0~~0~~0~~0009457~0009500~44~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009501~0009550~~~0009501~0009541~41~41~0~~0~~0~~0009542~0009550~9~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009551~0009600~~~0009551~0009596~46~46~0~~0~~0~~0009597~0009600~4~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009601~0009650~~~0009601~0009611~11~11~0~~0~~0~~0009612~0009650~39~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009651~0009700~~~~~0~0~0~~0~~0~~0009651~0009700~50~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009701~0009750~~~0009701~0009729~29~29~0~~0~~0~~0009730~0009750~21~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009751~0009800~~~0009751~0009759~9~9~0~~0~~0~~0009760~0009800~41~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009801~0009850~~~0009801~0009804~4~4~0~~0~~0~~0009805~0009850~46~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009851~0009900~~~0009851~0009900~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0009901~0009950~~~0009901~0009905~5~5~0~~0~~0~~0009906~0009950~45~0~H„a Æ¨n gi∏ tr"
str2 = str2 & ""
str2 = str2 & "ﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0009951~0010000~~~~~0~0~0~~0~~0~~0009951~0010000~50~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0010001~0010050~~~0010001~0010009~9~9~0~~0~~0~~0010010~0010050~41~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0010051~0010100~~~0010051~0010100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - B"
str2 = str2 & "C012L~01GTKT2/001~AA/11P~50~0010101~0010150~~~0010101~0010150~50~49~1~0010128~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~"
str2 = str2 & "01GTKT2/001~AA/11P~100~0010151~0010250~~~0010151~0010250~100~100~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/00"
str2 = str2 & "1~AA/11P~50~0010251~0010300~~~0010251~0010300~50~49~1~0010271~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/1"
str2 = str2 & "1P~50~0010301~0010350~~~0010301~0010350~50~49~1~0010301~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010351~0010400~~~0010351~0010352~2~2~0~~0~~0~~0010353~0010400~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010401~0010450~~~0010401~0010407~7~7~0~~0~~0~~0010408~0010450~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010451~0010500~~~0010451~0010456~6~6~0~~0~~0~~0010457~0010500~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010501~0010550~~~0010501~0010507~7~7~0~~0~~0~~0010508~0010550~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010551~0010600~~~0010551~0010557~7~7~0~~0~~0~~0010558~0010600~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010601~0010650~~~0010601~0010609~9~9~0~~0~~0~~0010610~0010650~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0010651~0010700~~~0010651~0010679~29~26~3~10662;10667;10668~0~~0~~0010680~0010700~21~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01G"
str2 = str2 & "TKT2/001~AA/11P~50~0010701~0010750~~~0010701~0010710~10~9~1~0010707~0~~0~~0010711~0010750~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC"
str2 = str2 & "012L~01GTKT2/001~AA/11P~50~0010751~0010800~~~0010751~0010754~4~3~1~0010752~0~~0~~0010755~0010800~46~0~H„a Æ¨n gi∏"
str2 = str2 & ""
str2 = str2 & " trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0010801~0010850~~~0010801~0010809~9~8~1~0010807~0~~0~~0010810~0010"
str2 = str2 & "850~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0010851~0010900~~~0010851~0010851~1~1~0~~0~~0~~0010852~0010"
str2 = str2 & "900~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0010901~0010950~~~0010901~0010909~9~9~0~~0~~0~~0010910~0010"
str2 = str2 & "950~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0010951~0011000~~~0010951~0010962~12~12~0~~0~~0~~0010963~00"
str2 = str2 & "11000~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0011001~0011050~~~0011001~0011013~13~13~0~~0~~0~~0011014~"
str2 = str2 & "0011050~37~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~250~0011051~0011300~~~0011051~0011300~250~245~5~11066;1107"
str2 = str2 & "6;"
str2 = str2 & "11077"
str2 = str2 & ";11090;11109~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0011301~0011350~~~0011301~0011305~"
str2 = str2 & "5~4~1~11303~0~~0~~0011306~0011350~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~100~0011351~0011450~~~0011351~00"
str2 = str2 & "11450~100~99~1~11400~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0011451~0011500~~~~~0~0~0~~0~~0~~00"
str2 = str2 & "11451~0011500~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~500~0011501~0012000~~~0011501~0012000~500~484~16~115"
str2 = str2 & "85;11609;11658;11665;11695;11728;11778;11779;11806;11807;11855;11885;11894;11913;11924;11937~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0012001~0012050~~~0012001~0012048~48~48~0~~0~~0~~0012049~0012050~2~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng - BC012L~01GTKT2/001~AA/11P~50~0012051~0012100~~~0012051~0012053~3~3~0~~0~~0~~0012054~0012100~47~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng - BC012L~01GTKT2/001~AA/11P~300~0012101~0012400~~~0012101~0012400~300~294~6~12132;12210;12223;12248;12269;12325~0~~0~"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012401~0012450~~~0012401~0012420~20~18~2~12410;12416~0~~0~~0"
str2 = str2 & "012421~0012450~30~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012451~0012500~~~0012451~0012500~50~"
str2 = str2 & ""
str2 = str2 & "50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012501~0012550~~~0012501~0012531~31~"
str2 = str2 & "31~0~~0~~0~~0012532~0012550~19~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012551~0012600~~~0012551~0012600~5"
str2 = str2 & "0~48~2~0012563;0012572~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012601~0012650~~~0012601~0012650"
str2 = str2 & "~50~48~2~0012608;0012615~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012651~0012700~~~0012651~00126"
str2 = str2 & "99~49~49~0~~0~~0~~0012700~0012700~1~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012701~0012750~~~0012701~0012"
str2 = str2 & "728~28~27~1~0012702~0~~0~~0012729~0012750~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012751~0012800~~~001"
str2 = str2 & "2751~0012756~6~6~0~~0~~0~~0012757~0012800~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012801~0012850~~~001"
str2 = str2 & "2801~0012803~3~2~1~0012803~0~~0~~0012804~0012850~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0012851~001290"
str2 = str2 & "0~~~0012851~0012868~18~17~1~0012863~0~~0~~0012869~0012900~32~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00129"
str2 = str2 & "01~0012950~~~0012901~0012935~35~35~0~~0~~0~~0012936~0012950~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~001"
str2 = str2 & "2951~0013000~~~0012951~0012953~3~3~0~~0~~0~~0012954~0013000~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~001"
str2 = str2 & "3001~0013050~~~0013001~0013050~50~49~1~0013001~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013051~0"
str2 = str2 & "013100~~~0013051~0013100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013101~0013150~~~0013"
str2 = str2 & "101~0013150~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013151~0013200~~~0013151~0013200~5"
str2 = str2 & "0~45~5~13160;13166;"
str2 = str2 & "13175;13176;13196~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013201~0013250~~~"
str2 = str2 & "0013201~0013250~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013251~0013300~~~00"
str2 = str2 & ""
str2 = str2 & "13251~0013300~50~49~1~0013289~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013301~00133"
str2 = str2 & "50~~~0013301~0013350~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013351~0013400~~~0013351~"
str2 = str2 & "0013380~30~30~0~~0~~0~~0013381~0013400~20~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013401~0013450~~~001340"
str2 = str2 & "1~0013437~37~36~1~0013427~0~~0~~0013438~0013450~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013451~0013500"
str2 = str2 & "~~~0013451~0013487~37~36~1~0013467~0~~0~~0013488~0013500~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~001350"
str2 = str2 & "1~0013550~~~0013501~0013550~50~49~1~0013517~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013551~0013"
str2 = str2 & "600~~~0013551~0013576~26~26~0~~0~~0~~0013577~0013600~24~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013601~00"
str2 = str2 & "13650~~~0013601~0013608~8~8~0~~0~~0~~0013609~0013650~42~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013651~00"
str2 = str2 & "13700~~~0013651~0013651~1~1~0~~0~~0~~0013652~0013700~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013701~00"
str2 = str2 & "13750~~~0013701~0013738~38~38~0~~0~~0~~0013739~0013750~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0013751~"
str2 = str2 & "0013800~~~0013751~0013778~28~27~1~0013773~0~~0~~0013779~0013800~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50"
str2 = str2 & "~0013801~0013850~~~0013801~0013805~5~5~0~~0~~0~~0013806~0013850~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50"
str2 = str2 & "~0013851~0013900~~~0013851~0013890~40~40~0~~0~~0~~0013891~0013900~10~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~"
str2 = str2 & "50~0013901~0013950~~~0013901~0013901~1~1~0~~0~~0~~0013902~0013950~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~"
str2 = str2 & "50~0013951~0014000~~~0013951~0013974~24~24~0~~0~~0~~0013975~0014000~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11"
str2 = str2 & "P~50~0014001~0014050~~~0014001~0014021~21~21~0~~0~~0~~0014022~0014050~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GT"
str2 = str2 & ""
str2 = str2 & "KT2/001~AA/11P~50~0014051~0014100~~~0014051~0014052~2~2~0~~0~~0~~0014053~0014100~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0014101~0014150~~~0014101~0014107~7~6~1~0014104~0~~0~~0014108~0014150~43~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng - BC012L~01GTKT2/001~AA/11P~50~0014151~0014200~~~0014151~0014199~49~49~0~~0~~0~~0014200~0014200~1~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014201~0014250~~~0014201~0014217~17~17~0~~0~~0~~0014218~0014250~33~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014251~0014300~~~0014251~0014257~7~7~0~~0~~0~~0014258~0014300~43~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014301~0014350~~~0014301~0014315~15~14~1~0014309~0~~0~~0014316~0014350~35~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014351~0014400~~~0014351~0014400~50~49~1~0014380~0~~0~~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014401~0014450~~~0014401~0014450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "- BC012L~01GTKT2/001~AA/11P~50~0014451~0014500~~~0014451~0014500~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GT"
str2 = str2 & "KT2/001~AA/11P~50~0014501~0014550~~~0014501~0014550~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11"
str2 = str2 & "P~50~0014551~0014600~~~0014551~0014579~29~29~0~~0~~0~~0014580~0014600~21~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/"
str2 = str2 & "11P~50~0014601~0014650~~~0014601~0014650~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~001465"
str2 = str2 & "1~0014700~~~0014651~0014700~50~49~1~0014686~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014701~0014"
str2 = str2 & "750~~~0014701~0014750~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014751~0014800~~~0014751"
str2 = str2 & "~0014800~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014801~0014850~~~0014801~0014821~21~2"
str2 = str2 & "1~0~~0~~0~~0014822~0014850~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014851~0014900~~~0014851"
str2 = str2 & ""
str2 = str2 & "~0014889~39~39~0~~0~~0~~0014890~0014900~11~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014901~00"
str2 = str2 & "14950~~~0014901~0014906~6~6~0~~0~~0~~0014907~0014950~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0014951~00"
str2 = str2 & "15000~~~0014951~0014974~24~24~0~~0~~0~~0014975~0015000~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015001~"
str2 = str2 & "0015050~~~0015001~0015050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015051~0015100~~~001"
str2 = str2 & "5051~0015100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015101~0015150~~~0015101~0015103~"
str2 = str2 & "3~3~0~~0~~0~~0015104~0015150~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015151~0015200~~~0015151~0015152~"
str2 = str2 & "2~2~0~~0~~0~~0015153~0015200~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015201~0015250~~~~~0~0~0~~0~~0~~0"
str2 = str2 & "015201~0015250~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015251~0015300~~~0015251~0015289~39~35~4~15288;"
str2 = str2 & "15264;15272;15275~0~~0~~0015290~0015300~11~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015301~0015350~~~00153"
str2 = str2 & "01~0015303~3~3~0~~0~~0~~0015304~0015350~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015351~0015400~~~00153"
str2 = str2 & "51~0015367~17~17~0~~0~~0~~0015368~0015400~33~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015401~0015450~~~001"
str2 = str2 & "5401~0015419~19~19~0~~0~~0~~0015420~0015450~31~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015451~0015500~~~0"
str2 = str2 & "015451~0015500~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015501~0015550~~~0015501~001551"
str2 = str2 & "2~12~12~0~~0~~0~~0015513~0015550~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015551~0015600~~~0015551~0015"
str2 = str2 & "553~3~3~0~~0~~0~~0015554~0015600~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015601~0015650~~~0015601~0015"
str2 = str2 & "603~3~3~0~~0~~0~~0015604~0015650~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015651~0015700~~~0"
str2 = str2 & ""
str2 = str2 & "015651~0015667~17~17~0~~0~~0~~0015668~0015700~33~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015"
str2 = str2 & "701~0015750~~~0015701~0015703~3~3~0~~0~~0~~0015704~0015750~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015"
str2 = str2 & "751~0015800~~~0015751~0015754~4~4~0~~0~~0~~0015755~0015800~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015"
str2 = str2 & "801~0015850~~~0015801~0015806~6~6~0~~0~~0~~0015807~0015850~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015"
str2 = str2 & "851~0015900~~~0015851~0015856~6~6~0~~0~~0~~0015857~0015900~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015"
str2 = str2 & "901~0015950~~~0015901~0015906~6~6~0~~0~~0~~0015907~0015950~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0015"
str2 = str2 & "951~0016000~~~0015951~0015969~19~19~0~~0~~0~~0015970~0016000~31~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00"
str2 = str2 & "16001~0016050~~~0016001~0016006~6~6~0~~0~~0~~0016007~0016050~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00"
str2 = str2 & "16051~0016100~~~0016051~0016053~3~3~0~~0~~0~~0016054~0016100~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~00"
str2 = str2 & "16101~0016150~~~0016101~0016123~23~23~0~~0~~0~~0016124~0016150~27~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "0016151~0016200~~~0016151~0016166~16~16~0~~0~~0~~0016167~0016200~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~1"
str2 = str2 & "00~0016201~0016300~~~~~0~0~0~~0~~0~~0016201~0016300~100~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0016301~00"
str2 = str2 & "16350~~~0016301~0016310~10~10~0~~0~~0~~0016311~0016350~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0016351~"
str2 = str2 & "0016400~~~0016351~0016360~10~10~0~~0~~0~~0016361~0016400~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~001640"
str2 = str2 & "1~0016450~~~0016401~0016410~10~10~0~~0~~0~~0016411~0016450~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0016"
str2 = str2 & "451~0016500~~~0016451~0016457~7~7~0~~0~~0~~0016458~0016500~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/"
str2 = str2 & ""
str2 = str2 & "11P~50~0016501~0016550~~~0016501~0016502~2~2~0~~0~~0~~0016503~0016550~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01"
str2 = str2 & "GTKT2/001~AA/11P~50~0016551~0016600~~~0016551~0016600~50~49~1~0016596~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/"
str2 = str2 & "001~AA/11P~50~0016601~0016650~~~0016601~0016650~50~49~1~0016619~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA"
str2 = str2 & "/11P~50~0016651~0016700~~~0016651~0016699~49~48~1~0016698~0~~0~~0016700~0016700~1~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT"
str2 = str2 & "2/001~AA/11P~50~0016701~0016750~~~0016701~0016743~43~42~1~0016701~0~~0~~0016744~0016750~7~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012"
str2 = str2 & "L~01GTKT2/001~AA/11P~50~0016751~0016800~~~0016751~0016775~25~24~1~0016769~0~~0~~0016776~0016800~25~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g - BC012L~01GTKT2/001~AA/11P~50~0016801~0016850~~~0016801~0016802~2~2~0~~0~~0~~0016803~0016850~48~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g - BC012L~01GTKT2/001~AA/11P~50~0016851~0016900~~~0016851~0016859~9~9~0~~0~~0~~0016860~0016900~41~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g - BC012L~01GTKT2/001~AA/11P~50~0016901~0016950~~~0016901~0016950~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01"
str2 = str2 & "GTKT2/001~AA/11P~50~0016951~0017000~~~0016951~0017000~50~49~1~0016991~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/"
str2 = str2 & "001~AA/11P~50~0017001~0017050~~~0017001~0017005~5~4~1~0017002~0~~0~~0017006~0017050~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~0"
str2 = str2 & "1GTKT2/001~AA/11P~50~0017051~0017100~~~0017051~0017062~12~12~0~~0~~0~~0017063~0017100~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L"
str2 = str2 & "~01GTKT2/001~AA/11P~50~0017101~0017150~~~0017101~0017136~36~36~0~~0~~0~~0017137~0017150~14~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC01"
str2 = str2 & "2L~01GTKT2/001~AA/11P~50~0017151~0017200~~~0017151~0017181~31~31~0~~0~~0~~0017182~0017200~19~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC"
str2 = str2 & "012L~01GTKT2/001~AA/11P~50~0017201~0017250~~~0017201~0017226~26~26~0~~0~~0~~0017227~0017250~24~0~H„a Æ¨n gi∏ trﬁ gia t®ng -"
str2 = str2 & "BC012L~01GTKT2/001~AA/11P~50~0017251~0017300~~~0017251~0017258~8~8~0~~0~~0~~0017259~0017300~42~0~H„a Æ¨n gi∏ trﬁ "
str2 = str2 & ""
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017301~0017350~~~0017301~0017312~12~12~0~~0~~0~~0017313~0017350~38~0~H"
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017351~0017400~~~0017351~0017362~12~12~0~~0~~0~~0017363~0017400~38~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017401~0017450~~~0017401~0017403~3~3~0~~0~~0~~0017404~0017450~47~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017451~0017500~~~0017451~0017453~3~3~0~~0~~0~~0017454~0017500~47~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017501~0017550~~~0017501~0017508~8~8~0~~0~~0~~0017509~0017550~42~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017551~0017600~~~0017551~0017557~7~7~0~~0~~0~~0017558~0017600~43~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017601~0017650~~~0017601~0017603~3~3~0~~0~~0~~0017604~0017650~47~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017651~0017700~~~0017651~0017654~4~4~0~~0~~0~~0017655~0017700~46~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017701~0017750~~~0017701~0017703~3~3~0~~0~~0~~0017704~0017750~47~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017751~0017800~~~0017751~0017753~3~3~0~~0~~0~~0017754~0017800~47~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017801~0017850~~~0017801~0017850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0017851~0017900~~~0017851~0017900~50~49~1~17863~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0017901~0017950~~~0017901~0017907~7~7~0~~0~~0~~0017908~0017950~43~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0017951~0018000~~~0017951~0018000~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L"
str2 = str2 & "~01GTKT2/001~AA/11P~50~0018001~0018050~~~0018001~0018006~6~6~0~~0~~0~~0018007~0018050~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L"
str2 = str2 & "~01GTKT2/001~AA/11P~50~0018051~0018100~~~0018051~0018066~16~16~0~~0~~0~~0018067~0018100~34~0~H„a Æ¨n gi∏ trﬁ gia "
str2 = str2 & ""
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0018101~0018150~~~0018101~0018150~50~49~1~0018132~0~~0~~~~0~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018151~0018200~~~0018151~0018200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng -"
str2 = str2 & "BC012L~01GTKT2/001~AA/11P~50~0018201~0018250~~~0018201~0018250~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTK"
str2 = str2 & "T2/001~AA/11P~50~0018251~0018300~~~0018251~0018300~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P"
str2 = str2 & "~50~0018301~0018350~~~0018301~0018350~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018351~0"
str2 = str2 & "018400~~~0018351~0018400~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018401~0018450~~~0018"
str2 = str2 & "401~0018403~3~3~0~~0~~0~~0018404~0018450~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018451~0018500~~~0018"
str2 = str2 & "451~0018494~44~44~0~~0~~0~~0018495~0018500~6~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018501~0018550~~~001"
str2 = str2 & "8501~0018550~50~49~1~18528~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018551~0018600~~~0018551~001"
str2 = str2 & "8600~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018601~0018650~~~0018601~0018650~50~50~0~"
str2 = str2 & "~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018651~0018700~~~0018651~0018700~50~49~1~18658~0~~0~~~"
str2 = str2 & "~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018701~0018750~~~0018701~0018738~38~38~0~~0~~0~~0018739~001875"
str2 = str2 & "0~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018751~0018800~~~0018751~0018800~50~48~2~18766;18768~0~~0~~~"
str2 = str2 & "~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018801~0018850~~~0018801~0018850~50~48~2~18816;18840~0~~0~~~~0"
str2 = str2 & "~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018851~0018900~~~0018851~0018900~50~49~1~0018856~0~~0~~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018901~0018950~~~0018901~0018950~50~47~3~18909;18915;18929~"
str2 = str2 & ""
str2 = str2 & "0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0018951~0019000~~~0018951~0018985~35~35~0~~"
str2 = str2 & "0~~0~~0018986~0019000~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019001~0019050~~~0019001~0019050~50~48~2"
str2 = str2 & "~19017;19048~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019051~0019100~~~0019051~0019100~50~47~3~1"
str2 = str2 & "9051;19052;19054~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019101~0019150~~~0019101~0019150~50~49"
str2 = str2 & "~1~0019136~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019151~0019200~~~0019151~0019200~50~50~0~~0~"
str2 = str2 & "~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019201~0019250~~~0019201~0019250~50~49~1~0019228~0~~0~~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019251~0019300~~~0019251~0019300~50~50~0~~0~~0~~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019301~0019350~~~~~0~0~0~~0~~0~~0019301~0019350~50~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g - BC012L~01GTKT2/001~AA/11P~50~0019351~0019400~~~0019351~0019351~1~1~0~~0~~0~~0019352~0019400~49~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g - BC012L~01GTKT2/001~AA/11P~50~0019401~0019450~~~~~0~0~0~~0~~0~~0019401~0019450~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01G"
str2 = str2 & "TKT2/001~AA/11P~50~0019451~0019500~~~0019451~0019471~21~21~0~~0~~0~~0019472~0019500~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~0"
str2 = str2 & "1GTKT2/001~AA/11P~50~0019501~0019550~~~~~0~0~0~~0~~0~~0019501~0019550~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/"
str2 = str2 & "11P~50~0019551~0019600~~~0019551~0019588~38~38~0~~0~~0~~0019589~0019600~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~A"
str2 = str2 & "A/11P~50~0019601~0019650~~~0019601~0019616~16~16~0~~0~~0~~0019617~0019650~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001"
str2 = str2 & "~AA/11P~50~0019651~0019700~~~0019651~0019689~39~39~0~~0~~0~~0019690~0019700~11~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/0"
str2 = str2 & "01~AA/11P~50~0019701~0019750~~~0019701~0019701~1~1~0~~0~~0~~0019702~0019750~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012"
str2 = str2 & ""
str2 = str2 & "L~01GTKT2/001~AA/11P~50~0019751~0019800~~~0019751~0019758~8~8~0~~0~~0~~0019759~0019800~42~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0019801~0019850~~~0019801~0019806~6~6~0~~0~~0~~0019807~0019850~44~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~50~0019851~0019900~~~0019851~0019890~40~40~0~~0~~0~~0019891~0019900~10~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019901~0019950~~~0019901~0019940~40~40~0~~0~~0~~0019941~0019950~10~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~0019951~0020000~~~0019951~0019951~1~1~0~~0~~0~~0019952~0020000~49~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0020001~0020050~0020001~0020010~10~10~0~~0~~0~~0020011~0020050~40~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0020051~0020100~0020051~0020080~30~30~0~~0~~0~~0020081~0020100~20~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0020101~0020150~0020101~0020114~14~14~0~~0~~0~~0020115~0020150~36~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0020151~0020200~0020151~0020195~45~45~0~~0~~0~~0020196~0020200~5~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~800~~~0020201~0021000~~~0~0~0~~0~~0~~0020201~0021000~800~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng - BC012L~01GTKT2/001~AA/11P~1500~~~0021001~0022500~~~0~0~0~~0~~0~~0021001~0022500~1500~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC0"
str2 = str2 & "12L~01GTKT2/001~AA/11P~50~~~0022501~0022550~~~0~0~0~~0~~0~~0022501~0022550~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/00"
str2 = str2 & "1~AA/11P~50~~~0022551~0022600~0022551~0022555~5~5~0~~0~~0~~0022556~0022600~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/00"
str2 = str2 & "1~AA/11P~50~~~0022601~0022650~0022601~0022615~15~15~0~~0~~0~~0022616~0022650~35~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/"
str2 = str2 & "001~AA/11P~50~~~0022651~0022700~~~0~0~0~~0~~0~~0022651~0022700~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~"
str2 = str2 & "~~0022701~0022750~0022701~0022705~5~5~0~~0~~0~~0022706~0022750~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001"
str2 = str2 & ""
str2 = str2 & "~AA/11P~50~~~0022751~0022800~~~0~0~0~~0~~0~~0022751~0022800~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~~~0022801~0022850~0022801~0022801~1~1~0~~0~~0~~0022802~0022850~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~~~0022851~0022900~0022851~0022851~1~1~0~~0~~0~~0022852~0022900~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~50~~~0022901~0022950~0022901~0022905~5~5~0~~0~~0~~0022906~0022950~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~"
str2 = str2 & "AA/11P~550~~~0022951~0023500~~~0~0~0~~0~~0~~0022951~0023500~550~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~"
str2 = str2 & "0023501~0023550~0023501~0023550~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0023551~00236"
str2 = str2 & "00~0023551~0023597~47~47~0~~0~~0~~0023598~0023600~3~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0023601~0023"
str2 = str2 & "650~0023601~0023648~48~47~1~0023639~0~~0~~0023649~0023650~2~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0023"
str2 = str2 & "651~0023700~~~0~0~0~~0~~0~~0023651~0023700~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0023701~0023750~~~"
str2 = str2 & "0~0~0~~0~~0~~0023701~0023750~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0023751~0023800~0023751~0023752~"
str2 = str2 & "2~2~0~~0~~0~~0023753~0023800~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~200~~~0023801~0024000~~~0~0~0~~0~~0~~"
str2 = str2 & "0023801~0024000~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~500~~~0024001~0024500~~~0~0~0~~0~~0~~0024001~0024"
str2 = str2 & "500~500~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0024501~0024550~~~0~0~0~~0~~0~~0024501~0024550~50~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0024551~0024600~0024551~0024585~35~35~0~~0~~0~~0024586~0024600~15~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0024601~0024650~0024601~0024616~16~15~1~0024606~0~~0~~0024617~002465"
str2 = str2 & "0~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0024651~0024700~0024651~0024662~12~12~0~~0~~0~~0"
str2 = str2 & ""
str2 = str2 & "024663~0024700~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0024701~0024750~0024701~0024709~9"
str2 = str2 & "~9~0~~0~~0~~0024710~0024750~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0024751~0024800~0024751~0024753~3"
str2 = str2 & "~3~0~~0~~0~~0024754~0024800~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~200~~~0024801~0025000~~~0~0~0~~0~~0~~0"
str2 = str2 & "024801~0025000~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0025001~0025050~0025001~0025025~25~25~0~~0~~0"
str2 = str2 & "~~0025026~0025050~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~450~~~0025051~0025500~~~0~0~0~~0~~0~~0025051~002"
str2 = str2 & "5500~450~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0025501~0025550~0025501~0025501~1~0~1~0025501~0~~0~~002"
str2 = str2 & "5502~0025550~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~450~~~0025551~0026000~~~0~0~0~~0~~0~~0025551~0026000~"
str2 = str2 & "450~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0026001~0026050~0026001~0026006~6~6~0~~0~~0~~0026007~0026050"
str2 = str2 & "~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0026051~0026100~0026051~0026064~14~14~0~~0~~0~~0026065~00261"
str2 = str2 & "00~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0026101~0026150~0026101~0026110~10~10~0~~0~~0~~0026111~002"
str2 = str2 & "6150~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0026151~0026200~0026151~0026155~5~5~0~~0~~0~~0026156~002"
str2 = str2 & "6200~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0026201~0026250~0026201~0026201~1~1~0~~0~~0~~0026202~002"
str2 = str2 & "6250~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~250~~~0026251~0026500~~~0~0~0~~0~~0~~0026251~0026500~250~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~1000~~~0026501~0027500~~~0~0~0~~0~~0~~0026501~0027500~1000~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~50~~~0027501~0027550~0027501~0027516~16~16~0~~0~~0~~0027517~0027550~34~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~450~~~0027551~0028000~~~0~0~0~~0~~0~~0027551~0028000~450~0~H„a Æ¨n gi∏"
str2 = str2 & ""
str2 = str2 & " trﬁ gia t®ng - BC012L~01GTKT2/001~AA/11P~12000~~~0028001~0040000~~~0~0~0~~0~~0~~0028001~0040000~12000~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~1600~~~0000001~0001600~0000001~0001245~1245~1218~27~005;0038;0040;0043;308;32"
str2 = str2 & "1;328;350;370;146;209;216;279;533;534;581;831;880;888;776;631;293;1030;1077;1174;704;985~0~~0~~0001246~0001600~355~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~200~~~0001601~0001800~0001601~0001655~55~55~0~~0~~0~~0001656~0001800~145~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~200~~~0001801~0002000~0001801~0001843~43~43~0~~0~~0~~0001844~0002000~157~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~3500~~~0002001~0005500~~~0~0~0~~0~~0~~0002001~0005500~3500~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~50~~~0005501~0005550~0005501~0005524~24~24~0~~0~~0~~0005525~0005550~26~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~50~~~0005551~0005600~0005551~0005596~46~46~0~~0~~0~~0005597~0005600~4~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~50~~~0005601~0005650~0005601~0005623~23~23~0~~0~~0~~0005624~0005650~27~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~150~~~0005651~0005800~0005651~0005666~16~16~0~~0~~0~~0005667~0005800~134~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~100~~~0005801~0005900~0005801~0005837~37~37~0~~0~~0~~0005838~0005900~63"
str2 = str2 & "~0~H„a Æ¨n gi∏ trﬁ gia t®ng - BC012L~01GTKT2/001~AB/11P~100~~~0005901~0006000~0005901~0005917~17~17~0~~0~~0~~0005918~0006000"
str2 = str2 & "~83~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~5000~0030001~0035000~~~0030001~0035000~5000~0~0~~0~~5000~30001-"
str2 = str2 & "35000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~13~0093438~0093450~~~0093438~0093450~13~0~0~~0~~13~93438-"
str2 = str2 & "93450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~50~0093451~0093500~~~0093451~0093500~50~0~0~~0~~50~93451-"
str2 = str2 & "93500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~14~0098237~0098250~~~0098237~0098250~14~0~0~~0"
str2 = str2 & ""
str2 = str2 & "~~14~98237-98250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~50~0098251~0098300~~~0098251~0098"
str2 = str2 & "300~50~0~0~~0~~50~98251-98300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~250~0098301~0098550~~~0098301~009"
str2 = str2 & "8550~250~0~0~~0~~250~98301-98550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2008T~1~0070900~0070900~~~0070900~00"
str2 = str2 & "70900~1~0~0~~0~~1~70900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~32~0018219~0018250~~~0018219~0018250~32"
str2 = str2 & "~0~0~~0~~32~018219-018250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~20~0018331~0018350~~~0018331~0018350~"
str2 = str2 & "20~0~0~~0~~20~018331-018350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~200~0018351~0018550~~~0018351~00185"
str2 = str2 & "50~200~0~0~~0~~200~018351-018550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~42~0007009~0007050~~~0007009~0"
str2 = str2 & "007050~42~0~0~~0~~42~007009-007050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~19~0016382~0016400~~~0016382"
str2 = str2 & "~0016400~19~0~0~~0~~19~016382-016400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~2~0016749~0016750~~~001674"
str2 = str2 & "9~0016750~2~0~0~~0~~2~016749-016750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~50~0016751~0016800~~~001675"
str2 = str2 & "1~0016800~50~0~0~~0~~50~016751-016800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~19~0019182~0019200~~~0019"
str2 = str2 & "182~0019200~19~0~0~~0~~19~019182-019200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~50~0019501~0019550~~~00"
str2 = str2 & "19501~0019550~50~0~0~~0~~50~019501-019550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~17~0019134~0019150~~~"
str2 = str2 & "0019134~0019150~17~0~0~~0~~17~019134-019150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~26~0019875~0019900~"
str2 = str2 & "~~0019875~0019900~26~0~0~~0~~26~019875-019900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~50~0019901~001995"
str2 = str2 & "0~~~0019901~0019950~50~0~0~~0~~50~019901-019950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2009T~50~0"
str2 = str2 & ""
str2 = str2 & "019951~0020000~~~0019951~0020000~50~0~0~~0~~50~019951-020000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-0"
str2 = str2 & "2~XQ/2009T~9~0017492~0017500~~~0017492~0017500~9~0~0~~0~~9~017492-017500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02"
str2 = str2 & "~XQ/2009T~30~0017571~0017600~~~0017571~0017600~30~0~0~~0~~30~017571-017600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-"
str2 = str2 & "02~XQ/2009T~35~0017666~0017700~~~0017666~0017700~35~0~0~~0~~35~017666-017700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3L"
str2 = str2 & "N-02~XQ/2009T~350~0017701~0018050~~~0017701~0018050~350~0~0~~0~~350~017701-018050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GT"
str2 = str2 & "KT-3LN-02~XQ/2009T~350~0011951~0012300~~~0011951~0012300~350~0~0~~0~~350~011951-012300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L"
str2 = str2 & "~01GTKT-3LN-02~XQ/2009T~50~0011851~0011900~~~0011851~0011900~50~0~0~~0~~50~011851-011900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC01"
str2 = str2 & "3L~01GTKT-3LN-02~XQ/2009T~42~0011909~0011950~~~0011909~0011950~42~0~0~~0~~42~011909-011950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC"
str2 = str2 & "013L~01GTKT-3LN-02~XQ/2009T~30~0015771~0015800~~~0015771~0015800~30~0~0~~0~~30~015771-015800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-"
str2 = str2 & "BC013L~01GTKT-3LN-02~XQ/2009T~50~0010201~0010250~~~0010201~0010250~50~0~0~~0~~50~010201-010250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g-BC013L~01GTKT-3LN-02~XQ/2010T~550~0000001~0000550~~~0000001~0000550~550~0~0~~0~~550~000001-000550~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~36~0000965~0001000~~~0000965~0001000~36~0~0~~0~~36~965-1000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng-BC013L~01GTKT-3LN-02~XQ/2010T~50~0001001~0001050~~~0001001~0001050~50~0~0~~0~~50~1001-001050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng-BC013L~01GTKT-3LN-02~XQ/2010T~47~0000554~0000600~~~0000554~0000600~47~0~0~~0~~47~554-000600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®"
str2 = str2 & "ng-BC013L~01GTKT-3LN-02~XQ/2010T~47~0002104~0002150~~~0002104~0002150~47~0~0~~0~~47~2104-002150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®"
str2 = str2 & "ng-BC013L~01GTKT-3LN-02~XQ/2010T~100~0002151~0002250~~~0002151~0002250~100~0~0~~0~~100~2151-2250~~~0~0~H„a Æ¨n gi"
str2 = str2 & ""
str2 = str2 & "∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~44~0002407~0002450~~~0002407~0002450~44~0~0~~0~~44~2407-002450~~~0"
str2 = str2 & "~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~50~0002451~0002500~~~0002451~0002500~50~0~0~~0~~50~2451-002500~~~0"
str2 = str2 & "~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~350~0002651~0003000~~~0002651~0003000~350~0~0~~0~~350~2651-3000~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~48~0003153~0003200~~~0003153~0003200~48~0~0~~0~~48~3153-003200~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~50~0003201~0003250~~~0003201~0003250~50~0~0~~0~~50~3201-003250~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~250~0003351~0003600~~~0003351~0003600~250~0~0~~0~~250~3351-003600"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT-3LN-02~XQ/2010T~3900~0003601~0007500~~~0003601~0007500~3900~0~0~~0~~3900~3601-"
str2 = str2 & "007500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~950~0000001~0000950~~~0000001~0000950~950~948~2~410;754~0~~0"
str2 = str2 & "~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0000951~0001000~~~0000951~0000993~43~43~0~~0~~0~~0000994~00010"
str2 = str2 & "00~7~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~2950~0001001~0003950~~~~~0~~0~~0~~0~~0001001~0003950~2950~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0003951~0004000~~~0003951~0003955~5~5~0~~0~~0~~0003956~0004000~45~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0004001~0004050~~~0004001~0004050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g-BC013L~01GTKT3/001~AA/11P~300~0004051~0004350~~~~~0~~0~~0~~0~~0004051~0004350~300~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT"
str2 = str2 & "3/001~AA/11P~150~0004351~0004500~~~0004351~0004353~3~3~0~~0~~0~~0004354~0004500~147~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT"
str2 = str2 & "3/001~AA/11P~50~0004501~0004550~~~0004501~0004550~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50"
str2 = str2 & "~0004551~0004600~~~0004551~0004551~1~0~1~0004551~0~~0~~0004552~0004600~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTK"
str2 = str2 & ""
str2 = str2 & "T3/001~AA/11P~50~0004601~0004650~~~0004601~0004631~31~29~2~4601;4602~0~~0~~0004632~0004650~19~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0004651~0004700~~~0004651~0004700~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013"
str2 = str2 & "L~01GTKT3/001~AA/11P~50~0004701~0004750~~~0004701~0004750~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~A"
str2 = str2 & "A/11P~50~0004751~0004800~~~0004751~0004768~18~18~0~~0~~0~~0004769~0004800~32~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~A"
str2 = str2 & "A/11P~50~0004801~0004850~~~0004801~0004846~46~46~0~~0~~0~~0004847~0004850~4~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA"
str2 = str2 & "/11P~50~0004851~0004900~~~0004851~0004900~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0004901"
str2 = str2 & "~0004950~~~0004901~0004913~13~13~0~~0~~0~~0004914~0004950~37~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~200~000495"
str2 = str2 & "1~0005150~~~~~0~~0~~0~~0~~0004951~0005150~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005151~0005200~~~0005"
str2 = str2 & "151~0005200~50~49~1~5156~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005201~0005250~~~0005201~0005250"
str2 = str2 & "~50~48~2~5210;5247~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005251~0005300~~~0005251~0005300~50~48"
str2 = str2 & "~2~5257;005285~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005301~0005350~~~0005301~0005350~50~47~3~0"
str2 = str2 & "05308;5331;5349~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005351~0005400~~~0005351~0005400~50~50~0~"
str2 = str2 & "~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005401~0005450~~~0005401~0005416~16~16~0~~0~~0~~0005417~"
str2 = str2 & "0005450~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0005451~0005500~~~0005451~0005475~25~25~0~~0~~0~~0005476~"
str2 = str2 & "0005500~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~650~0005501~0006150~~~~~0~~0~~0~~0~~0005501~0006150~650~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0006151~0006200~~~0006151~0006200~50~50~0~~0~~0~~~~0~0~H„a Æ¨"
str2 = str2 & ""
str2 = str2 & "n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0006201~0006250~~~0006201~0006250~50~50~0~~0~~0~~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0006251~0006300~~~0006251~0006267~17~17~0~~0~~0~~0006268~0006300~33~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0006301~0006350~~~0006301~0006334~34~34~0~~0~~0~~0006335~0006350~16~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0006351~0006400~~~~~0~~0~~0~~0~~0006351~0006400~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-"
str2 = str2 & "BC013L~01GTKT3/001~AA/11P~50~0006401~0006450~~~0006401~0006450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/"
str2 = str2 & "001~AA/11P~50~0006451~0006500~~~0006451~0006500~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0"
str2 = str2 & "006501~0006550~~~0006501~0006542~42~42~0~~0~~0~~0006543~0006550~8~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~00"
str2 = str2 & "06551~0006600~~~~~0~~0~~0~~0~~0006551~0006600~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~550~0006601~0007150~~~"
str2 = str2 & "~~0~~0~~0~~0~~0006601~0007150~550~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0007151~0007200~~~0007151~0007200~"
str2 = str2 & "50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0007201~0007250~~~0007201~0007250~50~49~1~0007248"
str2 = str2 & "~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0007251~0007300~~~0007251~0007287~37~37~0~~0~~0~~0007288~"
str2 = str2 & "0007300~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0007301~0007350~~~0007301~0007309~9~9~0~~0~~0~~0007310~00"
str2 = str2 & "07350~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0007351~0007400~~~0007351~0007385~35~35~0~~0~~0~~0007386~00"
str2 = str2 & "07400~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0007401~0007450~~~~~0~~0~~0~~0~~0007401~0007450~50~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~200~0007451~0007650~~~~~0~~0~~0~~0~~0007451~0007650~200~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng-BC013L~01GTKT3/001~AA/11P~50~0007651~0007700~~~0007651~0007700~50~49~1~0007653~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & ""
str2 = str2 & "a t®ng-BC013L~01GTKT3/001~AA/11P~50~0007701~0007750~~~0007701~0007750~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BC013L~01GTKT3/001~AA/11P~50~0007751~0007800~~~0007751~0007800~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01G"
str2 = str2 & "TKT3/001~AA/11P~50~0007801~0007850~~~0007801~0007850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P"
str2 = str2 & "~50~0007851~0007900~~~0007851~0007856~6~6~0~~0~~0~~0007857~0007900~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~5"
str2 = str2 & "0~0007901~0007950~~~0007901~0007906~6~6~0~~0~~0~~0007907~0007950~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~"
str2 = str2 & "0007951~0008000~~~0007951~0008000~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008001~0008050"
str2 = str2 & "~~~0008001~0008038~38~38~0~~0~~0~~0008039~0008050~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~100~0008051~000815"
str2 = str2 & "0~~~~~0~0~0~~0~~0~~0008051~0008150~100~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008151~0008200~~~0008151~000"
str2 = str2 & "8200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008201~0008250~~~0008201~0008250~50~50~0~~0"
str2 = str2 & "~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008251~0008300~~~0008251~0008300~50~50~0~~0~~0~~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008301~0008350~~~0008301~0008350~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BC013L~01GTKT3/001~AA/11P~50~0008351~0008400~~~0008351~0008400~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01"
str2 = str2 & "GTKT3/001~AA/11P~50~0008401~0008450~~~0008401~0008450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11"
str2 = str2 & "P~50~0008451~0008500~~~0008451~0008500~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008501~00"
str2 = str2 & "08550~~~0008501~0008550~50~49~1~0008536~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008551~0008600~~~"
str2 = str2 & "0008551~0008600~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008601~0008650~~~0008"
str2 = str2 & ""
str2 = str2 & "601~0008650~50~49~1~0008605~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0008651~0008700~~"
str2 = str2 & "~0008651~0008675~25~25~0~~0~~0~~0008676~0008700~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~950~0008701~0009650~"
str2 = str2 & "~~~~0~0~0~~0~~0~~0008701~0009650~950~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0009651~0009700~~~0009651~00097"
str2 = str2 & "00~50~42~8~0009683-9690~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0009701~0009750~~~0009701~0009750~"
str2 = str2 & "50~49~1~0009730~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0009751~0009800~~~0009751~0009800~50~50~0~"
str2 = str2 & "~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0009801~0009850~~~0009801~0009850~50~50~0~~0~~0~~~~0~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0009851~0009900~~~0009851~0009900~50~49~1~0009868~0~~0~~~~0~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0009901~0009950~~~0009901~0009950~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-"
str2 = str2 & "BC013L~01GTKT3/001~AA/11P~50~0009951~0010000~~~0009951~0010000~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/"
str2 = str2 & "001~AA/11P~50~0010001~0010050~~~0010001~0010003~3~3~0~~0~~0~~0010004~0010050~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/00"
str2 = str2 & "1~AA/11P~50~0010051~0010100~~~~~0~0~0~~0~~0~~0010051~0010100~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0010"
str2 = str2 & "101~0010150~~~~~0~0~0~~0~~0~~0010101~0010150~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~800~0010151~0010950~~~0"
str2 = str2 & "010151~0010929~779~769~10~10371;10375;"
str2 = str2 & "10389;10503;10543;10582;10627;10654;10743;10824~0~~0~~0010930~0010950~21~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~200~0010951~0011150~~~~~0~0~0~~0~~0~~0010951~0011150~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BC013L~01GTKT3/001~AA/11P~50~0011151~0011200~~~0011151~0011175~25~25~0~~0~~0~~0011176~0011200~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BC013L~01GTKT3/001~AA/11P~450~0011201~0011650~~~~~0~0~0~~0~~0~~0011201~0011650~450~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC"
str2 = str2 & "013L~01GTKT3/001~AA/11P~50~0011651~0011700~~~0011651~0011700~50~48~2~0011663;0011676~0~~0~~~~0~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0011701~0011750~~~0011701~0011708~8~8~0~~0~~0~~0011709~0011750~42~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0011751~0011800~~~0011751~0011800~50~49~1~0011759~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BC013L~01GTKT3/001~AA/11P~50~0011801~0011850~~~0011801~0011850~50~49~1~0011802~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~"
str2 = str2 & "01GTKT3/001~AA/11P~50~0011851~0011900~~~0011851~0011900~50~49~1~0011865~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/"
str2 = str2 & "001~AA/11P~50~0011901~0011950~~~0011901~0011945~45~45~0~~0~~0~~0011946~0011950~5~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/0"
str2 = str2 & "01~AA/11P~200~0011951~0012150~~~~~0~0~0~~0~~0~~0011951~0012150~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0"
str2 = str2 & "012151~0012200~~~0012151~0012192~42~41~1~0012192~0~~0~~0012193~0012200~8~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11"
str2 = str2 & "P~400~0012201~0012600~~~0012201~0012600~400~400~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0012601"
str2 = str2 & "~0012650~~~0012601~0012607~7~7~0~~0~~0~~0012608~0012650~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~200~0012651~"
str2 = str2 & "0012850~~~~~0~0~0~~0~~0~~0012651~0012850~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~200~0012851~0013050~~~~~0~"
str2 = str2 & "0~0~~0~~0~~0012851~0013050~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0013051~0013100~~~0013051~0013059~9~9"
str2 = str2 & "~0~~0~~0~~0013060~0013100~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~300~0013101~0013400~~~~~0~0~0~~0~~0~~00131"
str2 = str2 & "01~0013400~300~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0013401~0013450~~~~~0~0~0~~0~~0~~0013401~0013450~50~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0013451~0013500~~~~~0~0~0~~0~~0~~0013451~0013500~50~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BC013L~01GTKT3/001~AA/11P~50~0013501~0013550~~~~~0~0~0~~0~~0~~0013501~0013550~50~0~H„a Æ¨n gi∏ trﬁ gia "
str2 = str2 & ""
str2 = str2 & "t®ng-BC013L~01GTKT3/001~AA/11P~6450~0013551~0020000~~~~~0~0~0~~0~~0~~0013551~0020000~6450~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BC013L~01GTKT3/001~AA/11P~5000~~~0020001~0025000~~~0~0~0~~0~~0~~0020001~0025000~5000~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL"
str2 = str2 & "~01GTKT-2LL-07~XQ/2008T~20~0025331~0025350~~~0025331~0025350~20~0~0~~0~~20~025331-025350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02"
str2 = str2 & "LL~01GTKT-2LL-07~XQ/2008T~41~0034610~0034650~~~0034610~0034650~41~0~0~~0~~41~034610-034650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "02LL~01GTKT-2LL-07~XQ/2008T~24~0025377~0025400~~~0025377~0025400~24~0~0~~0~~24~025377-025400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-"
str2 = str2 & "BD02LL~01GTKT-2LL-07~XQ/2008T~40~0034261~0034300~~~0034261~0034300~40~0~0~~0~~40~034261-034300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g-BD02LL~01GTKT-2LL-07~XQ/2008T~47~0025404~0025450~~~0025404~0025450~47~0~0~~0~~47~025404-025450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~35~0034316~0034350~~~0034316~0034350~35~0~0~~0~~35~034316-034350~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~16~0025485~0025500~~~0025485~0025500~16~0~0~~0~~16~025485-025500~~~0~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~10~0025591~0025600~~~0025591~0025600~10~0~0~~0~~10~025591-025600~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~16~0025635~0025650~~~0025635~0025650~16~0~0~~0~~16~025635-025650~~~0~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~20~0025681~0025700~~~0025681~0025700~20~0~0~~0~~20~025681-025700~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~41~0034360~0034400~~~0034360~0034400~41~0~0~~0~~41~034360-034400~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~50~0034401~0034450~~~0034401~0034450~50~0~0~~0~~50~034401-034450~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~43~0034458~0034500~~~0034458~0034500~43~0~0~~0~~43~034458-034500~~~0~0~H„a Æ"
str2 = str2 & "¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~31~0034520~0034550~~~0034520~0034550~31~0~0~~0~~31~034520-03455"
str2 = str2 & ""
str2 = str2 & "0~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~19~0025182~0025200~~~0025182~0025200~19~0~0~~0~~"
str2 = str2 & "19~025182-025200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~9~0025242~0025250~~~0025242~0025250~9~0~0~~0~~"
str2 = str2 & "9~025242-025250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~34~0025267~0025300~~~0025267~0025300~34~0~0~~0~"
str2 = str2 & "~34~025267-025300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~31~0025720~0025750~~~0025720~0025750~31~0~0~~"
str2 = str2 & "0~~31~025720-025750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~50~0025751~0025800~~~0025751~0025800~50~0~0"
str2 = str2 & "~~0~~50~025751-025800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~50~0025801~0025850~~~0025801~0025850~50~0"
str2 = str2 & "~0~~0~~50~025801-025850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~34~0025867~0025900~~~0025867~0025900~34"
str2 = str2 & "~0~0~~0~~34~025867-025900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~27~0025924~0025950~~~0025924~0025950~"
str2 = str2 & "27~0~0~~0~~27~025924-025950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~27~0025074~0025100~~~0025074~002510"
str2 = str2 & "0~27~0~0~~0~~27~025074-025100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~40~0034011~0034050~~~0034011~0034"
str2 = str2 & "050~40~0~0~~0~~40~034011-034050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~41~0034060~0034100~~~0034060~00"
str2 = str2 & "34100~41~0~0~~0~~41~034060-034100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~38~0034113~0034150~~~0034113~"
str2 = str2 & "0034150~38~0~0~~0~~38~034113-034150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~49~0034152~0034200~~~003415"
str2 = str2 & "2~0034200~49~0~0~~0~~49~034152-034200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~47~0034654~0034700~~~0034"
str2 = str2 & "654~0034700~47~0~0~~0~~47~034654-034700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~40~0034211~0034250~~~00"
str2 = str2 & "34211~0034250~40~0~0~~0~~40~034211-034250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~42~0034559"
str2 = str2 & ""
str2 = str2 & "~0034600~~~0034559~0034600~42~0~0~~0~~42~034559-034600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2"
str2 = str2 & "008T~300~0034701~0035000~~~0034701~0035000~300~0~0~~0~~300~034701-035000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07"
str2 = str2 & "~XQ/2008T~1000~0035001~0036000~~~0035001~0036000~1000~0~0~~0~~1000~035001-036000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTK"
str2 = str2 & "T-2LL-07~XQ/2008T~7~0028744~0028750~~~0028744~0028750~7~0~0~~0~~7~028744-028750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT"
str2 = str2 & "-2LL-07~XQ/2008T~38~0028513~0028550~~~0028513~0028550~38~0~0~~0~~38~028513-028550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GT"
str2 = str2 & "KT-2LL-07~XQ/2008T~37~0028614~0028650~~~0028614~0028650~37~0~0~~0~~37~028614-028650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01"
str2 = str2 & "GTKT-2LL-07~XQ/2008T~39~0028962~0029000~~~0028962~0029000~39~0~0~~0~~39~028962-029000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~"
str2 = str2 & "01GTKT-2LL-07~XQ/2008T~19~0029132~0029150~~~0029132~0029150~19~0~0~~0~~19~029132-029150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02L"
str2 = str2 & "L~01GTKT-2LL-07~XQ/2008T~50~0029151~0029200~~~0029151~0029200~50~0~0~~0~~50~029151-029200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD0"
str2 = str2 & "2LL~01GTKT-2LL-07~XQ/2008T~50~0029001~0029050~~~0029001~0029050~50~0~0~~0~~50~029001-029050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D02LL~01GTKT-2LL-07~XQ/2008T~22~0029229~0029250~~~0029229~0029250~22~0~0~~0~~22~029229-029250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2008T~50~0029251~0029300~~~0029251~0029300~50~0~0~~0~~50~029251-029300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®"
str2 = str2 & "ng-BD02LL~01GTKT-2LL-07~XQ/2008T~35~0029316~0029350~~~0029316~0029350~35~0~0~~0~~35~029316-029350~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~28~0028673~0028700~~~0028673~0028700~28~0~0~~0~~28~028673-028700~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~24~0028577~0028600~~~0028577~0028600~24~0~0~~0~~24~028577-028600~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~150~0029351~0029500~~~0029351~0029500~150~0~0~~0~~150~029351-029500~~~0~0~"
str2 = str2 & ""
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~1000~0032001~0033000~~~0032001~0033000~1000~0~0~~0~~1000"
str2 = str2 & "~032001-033000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~30~0029621~0029650~~~0029621~0029650~30~0~0~~0~~"
str2 = str2 & "30~029621-029650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~32~0029669~0029700~~~0029669~0029700~32~0~0~~0"
str2 = str2 & "~~32~029669-029700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~100~0029701~0029800~~~0029701~0029800~100~0~"
str2 = str2 & "0~~0~~100~029701-029800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~40~0029811~0029850~~~0029811~0029850~40"
str2 = str2 & "~0~0~~0~~40~029811-029850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~28~0029873~0029900~~~0029873~0029900~"
str2 = str2 & "28~0~0~~0~~28~029873-029900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~24~0029927~0029950~~~0029927~002995"
str2 = str2 & "0~24~0~0~~0~~24~029927-029950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~41~0029960~0030000~~~0029960~0030"
str2 = str2 & "000~41~0~0~~0~~41~029960-030000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~20~0030031~0030050~~~0030031~00"
str2 = str2 & "30050~20~0~0~~0~~20~030031-030050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~48~0030053~0030100~~~0030053~"
str2 = str2 & "0030100~48~0~0~~0~~48~030053-030100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~45~0030156~0030200~~~003015"
str2 = str2 & "6~0030200~45~0~0~~0~~45~030156-030200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~42~0030209~0030250~~~0030"
str2 = str2 & "209~0030250~42~0~0~~0~~42~030209-030250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~250~0030251~0030500~~~0"
str2 = str2 & "030251~0030500~250~0~0~~0~~250~030251-030500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~1000~0033001~00340"
str2 = str2 & "00~~~0033001~0034000~1000~0~0~~0~~1000~033001-034000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~13000~0037"
str2 = str2 & "001~0050000~~~0037001~0050000~13000~0~0~~0~~13000~037001-050000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-"
str2 = str2 & ""
str2 = str2 & "07~XQ/2008T~45~0030906~0030950~~~0030906~0030950~45~0~0~~0~~45~030906-030950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD0"
str2 = str2 & "2LL~01GTKT-2LL-07~XQ/2008T~14~0030737~0030750~~~0030737~0030750~14~0~0~~0~~14~030737-030750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D02LL~01GTKT-2LL-07~XQ/2008T~39~0036062~0036100~~~0036062~0036100~39~0~0~~0~~39~036062-036100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2008T~19~0030532~0030550~~~0030532~0030550~19~0~0~~0~~19~030532-030550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®"
str2 = str2 & "ng-BD02LL~01GTKT-2LL-07~XQ/2008T~45~0030656~0030700~~~0030656~0030700~45~0~0~~0~~45~030656-030700~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~47~0036004~0036050~~~0036004~0036050~47~0~0~~0~~47~036004-036050~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~46~0030805~0030850~~~0030805~0030850~46~0~0~~0~~46~030805-030850~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~37~0030614~0030650~~~0030614~0030650~37~0~0~~0~~37~030614-030650~~~0~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~900~0036101~0037000~~~0036101~0037000~900~0~0~~0~~900~036101-037000~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~29~0027022~0027050~~~0027022~0027050~29~0~0~~0~~29~027022-027050~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~30~0027121~0027150~~~0027121~0027150~30~0~0~~0~~30~027121-027150~~~0~0~H„a Æ"
str2 = str2 & "¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~600~0027151~0027750~~~0027151~0027750~600~0~0~~0~~600~027151-027750~~~0~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~10~0027841~0027850~~~0027841~0027850~10~0~0~~0~~10~027841-027850~~~0~"
str2 = str2 & "0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~38~0027913~0027950~~~0027913~0027950~38~0~0~~0~~38~027913-027950~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~39~0027962~0028000~~~0027962~0028000~39~0~0~~0~~39~027962-028000~"
str2 = str2 & "~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~500~0028001~0028500~~~0028001~0028500~500~0~0~~0~~50"
str2 = str2 & ""
str2 = str2 & "0~028001-028500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~12~0026139~0026150~~~0026139~00261"
str2 = str2 & "50~12~0~0~~0~~12~026139-026150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~21~0026180~0026200~~~0026180~002"
str2 = str2 & "6200~21~0~0~~0~~21~026180-026200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~50~0026201~0026250~~~0026201~0"
str2 = str2 & "026250~50~0~0~~0~~50~026201-026250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~150~0026501~0026650~~~002650"
str2 = str2 & "1~0026650~150~0~0~~0~~150~026501-026650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~350~0026651~0027000~~~0"
str2 = str2 & "026651~0027000~350~0~0~~0~~350~026651-027000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~31~0031220~0031250"
str2 = str2 & "~~~0031220~0031250~31~0~0~~0~~31~031220-031250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~13~0031288~00313"
str2 = str2 & "00~~~0031288~0031300~13~0~0~~0~~13~031288-031300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~47~0031404~003"
str2 = str2 & "1450~~~0031404~0031450~47~0~0~~0~~47~031404-031450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~45~0031656~0"
str2 = str2 & "031700~~~0031656~0031700~45~0~0~~0~~45~031656-031700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~300~003170"
str2 = str2 & "1~0032000~~~0031701~0032000~300~0~0~~0~~300~031701-032000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~24~01"
str2 = str2 & "21977~0122000~~~0121977~0122000~24~0~0~~0~~24~121977-122000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~33~"
str2 = str2 & "0122518~0122550~~~0122518~0122550~33~0~0~~0~~33~122518-122550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~3"
str2 = str2 & "~0097948~0097950~~~0097948~0097950~3~0~0~~0~~3~97948-97950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~25~0"
str2 = str2 & "122576~0122600~~~0122576~0122600~25~0~0~~0~~25~122576-122600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~30"
str2 = str2 & "~0127721~0127750~~~0127721~0127750~30~0~0~~0~~30~127721-127750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-0"
str2 = str2 & ""
str2 = str2 & "7~XQ/2008T~31~0132870~0132900~~~0132870~0132900~31~0~0~~0~~31~132870-132900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02"
str2 = str2 & "LL~01GTKT-2LL-07~XQ/2008T~22~0129229~0129250~~~0129229~0129250~22~0~0~~0~~22~129229-129250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "02LL~01GTKT-2LL-07~XQ/2008T~42~0079809~0079850~~~0079809~0079850~42~0~0~~0~~42~79809-79850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "02LL~01GTKT-2LL-07~XQ/2008T~5~0099396~0099400~~~0099396~0099400~5~0~0~~0~~5~99396-99400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02L"
str2 = str2 & "L~01GTKT-2LL-07~XQ/2008T~18~0129483~0129500~~~0129483~0129500~18~0~0~~0~~18~129483-129500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD0"
str2 = str2 & "2LL~01GTKT-2LL-07~XQ/2008T~21~0129080~0129100~~~0129080~0129100~21~0~0~~0~~21~129080-129100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D02LL~01GTKT-2LL-07~XQ/2008T~29~0079172~0079200~~~0079172~0079200~29~0~0~~0~~29~79172-79200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D02LL~01GTKT-2LL-07~XQ/2008T~50~0129151~0129200~~~0129151~0129200~50~0~0~~0~~50~129151-129200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2008T~29~0079072~0079100~~~0079072~0079100~29~0~0~~0~~29~79072-79100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2008T~44~0079857~0079900~~~0079857~0079900~44~0~0~~0~~44~79857-79900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2008T~20~0129031~0129050~~~0129031~0129050~20~0~0~~0~~20~129031-129050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®"
str2 = str2 & "ng-BD02LL~01GTKT-2LL-07~XQ/2008T~26~0120825~0120850~~~0120825~0120850~26~0~0~~0~~26~120825-120850~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~41~0120860~0120900~~~0120860~0120900~41~0~0~~0~~41~120860-120900~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~26~0121475~0121500~~~0121475~0121500~26~0~0~~0~~26~121475-121500~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~46~0120255~0120300~~~0120255~0120300~46~0~0~~0~~46~120255-120300~~~0~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~32~0097419~0097450~~~0097419~0097450~32~0~0~~0~~32~097419-097450~~~0~0~H"
str2 = str2 & ""
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~28~0120973~0121000~~~0120973~0121000~28~0~0~~0~~28~120973"
str2 = str2 & "-121000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~30~0120921~0120950~~~0120921~0120950~30~0~0~~0~~30~1209"
str2 = str2 & "21-120950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~10~0097391~0097400~~~0097391~0097400~10~0~0~~0~~10~09"
str2 = str2 & "7391-097400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~35~0098266~0098300~~~0098266~0098300~35~0~0~~0~~35~"
str2 = str2 & "098266-098300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~27~0127324~0127350~~~0127324~0127350~27~0~0~~0~~2"
str2 = str2 & "7~127324-127350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~25~0097676~0097700~~~0097676~0097700~25~0~0~~0~"
str2 = str2 & "~25~097676-097700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~29~0121272~0121300~~~0121272~0121300~29~0~0~~"
str2 = str2 & "0~~29~121272-121300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~28~0078173~0078200~~~0078173~0078200~28~0~0"
str2 = str2 & "~~0~~28~078173-078200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~17~0094134~0094150~~~0094134~0094150~17~0"
str2 = str2 & "~0~~0~~17~094134-094150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~7~0082694~0082700~~~0082694~0082700~7~0"
str2 = str2 & "~0~~0~~7~082694-082700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~11~0082840~0082850~~~0082840~0082850~11~"
str2 = str2 & "0~0~~0~~11~082840-082850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~100~0082901~0083000~~~0082901~0083000~"
str2 = str2 & "100~0~0~~0~~100~082901-083000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~35~0082866~0082900~~~0082866~0082"
str2 = str2 & "900~35~0~0~~0~~35~082866-082900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~8~0130643~0130650~~~0130643~013"
str2 = str2 & "0650~8~0~0~~0~~8~130643-130650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2008T~22~0081329~0081350~~~0081329~008"
str2 = str2 & "1350~22~0~0~~0~~22~081329-081350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~37~0016264~0016300~"
str2 = str2 & ""
str2 = str2 & "~~0016264~0016300~37~0~0~~0~~37~016264-016300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~12~0"
str2 = str2 & "016339~0016350~~~0016339~0016350~12~0~0~~0~~12~016339-016350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~34"
str2 = str2 & "~0016367~0016400~~~0016367~0016400~34~0~0~~0~~34~016367-016400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~"
str2 = str2 & "19~0016482~0016500~~~0016482~0016500~19~0~0~~0~~19~016482-016500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009"
str2 = str2 & "T~32~0042519~0042550~~~0042519~0042550~32~0~0~~0~~32~042519-042550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/20"
str2 = str2 & "09T~200~0042551~0042750~~~0042551~0042750~200~0~0~~0~~200~042551-042750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~8~0016543~0016550~~~0016543~0016550~8~0~0~~0~~8~016543-016550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~X"
str2 = str2 & "Q/2009T~21~0016630~0016650~~~0016630~0016650~21~0~0~~0~~21~016630-016650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07"
str2 = str2 & "~XQ/2009T~27~0016724~0016750~~~0016724~0016750~27~0~0~~0~~27~016724-016750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-"
str2 = str2 & "07~XQ/2009T~12~0014439~0014450~~~0014439~0014450~12~0~0~~0~~12~014439-014450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2L"
str2 = str2 & "L-07~XQ/2009T~42~0021759~0021800~~~0021759~0021800~42~0~0~~0~~42~021759-021800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-"
str2 = str2 & "2LL-07~XQ/2009T~21~0030880~0030900~~~0030880~0030900~21~0~0~~0~~21~030880-030900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTK"
str2 = str2 & "T-2LL-07~XQ/2009T~50~0030951~0031000~~~0030951~0031000~50~0~0~~0~~50~030951-031000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01G"
str2 = str2 & "TKT-2LL-07~XQ/2009T~14~0035737~0035750~~~0035737~0035750~14~0~0~~0~~14~035737-035750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~0"
str2 = str2 & "1GTKT-2LL-07~XQ/2009T~2~0024749~0024750~~~0024749~0024750~2~0~0~~0~~2~024749-024750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01"
str2 = str2 & "GTKT-2LL-07~XQ/2009T~24~0039877~0039900~~~0039877~0039900~24~0~0~~0~~24~039877-039900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & ""
str2 = str2 & "®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~42~0040609~0040650~~~0040609~0040650~42~0~0~~0~~42~040609-040650~~~0~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~13~0040788~0040800~~~0040788~0040800~13~0~0~~0~~13~040788-040800~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~46~0033005~0033050~~~0033005~0033050~46~0~0~~0~~46~033005-033050~~~0~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~22~0040929~0040950~~~0040929~0040950~22~0~0~~0~~22~040929-040950~~~0~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~37~0040964~0041000~~~0040964~0041000~37~0~0~~0~~37~040964-041000~~~0~"
str2 = str2 & "0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~30~0041421~0041450~~~0041421~0041450~30~0~0~~0~~30~041421-041450~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~37~0041364~0041400~~~0041364~0041400~37~0~0~~0~~37~041364-041400~"
str2 = str2 & "~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~22~0024329~0024350~~~0024329~0024350~22~0~0~~0~~22~024329-02435"
str2 = str2 & "0~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~36~0009965~0010000~~~0009965~0010000~36~0~0~~0~~36~009965-010"
str2 = str2 & "000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0010451~0010500~~~0010451~0010500~50~0~0~~0~~50~010451-0"
str2 = str2 & "10500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~29~0009572~0009600~~~0009572~0009600~29~0~0~~0~~29~009572"
str2 = str2 & "-009600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~41~0036210~0036250~~~0036210~0036250~41~0~0~~0~~41~0362"
str2 = str2 & "10-036250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~44~0035857~0035900~~~0035857~0035900~44~0~0~~0~~44~03"
str2 = str2 & "5857-035900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~1~0036400~0036400~~~0036400~0036400~1~0~0~~0~~1~036"
str2 = str2 & "400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0036401~0036450~~~0036401~0036450~50~0~0~~0~~50~036401-0"
str2 = str2 & "36450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0036151~0036200~~~0036151~0036200~50~0~0~~0"
str2 = str2 & ""
str2 = str2 & "~~50~036151-036200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~17~0035934~0035950~~~0035934~00"
str2 = str2 & "35950~17~0~0~~0~~17~035934-035950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~49~0036302~0036350~~~0036302~"
str2 = str2 & "0036350~49~0~0~~0~~49~036302-036350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0036451~0036500~~~003645"
str2 = str2 & "1~0036500~50~0~0~~0~~50~036451-036500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0036501~0036550~~~0036"
str2 = str2 & "501~0036550~50~0~0~~0~~50~036501-036550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~700~0036551~0037250~~~0"
str2 = str2 & "036551~0037250~700~0~0~~0~~700~036551-037250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~29~0036272~0036300"
str2 = str2 & "~~~0036272~0036300~29~0~0~~0~~29~036272-036300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~41~0044410~00444"
str2 = str2 & "50~~~0044410~0044450~41~0~0~~0~~41~044410-044450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~36~0012315~001"
str2 = str2 & "2350~~~0012315~0012350~36~0~0~~0~~36~012315-012350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~16~0005635~0"
str2 = str2 & "005650~~~0005635~0005650~16~0~0~~0~~16~005635-005650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~9~0020942~"
str2 = str2 & "0020950~~~0020942~0020950~9~0~0~~0~~9~020942-020950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~36~0041815~"
str2 = str2 & "0041850~~~0041815~0041850~36~0~0~~0~~36~041815-041850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~8~0029343"
str2 = str2 & "~0029350~~~0029343~0029350~8~0~0~~0~~8~029343-029350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~42~0035209"
str2 = str2 & "~0035250~~~0035209~0035250~42~0~0~~0~~42~035209-035250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~9~003454"
str2 = str2 & "2~0034550~~~0034542~0034550~9~0~0~~0~~9~034542-034550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~18~004423"
str2 = str2 & "3~0044250~~~0044233~0044250~18~0~0~~0~~18~044233-044250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/20"
str2 = str2 & ""
str2 = str2 & "09T~9~0044392~0044400~~~0044392~0044400~9~0~0~~0~~9~044392-044400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-"
str2 = str2 & "2LL-07~XQ/2009T~30~0006371~0006400~~~0006371~0006400~30~0~0~~0~~30~006371-006400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTK"
str2 = str2 & "T-2LL-07~XQ/2009T~3~0043848~0043850~~~0043848~0043850~3~0~0~~0~~3~043848-043850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT"
str2 = str2 & "-2LL-07~XQ/2009T~33~0043418~0043450~~~0043418~0043450~33~0~0~~0~~33~043418-043450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GT"
str2 = str2 & "KT-2LL-07~XQ/2009T~39~0032062~0032100~~~0032062~0032100~39~0~0~~0~~39~032062-032100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01"
str2 = str2 & "GTKT-2LL-07~XQ/2009T~36~0032265~0032300~~~0032265~0032300~36~0~0~~0~~36~032265-032300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~"
str2 = str2 & "01GTKT-2LL-07~XQ/2009T~25~0018326~0018350~~~0018326~0018350~25~0~0~~0~~25~018326-018350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02L"
str2 = str2 & "L~01GTKT-2LL-07~XQ/2009T~44~0043457~0043500~~~0043457~0043500~44~0~0~~0~~44~043457-043500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD0"
str2 = str2 & "2LL~01GTKT-2LL-07~XQ/2009T~1~0008700~0008700~~~0008700~0008700~1~0~0~~0~~1~008700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GT"
str2 = str2 & "KT-2LL-07~XQ/2009T~21~0032230~0032250~~~0032230~0032250~21~0~0~~0~~21~032230-032250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01"
str2 = str2 & "GTKT-2LL-07~XQ/2009T~10~0018491~0018500~~~0018491~0018500~10~0~0~~0~~10~018491-018500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~"
str2 = str2 & "01GTKT-2LL-07~XQ/2009T~43~0032108~0032150~~~0032108~0032150~43~0~0~~0~~43~032108-032150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02L"
str2 = str2 & "L~01GTKT-2LL-07~XQ/2009T~35~0019616~0019650~~~0019616~0019650~35~0~0~~0~~35~019616-019650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD0"
str2 = str2 & "2LL~01GTKT-2LL-07~XQ/2009T~50~0019651~0019700~~~0019651~0019700~50~0~0~~0~~50~019651-019700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D02LL~01GTKT-2LL-07~XQ/2009T~50~0020201~0020250~~~0020201~0020250~50~0~0~~0~~50~020201-020250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0020251~0020300~~~0020251~0020300~50~0~0~~0~~50~020251-020300~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & ""
str2 = str2 & " trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0020301~0020350~~~0020301~0020350~50~0~0~~0~~50~020301-020350~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0020351~0020400~~~0020351~0020400~50~0~0~~0~~50~020351-020400~"
str2 = str2 & "~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0020401~0020450~~~0020401~0020450~50~0~0~~0~~50~020401-02045"
str2 = str2 & "0~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~8~0020493~0020500~~~0020493~0020500~8~0~0~~0~~8~020493-020500"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~250~0020501~0020750~~~0020501~0020750~250~0~0~~0~~250~20501-20"
str2 = str2 & "750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~3~0020048~0020050~~~0020048~0020050~3~0~0~~0~~3~20048-20050"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0020051~0020100~~~0020051~0020100~50~0~0~~0~~50~20051-20100"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0020101~0020150~~~0020101~0020150~50~0~0~~0~~50~20101-20150"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0019851~0019900~~~0019851~0019900~50~0~0~~0~~50~19851-19900"
str2 = str2 & "~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~1~015900-~0015900~~~0015900~0015900~1~0~0~~0~~1~015900-015900~"
str2 = str2 & "~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~33~0028068~0028100~~~0028068~0028100~33~0~0~~0~~33~028068-02810"
str2 = str2 & "0~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~18~0028133~0028150~~~0028133~0028150~18~0~0~~0~~18~028133-028"
str2 = str2 & "150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~19~0028332~0028350~~~0028332~0028350~19~0~0~~0~~19~028332-0"
str2 = str2 & "28350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~24~0028427~0028450~~~0028427~0028450~24~0~0~~0~~24~028427"
str2 = str2 & "-028450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~25~0037876~0037900~~~0037876~0037900~25~0~0~~0~~25~0378"
str2 = str2 & "76-037900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~4~0037997~0038000~~~0037997~0038000~4~0~0~"
str2 = str2 & ""
str2 = str2 & "~0~~4~037997-038000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~19~0038332~0038350~~~0038332~0"
str2 = str2 & "038350~19~0~0~~0~~19~038332-038350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~11~0038540~0038550~~~0038540"
str2 = str2 & "~0038550~11~0~0~~0~~11~038540-038550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~39~0038562~0038600~~~00385"
str2 = str2 & "62~0038600~39~0~0~~0~~39~038562-038600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~40~0038711~0038750~~~003"
str2 = str2 & "8711~0038750~40~0~0~~0~~40~038711-038750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~7~0027194~0027200~~~00"
str2 = str2 & "27194~0027200~7~0~0~~0~~7~027194-027200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~31~0022370~0022400~~~00"
str2 = str2 & "22370~0022400~31~0~0~~0~~31~022370-022400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~25~0022526~0022550~~~"
str2 = str2 & "0022526~0022550~25~0~0~~0~~25~022526-022550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~8~0022843~0022850~~"
str2 = str2 & "~0022843~0022850~8~0~0~~0~~8~022843-022850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~14~0033937~0033950~~"
str2 = str2 & "~0033937~0033950~14~0~0~~0~~14~033937-033950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~34~0022867~0022900"
str2 = str2 & "~~~0022867~0022900~34~0~0~~0~~34~022867-022900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~44~0033807~00338"
str2 = str2 & "50~~~0033807~0033850~44~0~0~~0~~44~033807-033850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~50~0039151~003"
str2 = str2 & "9200~~~0039151~0039200~50~0~0~~0~~50~039151-039200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~48~0039653~0"
str2 = str2 & "039700~~~0039653~0039700~48~0~0~~0~~48~39653-39700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~11~0039740~0"
str2 = str2 & "039750~~~0039740~0039750~11~0~0~~0~~11~39740-39750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~3~0034048~00"
str2 = str2 & "34050~~~0034048~0034050~3~0~0~~0~~3~34048-34050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~10~0"
str2 = str2 & ""
str2 = str2 & "038791~0038800~~~0038791~0038800~10~0~0~~0~~10~38791-38800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~25~0038926~0038950~~~0038926~0038950~25~0~0~~0~~25~38926-38950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~24~0039627~0039650~~~0039627~0039650~24~0~0~~0~~24~39627-39650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~12~0039589~0039600~~~0039589~0039600~12~0~0~~0~~12~39589-39600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~18~0039433~0039450~~~0039433~0039450~18~0~0~~0~~18~39433-39450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~39~0039462~0039500~~~0039462~0039500~39~0~0~~0~~39~39462-39500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~14~0039087~0039100~~~0039087~0039100~14~0~0~~0~~14~39087-39100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~"
str2 = str2 & "XQ/2009T~5~0004796~0004800~~~0004796~0004800~5~0~0~~0~~5~004796-004800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~X"
str2 = str2 & "Q/2009T~39~0019012~0019050~~~0019012~0019050~39~0~0~~0~~39~019012-019050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07"
str2 = str2 & "~XQ/2009T~18~0019083~0019100~~~0019083~0019100~18~0~0~~0~~18~019083-019100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-"
str2 = str2 & "07~XQ/2009T~50~0023201~0023250~~~0023201~0023250~50~0~0~~0~~50~023201-023250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2L"
str2 = str2 & "L-07~XQ/2009T~27~0018774~0018800~~~0018774~0018800~27~0~0~~0~~27~018774-018800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-"
str2 = str2 & "2LL-07~XQ/2009T~50~0018801~0018850~~~0018801~0018850~50~0~0~~0~~50~018801-018850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTK"
str2 = str2 & "T-2LL-07~XQ/2009T~44~0023157~0023200~~~0023157~0023200~44~0~0~~0~~44~023157-023200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01G"
str2 = str2 & "TKT-2LL-07~XQ/2009T~10~0023041~0023050~~~0023041~0023050~10~0~0~~0~~10~023041-023050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~0"
str2 = str2 & "1GTKT-2LL-07~XQ/2009T~8~0022993~0023000~~~0022993~0023000~8~0~0~~0~~8~022993-023000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & ""
str2 = str2 & "g-BD02LL~01GTKT-2LL-07~XQ/2009T~31~0022220~0022250~~~0022220~0022250~31~0~0~~0~~31~022220-022250~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~4~0022047~0022050~~~0022047~0022050~4~0~0~~0~~4~022047-022050~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~4~0023097~0023100~~~0023097~0023100~4~0~0~~0~~4~023097-023100~~~0~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~4~0008297~0008300~~~0008297~0008300~4~0~0~~0~~4~008297-008300~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~33~0007368~0007400~~~0007368~0007400~33~0~0~~0~~33~007368-007400~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~16~0007485~0007500~~~0007485~0007500~16~0~0~~0~~16~007485-007500~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~17~0014684~0014700~~~0014684~0014700~17~0~0~~0~~17~014684-014700~~~0~0~H„a Æ"
str2 = str2 & "¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~31~0037320~0037350~~~0037320~0037350~31~0~0~~0~~31~037320-037350~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2009T~13~0037588~0037600~~~0037588~0037600~13~0~0~~0~~13~037588-037600~~~0~0~H"
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~19~0002032~0002050~~~0002032~0002050~19~0~0~~0~~19~002032-002050~~~0~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~5~0002446~0002450~~~0002446~0002450~5~0~0~~0~~5~002446-002450~~~0~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~30~0002471~0002500~~~0002471~0002500~30~0~0~~0~~30~002471-002500~~~0~"
str2 = str2 & "0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~12~0000089~0000100~~~0000089~0000100~12~0~0~~0~~12~000089-000100~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~5~0000196~0000200~~~0000196~0000200~5~0~0~~0~~5~000196-000200~~~0"
str2 = str2 & "~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~11~0000340~0000350~~~0000340~0000350~11~0~0~~0~~11~000340-000350~~"
str2 = str2 & "~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~25~0000376~0000400~~~0000376~0000400~25~0~0~~0~~25~00"
str2 = str2 & ""
str2 = str2 & "0376-000400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~15~0000436~0000450~~~0000436~0000450~1"
str2 = str2 & "5~0~0~~0~~15~000436-000450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~11~0000490~0000500~~~0000490~0000500"
str2 = str2 & "~11~0~0~~0~~11~000490-000500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~2~0003799~0003800~~~0003799~000380"
str2 = str2 & "0~2~0~0~~0~~2~003799-003800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~8~0003893~0003900~~~0003893~0003900"
str2 = str2 & "~8~0~0~~0~~8~003893-003900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~48~0003903~0003950~~~0003903~0003950"
str2 = str2 & "~48~0~0~~0~~48~003903-003950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~21~0003980~0004000~~~0003980~00040"
str2 = str2 & "00~21~0~0~~0~~21~003980-004000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0004751~0004800~~~0004751~000"
str2 = str2 & "4800~50~0~0~~0~~50~004751-004800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~450~0004801~0005250~~~0004801~"
str2 = str2 & "0005250~450~0~0~~0~~450~004801-5250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~39~0005612~0005650~~~000561"
str2 = str2 & "2~0005650~39~0~0~~0~~39~005612-5650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~500~0005751~0006250~~~00057"
str2 = str2 & "51~0006250~500~0~0~~0~~500~005751-006250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~48~0006403~0006450~~~0"
str2 = str2 & "006403~0006450~48~0~0~~0~~48~006403-006450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~37~0006464~0006500~~"
str2 = str2 & "~0006464~0006500~37~0~0~~0~~37~006464-006500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~43~0006508~0006550"
str2 = str2 & "~~~0006508~0006550~43~0~0~~0~~43~006508-006550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~6~0006595~000660"
str2 = str2 & "0~~~0006595~0006600~6~0~0~~0~~6~006595-006600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~30~0006621~000665"
str2 = str2 & "0~~~0006621~0006650~30~0~0~~0~~30~006621-006650~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~48~0"
str2 = str2 & ""
str2 = str2 & "006703~0006750~~~0006703~0006750~48~0~0~~0~~48~006703-006750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-0"
str2 = str2 & "7~XQ/2010T~23~0006778~0006800~~~0006778~0006800~23~0~0~~0~~23~006778-006800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL"
str2 = str2 & "-07~XQ/2010T~46~0006805~0006850~~~0006805~0006850~46~0~0~~0~~46~006805-006850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2"
str2 = str2 & "LL-07~XQ/2010T~9~0006892~0006900~~~0006892~0006900~9~0~0~~0~~9~006892-6900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-"
str2 = str2 & "07~XQ/2010T~8~0006943~0006950~~~0006943~0006950~8~0~0~~0~~8~006943-006950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-0"
str2 = str2 & "7~XQ/2010T~17~0007034~0007050~~~0007034~0007050~17~0~0~~0~~17~007034-007050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL"
str2 = str2 & "-07~XQ/2010T~42~0007059~0007100~~~0007059~0007100~42~0~0~~0~~42~007059-007100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2"
str2 = str2 & "LL-07~XQ/2010T~29~0007122~0007150~~~0007122~0007150~29~0~0~~0~~29~007122-7150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2"
str2 = str2 & "LL-07~XQ/2010T~50~0007151~0007200~~~0007151~0007200~50~0~0~~0~~50~007151-007200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT"
str2 = str2 & "-2LL-07~XQ/2010T~150~0007201~0007350~~~0007201~0007350~150~0~0~~0~~150~007201-007350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~0"
str2 = str2 & "1GTKT-2LL-07~XQ/2010T~37~0007364~0007400~~~0007364~0007400~37~0~0~~0~~37~007364-7400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~0"
str2 = str2 & "1GTKT-2LL-07~XQ/2010T~20~0007431~0007450~~~0007431~0007450~20~0~0~~0~~20~007431-7450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~0"
str2 = str2 & "1GTKT-2LL-07~XQ/2010T~150~0007451~0007600~~~0007451~0007600~150~0~0~~0~~150~007451-007600~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD0"
str2 = str2 & "2LL~01GTKT-2LL-07~XQ/2010T~42~0007659~0007700~~~0007659~0007700~42~0~0~~0~~42~007659-007700~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D02LL~01GTKT-2LL-07~XQ/2010T~47~0007854~0007900~~~0007854~0007900~47~0~0~~0~~47~007854-007900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD02LL~01GTKT-2LL-07~XQ/2010T~30~0007921~0007950~~~0007921~0007950~30~0~0~~0~~30~007921-007950~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & ""
str2 = str2 & " trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0007951~0008000~~~0007951~0008000~50~0~0~~0~~50~007951-008000~~~"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0008001~0008050~~~0008001~0008050~50~0~0~~0~~50~008001-008050~"
str2 = str2 & "~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~28~0008073~0008100~~~0008073~0008100~28~0~0~~0~~28~008073-00810"
str2 = str2 & "0~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~249~0008102~0008350~~~0008102~0008350~249~0~0~~0~~249~008102-"
str2 = str2 & "8350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~6~0008495~0008500~~~0008495~0008500~6~0~0~~0~~6~008495-850"
str2 = str2 & "0~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~36~0009065~0009100~~~0009065~0009100~36~0~0~~0~~36~009065-009"
str2 = str2 & "100~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~26~0009125~0009150~~~0009125~0009150~26~0~0~~0~~26~009125-0"
str2 = str2 & "09150~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~39~0009162~0009200~~~0009162~0009200~39~0~0~~0~~39~009162"
str2 = str2 & "-009200~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~42~0009209~0009250~~~0009209~0009250~42~0~0~~0~~42~0092"
str2 = str2 & "09-009250~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~27~0009274~0009300~~~0009274~0009300~27~0~0~~0~~27~00"
str2 = str2 & "9274-009300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~47~0009304~0009350~~~0009304~0009350~47~0~0~~0~~47~"
str2 = str2 & "009304-009350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~450~0009351~0009800~~~0009351~0009800~450~0~0~~0~"
str2 = str2 & "~450~009351-009800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~3~0009848~0009850~~~0009848~0009850~3~0~0~~0"
str2 = str2 & "~~3~009848-9850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0009851~0009900~~~0009851~0009900~50~0~0~~0~"
str2 = str2 & "~50~009851-9900~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0009901~0009950~~~0009901~0009950~50~0~0~~0~"
str2 = str2 & "~50~009901-9950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0009951~0010000~~~0009951~0010000"
str2 = str2 & ""
str2 = str2 & "~50~0~0~~0~~50~009951-10000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0010001~0010050~~~0"
str2 = str2 & "010001~0010050~50~0~0~~0~~50~010001-10050~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~500~0010051~0010550~~"
str2 = str2 & "~0010051~0010550~500~0~0~~0~~500~010051-010550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~4450~0010551~001"
str2 = str2 & "5000~~~0010551~0015000~4450~0~0~~0~~4450~010551-15000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~32~000276"
str2 = str2 & "9~0002800~~~0002769~0002800~32~0~0~~0~~32~002769-02800~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~30~00029"
str2 = str2 & "21~0002950~~~0002921~0002950~30~0~0~~0~~30~002921-002950~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~25~000"
str2 = str2 & "2976~0003000~~~0002976~0003000~25~0~0~~0~~25~002976-003000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~36~0"
str2 = str2 & "004265~0004300~~~0004265~0004300~36~0~0~~0~~36~004265-004300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~26"
str2 = str2 & "~0001525~0001550~~~0001525~0001550~26~0~0~~0~~26~001525-001550~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~"
str2 = str2 & "20~0001831~0001850~~~0001831~0001850~20~0~0~~0~~20~001831-001850~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010"
str2 = str2 & "T~26~0001975~0002000~~~0001975~0002000~26~0~0~~0~~26~001975-002000~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/20"
str2 = str2 & "10T~44~0003257~0003300~~~0003257~0003300~44~0~0~~0~~44~003257-003300~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/"
str2 = str2 & "2010T~39~0003362~0003400~~~0003362~0003400~39~0~0~~0~~39~003362-003400~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~X"
str2 = str2 & "Q/2010T~29~0003472~0003500~~~0003472~0003500~29~0~0~~0~~29~003472-003500~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07"
str2 = str2 & "~XQ/2010T~21~0003430~0003450~~~0003430~0003450~21~0~0~~0~~21~003430-003450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-"
str2 = str2 & "07~XQ/2010T~40~0038711~0038750~~~0038711~0038750~40~0~0~~0~~40~038711-038750~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02L"
str2 = str2 & ""
str2 = str2 & "L~01GTKT-2LL-07~XQ/2010T~34~0002567~0002600~~~0002567~0002600~34~0~0~~0~~34~002567-002600~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~48~0002603~0002650~~~0002603~0002650~48~0~0~~0~~48~002603-002650~~~0~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~48~0002653~0002700~~~0002653~0002700~48~0~0~~0~~48~002653-002700~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0002701~0002750~~~0002701~0002750~50~0~0~~0~~50~002701-002750~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~11~0000590~0000600~~~0000590~0000600~11~0~0~~0~~11~000590-00600~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~8~0000943~0000950~~~0000943~0000950~8~0~0~~0~~8~000943-000950~~~0~0~H„a Æ¨n g"
str2 = str2 & "i∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~25~0001026~0001050~~~0001026~0001050~25~0~0~~0~~25~001026-001050~~~0~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~32~0001119~0001150~~~0001119~0001150~32~0~0~~0~~32~001119-001150~~~0~0~H„a Æ"
str2 = str2 & "¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~35~0001316~0001350~~~0001316~0001350~35~0~0~~0~~35~001316-001350~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~9~0001392~0001400~~~0001392~0001400~9~0~0~~0~~9~001392-001400~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~42~0001409~0001450~~~0001409~0001450~42~0~0~~0~~42~001409-001450~~~0~0~H„"
str2 = str2 & "a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~17~0001484~0001500~~~0001484~0001500~17~0~0~~0~~17~001484-001500~~~0~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~7~0003144~0003150~~~0003144~0003150~7~0~0~~0~~7~003144-003150~~~0~0~H"
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~46~0003155~0003200~~~0003155~0003200~46~0~0~~0~~46~003155-003200~~~0~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~50~0003201~0003250~~~0003201~0003250~50~0~0~~0~~50~003201-3250~~~0~0"
str2 = str2 & "~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~39~0004312~0004350~~~0004312~0004350~39~0~0~~0~~39~004312"
str2 = str2 & ""
str2 = str2 & "-004350~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT-2LL-07~XQ/2010T~32~0004419~0004450~~~0004419~0004450~32~0~"
str2 = str2 & "0~~0~~32~004419-4450~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD02LL~01GTKT2/001BD~AA/11P~50~0000001~0000050~~~~~0~0~0~~0~~0~~0000001~"
str2 = str2 & "0000050~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000051~0000100~~~~~0~0~0~~0~~0~~0000051~0000100~50~0~H"
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000101~0000150~~~0000101~0000150~50~47~3~107;127;130~0~~0~~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000151~0000200~~~0000151~0000200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000201~0000250~~~0000201~0000233~33~30~3~203;227;231~0~~0~~0000234~0000250~17~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000251~0000300~~~0000251~0000300~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000301~0000350~~~0000301~0000350~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD01"
str2 = str2 & "2L~01GTKT2/001BD~AA/11P~50~0000351~0000400~~~0000351~0000400~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/00"
str2 = str2 & "1BD~AA/11P~50~0000401~0000450~~~0000401~0000450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50"
str2 = str2 & "~0000451~0000500~~~0000451~0000483~33~32~1~478~0~~0~~0000484~0000500~17~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/1"
str2 = str2 & "1P~50~0000501~0000550~~~0000501~0000514~14~14~0~~0~~0~~0000515~0000550~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA"
str2 = str2 & "/11P~50~0000551~0000600~~~0000551~0000600~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~00006"
str2 = str2 & "01~0000650~~~0000601~0000650~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000651~0000700~~~"
str2 = str2 & "0000651~0000657~7~7~0~~0~~0~~0000658~0000700~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0000701~0000750~~~"
str2 = str2 & "0000701~0000719~19~19~0~~0~~0~~0000720~0000750~31~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~00007"
str2 = str2 & ""
str2 = str2 & "51~0000800~~~0000751~0000788~38~38~0~~0~~0~~0000789~0000800~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~"
str2 = str2 & "AA/11P~50~0000801~0000850~~~0000801~0000850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000"
str2 = str2 & "0851~0000900~~~0000851~0000853~3~3~0~~0~~0~~0000854~0000900~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000"
str2 = str2 & "0901~0000950~~~0000901~0000938~38~38~0~~0~~0~~0000939~0000950~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0"
str2 = str2 & "000951~0001000~~~0000951~0000988~38~38~0~~0~~0~~0000989~0001000~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50"
str2 = str2 & "~0001001~0001050~~~0001001~0001050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001051~0001"
str2 = str2 & "100~~~0001051~0001100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001101~0001150~~~0001101"
str2 = str2 & "~0001114~14~11~3~1104;1109;1112~0~~0~~0001115~0001150~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001151~0"
str2 = str2 & "001200~~~0001151~0001200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001201~0001250~~~0001"
str2 = str2 & "201~0001250~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001251~0001300~~~0001251~0001272~2"
str2 = str2 & "2~22~0~~0~~0~~0001273~0001300~28~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001301~0001350~~~0001301~0001323"
str2 = str2 & "~23~23~0~~0~~0~~0001324~0001350~27~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001351~0001400~~~0001351~00013"
str2 = str2 & "51~1~1~0~~0~~0~~0001352~0001400~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001401~0001450~~~0001401~00014"
str2 = str2 & "38~38~38~0~~0~~0~~0001439~0001450~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001451~0001500~~~0001451~000"
str2 = str2 & "1488~38~38~0~~0~~0~~0001489~0001500~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001501~0001550~~~0001501~0"
str2 = str2 & "001550~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001551~0001600~~~0001551~000"
str2 = str2 & ""
str2 = str2 & "1600~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001601~0001650~~~0001601~000"
str2 = str2 & "1650~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001651~0001700~~~0001651~0001700~50~50~0~"
str2 = str2 & "~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001701~0001750~~~0001701~0001750~50~50~0~~0~~0~~~~0~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001751~0001800~~~0001751~0001800~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0001801~0001850~~~0001801~0001850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-B"
str2 = str2 & "D012L~01GTKT2/001BD~AA/11P~50~0001851~0001900~~~0001851~0001883~33~33~0~~0~~0~~0001884~0001900~17~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD012L~01GTKT2/001BD~AA/11P~50~0001901~0001950~~~0001901~0001950~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTK"
str2 = str2 & "T2/001BD~AA/11P~50~0001951~0002000~~~0001951~0002000~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/1"
str2 = str2 & "1P~50~0002001~0002050~~~0002001~0002050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002051"
str2 = str2 & "~0002100~~~0002051~0002088~38~38~0~~0~~0~~0002089~0002100~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~00021"
str2 = str2 & "01~0002150~~~0002101~0002117~17~17~0~~0~~0~~0002118~0002150~33~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000"
str2 = str2 & "2151~0002200~~~0002151~0002159~9~9~0~~0~~0~~0002160~0002200~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000"
str2 = str2 & "2201~0002250~~~0002201~0002238~38~38~0~~0~~0~~0002239~0002250~12~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0"
str2 = str2 & "002251~0002300~~~0002251~0002263~13~13~0~~0~~0~~0002264~0002300~37~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50"
str2 = str2 & "~0002301~0002350~~~~~0~0~0~~0~~0~~0002301~0002350~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002351~00024"
str2 = str2 & "00~~~0002351~0002400~50~49~1~0002393~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002401~"
str2 = str2 & ""
str2 = str2 & "0002450~~~0002401~0002450~50~49~1~0002411~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0"
str2 = str2 & "002451~0002500~~~0002451~0002477~27~27~0~~0~~0~~0002478~0002500~23~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50"
str2 = str2 & "~0002501~0002550~~~0002501~0002528~28~28~0~~0~~0~~0002529~0002550~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~"
str2 = str2 & "50~0002551~0002600~~~0002551~0002560~10~10~0~~0~~0~~0002561~0002600~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11"
str2 = str2 & "P~50~0002601~0002650~~~0002601~0002650~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002651~"
str2 = str2 & "0002700~~~0002651~0002700~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002701~0002750~~~000"
str2 = str2 & "2701~0002709~9~9~0~~0~~0~~0002710~0002750~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002751~0002800~~~000"
str2 = str2 & "2751~0002755~5~4~1~0002751~0~~0~~0002756~0002800~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002801~000285"
str2 = str2 & "0~~~0002801~0002850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002851~0002900~~~0002851~0"
str2 = str2 & "002853~3~3~0~~0~~0~~0002854~0002900~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002901~0002950~~~~~0~0~0~~"
str2 = str2 & "0~~0~~0002901~0002950~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0002951~0003000~~~0002951~0002961~11~11~0"
str2 = str2 & "~~0~~0~~0002962~0003000~39~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003001~0003050~~~0003001~0003003~3~3~0"
str2 = str2 & "~~0~~0~~0003004~0003050~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003051~0003100~~~0003051~0003066~16~16"
str2 = str2 & "~0~~0~~0~~0003067~0003100~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003101~0003150~~~~~0~0~0~~0~~0~~0003"
str2 = str2 & "101~0003150~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003151~0003200~~~0003151~0003151~1~1~0~~0~~0~~0003"
str2 = str2 & "152~0003200~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003201~0003250~~~0003201~0003207~7~7~0~"
str2 = str2 & ""
str2 = str2 & "~0~~0~~0003208~0003250~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003251~0003300~~~0003251~0"
str2 = str2 & "003268~18~18~0~~0~~0~~0003269~0003300~32~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003301~0003350~~~0003301"
str2 = str2 & "~0003304~4~4~0~~0~~0~~0003305~0003350~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003351~0003400~~~0003351"
str2 = str2 & "~0003375~25~25~0~~0~~0~~0003376~0003400~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003401~0003450~~~00034"
str2 = str2 & "01~0003402~2~2~0~~0~~0~~0003403~0003450~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003451~0003500~~~00034"
str2 = str2 & "51~0003471~21~21~0~~0~~0~~0003472~0003500~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003501~0003550~~~000"
str2 = str2 & "3501~0003501~1~1~0~~0~~0~~0003502~0003550~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003551~0003600~~~000"
str2 = str2 & "3551~0003558~8~8~0~~0~~0~~0003559~0003600~42~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003601~0003650~~~000"
str2 = str2 & "3601~0003603~3~3~0~~0~~0~~0003604~0003650~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003651~0003700~~~000"
str2 = str2 & "3651~0003658~8~8~0~~0~~0~~0003659~0003700~42~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003701~0003750~~~000"
str2 = str2 & "3701~0003703~3~3~0~~0~~0~~0003704~0003750~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003751~0003800~~~000"
str2 = str2 & "3751~0003754~4~4~0~~0~~0~~0003755~0003800~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003801~0003850~~~000"
str2 = str2 & "3801~0003802~2~2~0~~0~~0~~0003803~0003850~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003851~0003900~~~000"
str2 = str2 & "3851~0003900~50~47~3~0003863;0003866"
str2 = str2 & ";0003872~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0003901~00"
str2 = str2 & "03950~~~0003901~0003950~50~47~3~0003912;0003917;"
str2 = str2 & "0003943~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50"
str2 = str2 & "~0003951~0004000~~~0003951~0004000~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0"
str2 = str2 & ""
str2 = str2 & "004001~0004050~~~0004001~0004019~19~19~0~~0~~0~~0004020~0004050~31~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/00"
str2 = str2 & "1BD~AA/11P~50~0004051~0004100~~~0004051~0004100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50"
str2 = str2 & "~0004101~0004150~~~0004101~0004150~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004151~0004"
str2 = str2 & "200~~~0004151~0004178~28~28~0~~0~~0~~0004179~0004200~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004201~00"
str2 = str2 & "04250~~~0004201~0004201~1~1~0~~0~~0~~0004202~0004250~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004251~00"
str2 = str2 & "04300~~~0004251~0004279~29~29~0~~0~~0~~0004280~0004300~21~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004301~"
str2 = str2 & "0004350~~~0004301~0004332~32~32~0~~0~~0~~0004333~0004350~18~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000435"
str2 = str2 & "1~0004400~~~0004351~0004353~3~3~0~~0~~0~~0004354~0004400~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000440"
str2 = str2 & "1~0004450~~~0004401~0004407~7~7~0~~0~~0~~0004408~0004450~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000445"
str2 = str2 & "1~0004500~~~0004451~0004456~6~6~0~~0~~0~~0004457~0004500~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000450"
str2 = str2 & "1~0004550~~~0004501~0004506~6~6~0~~0~~0~~0004507~0004550~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000455"
str2 = str2 & "1~0004600~~~0004551~0004559~9~9~0~~0~~0~~0004560~0004600~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000460"
str2 = str2 & "1~0004650~~~0004601~0004611~11~11~0~~0~~0~~0004612~0004650~39~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004"
str2 = str2 & "651~0004700~~~0004651~0004655~5~5~0~~0~~0~~0004656~0004700~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004"
str2 = str2 & "701~0004750~~~0004701~0004703~3~3~0~~0~~0~~0004704~0004750~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0004"
str2 = str2 & "751~0004800~~~0004751~0004800~50~46~4~4758;4768;"
str2 = str2 & "4774;4798~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/00"
str2 = str2 & ""
str2 = str2 & "1BD~AA/11P~50~0004801~0004850~~~0004801~0004804~4~4~0~~0~~0~~0004805~0004850~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD01"
str2 = str2 & "2L~01GTKT2/001BD~AA/11P~50~0004851~0004900~~~0004851~0004854~4~4~0~~0~~0~~0004855~0004900~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD01"
str2 = str2 & "2L~01GTKT2/001BD~AA/11P~50~0004901~0004950~~~0004901~0004902~2~2~0~~0~~0~~0004903~0004950~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD01"
str2 = str2 & "2L~01GTKT2/001BD~AA/11P~50~0004951~0005000~~~0004951~0004953~3~3~0~~0~~0~~0004954~0005000~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD01"
str2 = str2 & "2L~01GTKT2/001BD~AA/11P~50~0005001~0005050~~~0005001~0005050~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/00"
str2 = str2 & "1BD~AA/11P~50~0005051~0005100~~~0005051~0005078~28~27~1~005071~0~~0~~0005079~0005100~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01"
str2 = str2 & "GTKT2/001BD~AA/11P~50~0005101~0005150~~~0005101~0005150~50~48~2~5121;5126~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~50~0005151~0005200~~~0005151~0005199~49~43~6~5157;5158;5182;5185;5194;5198~0~~0~~0005200~0005200~1~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005201~0005250~~~0005201~0005227~27~27~0~~0~~0~~0005228~0005250~23~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005251~0005300~~~0005251~0005255~5~5~0~~0~~0~~0005256~0005300~45~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005301~0005350~~~0005301~0005305~5~5~0~~0~~0~~0005306~0005350~45~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005351~0005400~~~0005351~0005364~14~14~0~~0~~0~~0005365~0005400~36~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005401~0005450~~~0005401~0005408~8~8~0~~0~~0~~0005409~0005450~42~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005451~0005500~~~0005451~0005453~3~3~0~~0~~0~~0005454~0005500~47~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005501~0005550~~~0005501~0005503~3~3~0~~0~~0~~0005504~0005550~47~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005551~0005600~~~0005551~0005551~1~1~0~~0~~0~~0005552~000560"
str2 = str2 & ""
str2 = str2 & "0~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005601~0005650~~~0005601~0005639~39~39~0~~0~~0~"
str2 = str2 & "~0005640~0005650~11~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005651~0005700~~~0005651~0005653~3~3~0~~0~~0~"
str2 = str2 & "~0005654~0005700~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005701~0005750~~~0005701~0005703~3~3~0~~0~~0~"
str2 = str2 & "~0005704~0005750~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005751~0005800~~~0005751~0005756~6~6~0~~0~~0~"
str2 = str2 & "~0005757~0005800~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005801~0005850~~~0005801~0005803~3~3~0~~0~~0~"
str2 = str2 & "~0005804~0005850~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005851~0005900~~~0005851~0005900~50~50~0~~0~~"
str2 = str2 & "0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005901~0005950~~~0005901~0005920~20~19~1~0005919~0~~0~~000"
str2 = str2 & "5921~0005950~30~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0005951~0006000~~~0005951~0005983~33~32~1~0005951~"
str2 = str2 & "0~~0~~0005984~0006000~17~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006001~0006050~~~0006001~0006050~50~49~1"
str2 = str2 & "~0006004~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006051~0006100~~~0006051~0006055~5~5~0~~0~~0~~"
str2 = str2 & "0006056~0006100~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006101~0006150~~~0006101~0006112~12~12~0~~0~~0"
str2 = str2 & "~~0006113~0006150~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006151~0006200~~~0006151~0006156~6~6~0~~0~~0"
str2 = str2 & "~~0006157~0006200~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006201~0006250~~~0006201~0006215~15~15~0~~0~"
str2 = str2 & "~0~~0006216~0006250~35~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006251~0006300~~~0006251~0006262~12~12~0~~"
str2 = str2 & "0~~0~~0006263~0006300~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006301~0006350~~~0006301~0006317~17~17~0"
str2 = str2 & "~~0~~0~~0006318~0006350~33~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006351~0006400~~~0006351~00"
str2 = str2 & ""
str2 = str2 & "06364~14~14~0~~0~~0~~0006365~0006400~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006401~00064"
str2 = str2 & "50~~~0006401~0006450~50~49~1~0006446~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006451~0006500~~~0"
str2 = str2 & "006451~0006462~12~12~0~~0~~0~~0006463~0006500~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006501~0006550~~"
str2 = str2 & "~0006501~0006503~3~3~0~~0~~0~~0006504~0006550~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006551~0006600~~"
str2 = str2 & "~0006551~0006565~15~13~2~0006551;0006552~0~~0~~0006566~0006600~35~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~"
str2 = str2 & "0006601~0006650~~~0006601~0006650~50~46~4~6627;6631"
str2 = str2 & ";6641;6614~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/"
str2 = str2 & "11P~50~0006651~0006700~~~0006651~0006700~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000670"
str2 = str2 & "1~0006750~~~0006701~0006735~35~35~0~~0~~0~~0006736~0006750~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006"
str2 = str2 & "751~0006800~~~0006751~0006756~6~6~0~~0~~0~~0006757~0006800~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006"
str2 = str2 & "801~0006850~~~0006801~0006807~7~7~0~~0~~0~~0006808~0006850~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006"
str2 = str2 & "851~0006900~~~0006851~0006856~6~6~0~~0~~0~~0006857~0006900~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0006"
str2 = str2 & "901~0006950~~~0006901~0006947~47~47~0~~0~~0~~0006948~0006950~3~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000"
str2 = str2 & "6951~0007000~~~0006951~0006964~14~13~1~0006955~0~~0~~0006965~0007000~36~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/1"
str2 = str2 & "1P~100~0007001~0007100~~~~~0~0~0~~0~~0~~0007001~0007100~100~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~000710"
str2 = str2 & "1~0007150~~~0007101~0007150~50~49~1~0007111~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007151~0007"
str2 = str2 & "200~~~0007151~0007200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007201~000725"
str2 = str2 & ""
str2 = str2 & "0~~~0007201~0007246~46~45~1~0007222~0~~0~~0007247~0007250~4~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/"
str2 = str2 & "11P~50~0007251~0007300~~~0007251~0007272~22~20~2~7265;7266~0~~0~~0007273~0007300~28~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~50~0007301~0007350~~~0007301~0007335~35~33~2~7301;7335~0~~0~~0007336~0007350~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "012L~01GTKT2/001BD~AA/11P~50~0007351~0007400~~~0007351~0007375~25~25~0~~0~~0~~0007376~0007400~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng-"
str2 = str2 & "BD012L~01GTKT2/001BD~AA/11P~50~0007401~0007450~~~0007401~0007418~18~17~1~7416~0~~0~~0007419~0007450~32~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007451~0007500~~~0007451~0007467~17~17~0~~0~~0~~0007468~0007500~33~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007501~0007550~~~0007501~0007512~12~11~1~0007504~0~~0~~0007513~0007550~38~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007551~0007600~~~0007551~0007570~20~20~0~~0~~0~~0007571~0007600~30~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007601~0007650~~~0007601~0007609~9~9~0~~0~~0~~0007610~0007650~41~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007651~0007700~~~0007651~0007663~13~13~0~~0~~0~~0007664~0007700~37~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007701~0007750~~~0007701~0007707~7~7~0~~0~~0~~0007708~0007750~43~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007751~0007800~~~0007751~0007752~2~2~0~~0~~0~~0007753~0007800~48~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007801~0007850~~~0007801~0007850~50~49~1~0007803~0~~0~~~~0~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007851~0007900~~~0007851~0007900~50~49~1~7859~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0007901~0007950~~~0007901~0007950~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~"
str2 = str2 & "01GTKT2/001BD~AA/11P~50~0007951~0008000~~~~~0~0~0~~0~~0~~0007951~0008000~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01G"
str2 = str2 & ""
str2 = str2 & "TKT2/001BD~AA/11P~50~0008001~0008050~~~0008001~0008004~4~3~1~8002~0~~0~~0008005~0008050~46~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0008051~0008100~~~0008051~0008052~2~1~1~8051~0~~0~~0008053~0008100~48~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0008101~0008150~~~0008101~0008150~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "012L~01GTKT2/001BD~AA/11P~50~0008151~0008200~~~0008151~0008200~50~49~1~0008174~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~0"
str2 = str2 & "1GTKT2/001BD~AA/11P~50~0008201~0008250~~~0008201~0008250~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~"
str2 = str2 & "AA/11P~50~0008251~0008300~~~0008251~0008256~6~6~0~~0~~0~~0008257~0008300~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~"
str2 = str2 & "AA/11P~50~0008301~0008350~~~0008301~0008303~3~3~0~~0~~0~~0008304~0008350~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~"
str2 = str2 & "AA/11P~50~0008351~0008400~~~0008351~0008370~20~20~0~~0~~0~~0008371~0008400~30~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001B"
str2 = str2 & "D~AA/11P~50~0008401~0008450~~~0008401~0008416~16~16~0~~0~~0~~0008417~0008450~34~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/00"
str2 = str2 & "1BD~AA/11P~50~0008451~0008500~~~0008451~0008461~11~10~1~0008451~0~~0~~0008462~0008500~39~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~0"
str2 = str2 & "1GTKT2/001BD~AA/11P~50~0008501~0008550~~~0008501~0008505~5~5~0~~0~~0~~0008506~0008550~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~0"
str2 = str2 & "1GTKT2/001BD~AA/11P~50~0008551~0008600~~~0008551~0008562~12~12~0~~0~~0~~0008563~0008600~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L"
str2 = str2 & "~01GTKT2/001BD~AA/11P~50~0008601~0008650~~~0008601~0008606~6~6~0~~0~~0~~0008607~0008650~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L"
str2 = str2 & "~01GTKT2/001BD~AA/11P~50~0008651~0008700~~~0008651~0008653~3~3~0~~0~~0~~0008654~0008700~47~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L"
str2 = str2 & "~01GTKT2/001BD~AA/11P~50~0008701~0008750~~~0008701~0008709~9~9~0~~0~~0~~0008710~0008750~41~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L"
str2 = str2 & "~01GTKT2/001BD~AA/11P~50~0008751~0008800~~~0008751~0008782~32~32~0~~0~~0~~0008783~0008800~18~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & ""
str2 = str2 & "a t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0008801~0008850~~~0008801~0008850~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0008851~0008900~~~0008851~0008900~50~46~4~8856;8857;8865;8895~0~~0~~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0008901~0008950~~~0008901~0008930~30~30~0~~0~~0~~0008931~0008950~20~0~H„a Æ¨n gi"
str2 = str2 & "∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0008951~0009000~~~~~0~0~0~~0~~0~~0008951~0009000~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD012L~01GTKT2/001BD~AA/11P~100~0009001~0009100~~~0009001~0009100~100~100~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01"
str2 = str2 & "GTKT2/001BD~AA/11P~100~0009101~0009200~~~0009101~0009200~100~95~5~09125;09128;09129;09141;09180~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009201~0009250~~~0009201~0009231~31~29~2~09220;09221~0~~0~~0009232~0009250~19~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009251~0009300~~~0009251~0009300~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009301~0009350~~~0009301~0009319~19~19~0~~0~~0~~0009320~0009350~31~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009351~0009400~~~0009351~0009366~16~15~1~09357~0~~0~~0009367~0009400~34~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009401~0009450~~~0009401~0009450~50~49~1~09415~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009451~0009500~~~0009451~0009470~20~20~0~~0~~0~~0009471~0009500~30~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009501~0009550~~~0009501~0009522~22~19~3~09509;09510;09521~0~~0~~0009523~0009550"
str2 = str2 & "~28~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009551~0009600~~~0009551~0009569~19~19~0~~0~~0~~0009570~00096"
str2 = str2 & "00~31~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009601~0009650~~~0009601~0009619~19~18~1~09613~0~~0~~000962"
str2 = str2 & "0~0009650~31~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009651~0009700~~~0009651~0009700~50~49~1~"
str2 = str2 & ""
str2 = str2 & "09696~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009701~0009750~~~0009701~0009720~20~"
str2 = str2 & "20~0~~0~~0~~0009721~0009750~30~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009751~0009800~~~0009751~0009800~5"
str2 = str2 & "0~46~4~09758;09772;09787;09795~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009801~0009850~~~0009801"
str2 = str2 & "~0009817~17~16~1~09812~0~~0~~0009818~0009850~33~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~100~0009851~0009950~~"
str2 = str2 & "~0009851~0009950~100~99~1~09886~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0009951~0010000~~~000995"
str2 = str2 & "1~0009987~37~36~1~09967~0~~0~~0009988~0010000~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~200~0010001~0010200~"
str2 = str2 & "~~0010001~0010200~200~192~8~10007;10014;10026;10028;10047;10079;10166;10198~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GT"
str2 = str2 & "KT2/001BD~AA/11P~50~0010201~0010250~~~0010201~0010248~48~47~1~10226~0~~0~~0010249~0010250~2~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012"
str2 = str2 & "L~01GTKT2/001BD~AA/11P~50~0010251~0010300~~~0010251~0010290~40~39~1~10260~0~~0~~0010291~0010300~10~0~H„a Æ¨n gi∏ trﬁ gia t®n"
str2 = str2 & "g-BD012L~01GTKT2/001BD~AA/11P~50~0010301~0010350~~~0010301~0010312~12~12~0~~0~~0~~0010313~0010350~38~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng-BD012L~01GTKT2/001BD~AA/11P~50~0010351~0010400~~~0010351~0010359~9~9~0~~0~~0~~0010360~0010400~41~0~H„a Æ¨n gi∏ trﬁ gia t"
str2 = str2 & "®ng-BD012L~01GTKT2/001BD~AA/11P~200~0010401~0010600~~~~~0~0~0~~0~~0~~0010401~0010600~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~0"
str2 = str2 & "1GTKT2/001BD~AA/11P~50~0010601~0010650~~~0010601~0010650~50~49~1~0010627~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2"
str2 = str2 & "/001BD~AA/11P~50~0010651~0010700~~~0010651~0010700~50~49~1~0010692~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD"
str2 = str2 & "~AA/11P~50~0010701~0010750~~~0010701~0010750~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~00"
str2 = str2 & "10751~0010800~~~0010751~0010800~50~48~2~0010785;0010786~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD"
str2 = str2 & ""
str2 = str2 & "~AA/11P~50~0010801~0010850~~~0010801~0010850~50~49~1~0010846~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~50~0010851~0010900~~~0010851~0010900~50~45~5~0010888;0010889;0010890;0010891;0010893~0~~0~~~~0~0~H„a Æ¨n gi∏"
str2 = str2 & "trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0010901~0010950~~~0010901~0010950~50~49~1~0010922~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0010951~0011000~~~0010951~0011000~50~49~1~0010991~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng"
str2 = str2 & "-BD012L~01GTKT2/001BD~AA/11P~50~0011001~0011050~~~0011001~0011050~50~49~1~0011004~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012"
str2 = str2 & "L~01GTKT2/001BD~AA/11P~50~0011051~0011100~~~0011051~0011086~36~29~7~0011053;0011057;0011060"
str2 = str2 & ";0011061;0011062;0011064;0011068"
str2 = str2 & "~0~~0~~0011087~0011100~14~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011101~0011150~~~0011101~0011110~10~9~1"
str2 = str2 & "~0011103~0~~0~~0011111~0011150~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011151~0011200~~~0011151~001115"
str2 = str2 & "4~4~4~0~~0~~0~~0011155~0011200~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011201~0011250~~~0011201~001125"
str2 = str2 & "0~50~48~2~0011225;0011242~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011251~0011300~~~0011251~0011"
str2 = str2 & "300~50~49~1~0011261~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011301~0011350~~~0011301~0011341~41"
str2 = str2 & "~40~1~0011311~0~~0~~0011342~0011350~9~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011351~0011400~~~0011351~00"
str2 = str2 & "11384~34~33~1~11361~0~~0~~0011385~0011400~16~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011401~0011450~~~001"
str2 = str2 & "1401~0011450~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011451~0011500~~~0011451~0011474~"
str2 = str2 & "24~24~0~~0~~0~~0011475~0011500~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011501~0011550~~~~~0~0~0~~0~~0~"
str2 = str2 & "~0011501~0011550~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011551~0011600~~~0011551~0011578~2"
str2 = str2 & ""
str2 = str2 & "8~28~0~~0~~0~~0011579~0011600~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011601~0011650~~~00"
str2 = str2 & "11601~0011624~24~23~1~0011621~0~~0~~0011625~0011650~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011651~001"
str2 = str2 & "1700~~~0011651~0011671~21~21~0~~0~~0~~0011672~0011700~29~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011701~0"
str2 = str2 & "011750~~~0011701~0011726~26~26~0~~0~~0~~0011727~0011750~24~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011751"
str2 = str2 & "~0011800~~~0011751~0011760~10~10~0~~0~~0~~0011761~0011800~40~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~00118"
str2 = str2 & "01~0011850~~~0011801~0011828~28~28~0~~0~~0~~0011829~0011850~22~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~001"
str2 = str2 & "1851~0011900~~~0011851~0011874~24~23~1~0011853~0~~0~~0011875~0011900~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/1"
str2 = str2 & "1P~50~0011901~0011950~~~0011901~0011950~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0011951"
str2 = str2 & "~0012000~~~0011951~0011987~37~37~0~~0~~0~~0011988~0012000~13~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~00120"
str2 = str2 & "01~0012050~~~0012001~0012050~50~49~1~0012013~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012051~001"
str2 = str2 & "2100~~~0012051~0012073~23~23~0~~0~~0~~0012074~0012100~27~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012101~0"
str2 = str2 & "012150~~~0012101~0012146~46~46~0~~0~~0~~0012147~0012150~4~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012151~"
str2 = str2 & "0012200~~~0012151~0012200~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012201~0012250~~~001"
str2 = str2 & "2201~0012226~26~25~1~0012222~0~~0~~0012227~0012250~24~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~150~0012251~001"
str2 = str2 & "2400~~~0012251~0012400~150~140~10~12263;12277;12294;12318;12337;12366;12388;12349;12394;12395~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gi"
str2 = str2 & "a t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012401~0012450~~~0012401~0012450~50~48~2~12413;12417~0~~0~~~~0~0~H„a Æ¨n g"
str2 = str2 & ""
str2 = str2 & "i∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012451~0012500~~~0012451~0012485~35~35~0~~0~~0~~0012486~0012500"
str2 = str2 & "~15~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~150~0012501~0012650~~~~~0~0~0~~0~~0~~0012501~0012650~150~0~H„a Æ¨"
str2 = str2 & "n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012651~0012700~~~0012651~0012700~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012701~0012750~~~0012701~0012703~3~3~0~~0~~0~~0012704~0012750~47~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012751~0012800~~~0012751~0012783~33~33~0~~0~~0~~0012784~0012800~17~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012801~0012850~~~0012801~0012833~33~33~0~~0~~0~~0012834~0012850~17~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012851~0012900~~~0012851~0012856~6~6~0~~0~~0~~0012857~0012900~44~0~H„a Æ¨n gi∏ trﬁ"
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012901~0012950~~~0012901~0012920~20~20~0~~0~~0~~0012921~0012950~30~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0012951~0013000~~~0012951~0012953~3~3~0~~0~~0~~0012954~0013000~47~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013001~0013050~~~0013001~0013013~13~11~2~0013001;0013007~0~~0~~0013014~0013050~3"
str2 = str2 & "7~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013051~0013100~~~0013051~0013062~12~12~0~~0~~0~~0013063~0013100"
str2 = str2 & "~38~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013101~0013150~~~0013101~0013106~6~6~0~~0~~0~~0013107~0013150"
str2 = str2 & "~44~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013151~0013200~~~0013151~0013200~50~47~3~13173;13178;13189~0~"
str2 = str2 & "~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013201~0013250~~~0013201~0013250~50~50~0~~0~~0~~~~0~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013251~0013300~~~0013251~0013300~50~48~2~13280;13298~0~~0~~~~0~0~H„a Æ"
str2 = str2 & "¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013301~0013350~~~0013301~0013302~2~2~0~~0~~0~~0013303~0013350"
str2 = str2 & ""
str2 = str2 & "~48~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013351~0013400~~~0013351~0013400~50~50~0~~0~~0~~"
str2 = str2 & "~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013401~0013450~~~~~0~0~0~~0~~0~~0013401~0013450~50~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013451~0013500~~~0013451~0013472~22~21~1~13468~0~~0~~0013473~0013500~28~0~H"
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~150~0013501~0013650~~~~~0~0~0~~0~~0~~0013501~0013650~150~0~H„a Æ¨n gi∏ t"
str2 = str2 & "rﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013651~0013700~~~0013651~0013662~12~11~1~0013660~0~~0~~0013663~0013700~38~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013701~0013750~~~0013701~0013746~46~46~0~~0~~0~~0013747~0013750~4~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013751~0013800~~~0013751~0013780~30~30~0~~0~~0~~0013781~0013800~20~0~H"
str2 = str2 & "„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013801~0013850~~~0013801~0013841~41~41~0~~0~~0~~0013842~0013850~9~0~"
str2 = str2 & "H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013851~0013900~~~~~0~0~0~~0~~0~~0013851~0013900~50~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0013901~0013950~~~0013901~0013950~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "012L~01GTKT2/001BD~AA/11P~50~0013951~0014000~~~0013951~0013957~7~7~0~~0~~0~~0013958~0014000~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "012L~01GTKT2/001BD~AA/11P~400~0014001~0014400~~~~~0~0~0~~0~~0~~0014001~0014400~400~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2"
str2 = str2 & "/001BD~AA/11P~500~0014401~0014900~~~~~0~0~0~~0~~0~~0014401~0014900~500~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11"
str2 = str2 & "P~150~0014901~0015050~~~0014901~0015050~150~142~8~14935;14946;14979;14981;14992;15004;15006;15025~0~~0~~~~0~0~H„a Æ¨n gi∏ tr"
str2 = str2 & "ﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015051~0015100~~~0015051~0015100~50~50~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD"
str2 = str2 & "012L~01GTKT2/001BD~AA/11P~50~0015101~0015150~~~0015101~0015106~6~6~0~~0~~0~~0015107~0015150~44~0~H„a Æ¨n gi∏ trﬁ "
str2 = str2 & ""
str2 = str2 & "gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015151~0015200~~~0015151~0015153~3~3~0~~0~~0~~0015154~0015200~47~0~H„a"
str2 = str2 & "Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015201~0015250~~~~~0~0~0~~0~~0~~0015201~0015250~50~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015251~0015300~~~0015251~0015255~5~5~0~~0~~0~~0015256~0015300~45~0~H„a Æ¨n gi∏ trﬁ g"
str2 = str2 & "ia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015301~0015350~~~0015301~0015319~19~18~1~0015318~0~~0~~0015320~0015350~31~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015351~0015400~~~0015351~0015355~5~5~0~~0~~0~~0015356~0015400~45~0~H„a Æ¨n"
str2 = str2 & "gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~150~0015401~0015550~~~~~0~0~0~~0~~0~~0015401~0015550~150~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0015551~0015600~~~0015551~0015558~8~8~0~~0~~0~~0015559~0015600~42~0~H„a Æ¨n gi∏ trﬁ gia"
str2 = str2 & "t®ng-BD012L~01GTKT2/001BD~AA/11P~200~0015601~0015800~~~~~0~0~0~~0~~0~~0015601~0015800~200~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~"
str2 = str2 & "01GTKT2/001BD~AA/11P~50~0015801~0015850~~~0015801~0015850~50~49~1~0015845~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~50~0015851~0015900~~~0015851~0015855~5~5~0~~0~~0~~0015856~0015900~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~50~0015901~0015950~~~0015901~0015904~4~4~0~~0~~0~~0015905~0015950~46~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~50~0015951~0016000~~~0015951~0015951~1~1~0~~0~~0~~0015952~0016000~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT"
str2 = str2 & "2/001BD~AA/11P~300~0016001~0016300~~~~~0~0~0~~0~~0~~0016001~0016300~300~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/1"
str2 = str2 & "1P~250~0016301~0016550~~~~~0~0~0~~0~~0~~0016301~0016550~250~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~001655"
str2 = str2 & "1~0016600~~~0016551~0016561~11~11~0~~0~~0~~0016562~0016600~39~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0016"
str2 = str2 & "601~0016650~~~0016601~0016605~5~5~0~~0~~0~~0016606~0016650~45~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/"
str2 = str2 & ""
str2 = str2 & "11P~400~0016651~0017050~~~~~0~0~0~~0~~0~~0016651~0017050~400~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA"
str2 = str2 & "/11P~50~0017051~0017100~~~0017051~0017051~1~1~0~~0~~0~~0017052~0017100~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA"
str2 = str2 & "/11P~50~0017101~0017150~~~0017101~0017107~7~7~0~~0~~0~~0017108~0017150~43~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA"
str2 = str2 & "/11P~50~0017151~0017200~~~~~0~0~0~~0~~0~~0017151~0017200~50~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~850~00172"
str2 = str2 & "01~0018050~~~~~0~0~0~~0~~0~~0017201~0018050~850~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~500~0018051~0018550~~"
str2 = str2 & "~~~0~0~0~~0~~0~~0018051~0018550~500~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0018551~0018600~~~0018551~0018"
str2 = str2 & "600~50~48~2~18560;18578~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0018601~0018650~~~0018601~001862"
str2 = str2 & "4~24~24~0~~0~~0~~0018625~0018650~26~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0018651~0018700~~~0018651~0018"
str2 = str2 & "675~25~25~0~~0~~0~~0018676~0018700~25~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~50~0018701~0018750~~~0018701~00"
str2 = str2 & "18701~1~1~0~~0~~0~~0018702~0018750~49~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~300~0018751~0019050~~~~~0~0~0~~"
str2 = str2 & "0~~0~~0018751~0019050~300~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~950~0019051~0020000~~~~~0~0~0~~0~~0~~001905"
str2 = str2 & "1~0020000~950~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~500~~~0020001~0020500~~~0~0~0~~0~~0~~0020001~0020500~50"
str2 = str2 & "0~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~200~~~0020501~0020700~0020501~0020555~55~53~2~020528;020508~0~~0~~0"
str2 = str2 & "020556~0020700~145~0~H„a Æ¨n gi∏ trﬁ gia t®ng-BD012L~01GTKT2/001BD~AA/11P~4300~~~0020701~0025000~~~0~0~0~~0~~0~~0020701~0025"
str2 = str2 & "000~4300~0</S><S>~nguy‘n khæc th?nh~20/07/2011</S></S01>"


Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)

'TAX_Utilities_Srv_New.Convert(strTemp, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000002026¥a Æ¨n s´¥ D/N IAS 1007002 nga`y 18/01/10~81559176~~90370278~50~10~4518514~85851764~5~0~4292588~8811102~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ D/N IAS 1008001 nga`y 10/02/10 ~111354036~~123383973~50~10~6169199~117214774~5~0~5860739~12029938~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ D/N IAS 1010001 nga`y 20/04/10~74365638~~82399599~50~10~4119980~78279619~5~0~3913981~8033961~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ D/N IAS 1010003 nga`y 20/04/10~8038747~~8907199~50~10~445360~8461839~5~0~423092~8684"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   03201200000000102601/0101/01/1900<S01><S></S><S>PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30400608 Ngµy 11/06/10~10814137~~11982423~50~10~599121~11383302~5~0~569165~1168286~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~H„a Æ¨n sË BKK29401946 Ngµy 28/12/09~31191363~~34561067~50~10~1728053~32833014~5~0~1641651~3369704~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~H„a Æ¨n sË BKK29401811 Ngµy 25/11/09~61492151~~68135347~50~10~3406767~64728580~5~0~3236429~6643196~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~Ho"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000004026y 25/03/09~237804000~~263494737~50~10~13174737~250320000~5~0~12516000~25690737~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29300868 Ngµy 25/03/09~19817000~~21957895~50~10~1097895~20860000~5~0~1043000~2140895~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29301299 Ngµy 24/04/09~99580425~~110338421~50~10~5516921~104821500~5~0~5241075~10757996~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29301300 Ngµy 24/04/09~69359500~~76852632~50~10~3842632~73010000~5~0~3650500~7493132~PricewaterhouseCoopers Legal &amp; Tax Consultants"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   03201200000000302652~PwC International Assignment Services (Thailand) Ltd. Professional Services - Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ D/N IAS 1011001 nga`y 07/05/10~37182819~~41199799~50~10~2059990~39139809~5~0~1956990~4016980~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29300208 Ngµy 22/01/09~257621000~~285452632~50~10~14272632~271180000~5~0~13559000~27831632~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29300188 Ngµy 24/02/09~45579100~~50503158~50~10~2525158~47978000~5~0~2398900~4924058~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29300867 Ngµ"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000006026821~4562632~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29302881 ngµy 20/08/2009~20524863~~22742231~50~10~1137112~21605119~5~0~1080256~2217368~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29303360 ngµy 18/09/2009~15989545~~17716947~50~10~885847~16831100~5~0~841555~1727402~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29304027 ngµy 10/11/2009~16162349~~17908420~50~10~895421~17012999~5~0~850650~1746071~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29304162 ngµy 24/1"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000005026Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29301302 Ngµy 24/04/09~59451000~~65873684~50~10~3293684~62580000~5~0~3129000~6422684~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29301313 Ngµy 24/04/09~59451000~~65873684~50~10~3293684~62580000~5~0~3129000~6422684~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29301986 ngµy 22/06/2009~198170000~~219578947~50~10~10978947~208600000~5~0~10430000~21408947~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË BKK29302959 ngµy 26/08/2009~42233594~~46796226~50~10~2339811~44456415~5~0~2222"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000008026Professional  fee~~H„a Æ¨n sË BKK29100327 Ngµy 21/01/09~34648439~~38391622~50~10~1919581~36472041~5~0~1823602~3743183~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29100326 Ngµy 24/02/09~117226473~~129890829~50~10~6494541~123396287~5~0~6169814~12664355~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29100819 Ngµy 24/02/09~158724856~~175872417~50~10~8793621~167078796~5~0~8353940~17147561~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29101435 Ngµy 25/03/09~59476168~~65901571~50~10~3295079~62606493~5~0~3130325~6425404~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0320120000000070261/2009~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~H„a Æ¨n sË D/N 0908007 Ngµy 26/02/09~367037756~~406690034~50~10~20334502~386355533~5~0~19317777~39652279~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~Ho¥a Æ¨n s´¥ BKK30300817 nga`y 31/03/10~199604500~~221168421~50~10~11058421~210110000~5~0~10505500~21563921~PricewaterhouseCoopers Legal &amp; Tax Consultants Ltd - Ph› dﬁch v- Professional fee~~Ho¥a Æ¨n s´¥ BKK30301130 nga`y 30/04/10~31969854~~35423661~50~10~1771183~33652478~5~0~1682624~3453807~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ -"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0320120000000100266594578~~62708674~50~10~3135434~59573240~5~0~2978662~6114096~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29102904 Ngµy 26/05/09~2330677~~2582468~50~10~129123~2453344~5~0~122667~251790~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29102905 Ngµy 26/05/09~26102754~~28922719~50~10~1446136~27476583~5~0~1373829~2819965~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29102906 Ngµy 26/05/09~25058597~~27765758~50~10~1388288~26377471~5~0~1318874~2707162~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29102186 Ngµy 28/05/09~197201345~~218505645~50~10~10925282~207580"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000009026sË BKK29101437 Ngµy 25/03/09~46856306~~51918345~50~10~2595917~49322427~5~0~2466121~5062038~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29101136 Ngµy 12/03/09~166352816~~184324449~50~10~9216222~175108227~5~0~8755411~17971633~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29102182 Ngµy 27/04/09~95365944~~105668636~50~10~5283432~100385204~5~0~5019260~10302692~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29102184 Ngµy 27/04/09~16774694~~18586919~50~10~929346~17657573~5~0~882879~1812225~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29101784 Ngµy 02/04/09~5"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   03201200000001202686~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29103481 ngµy 29/06/2009~28061070~~31092598~50~10~1554630~29537968~5~0~1476898~3031528~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104164 ngµy 28/07/2009~41904434~~46431506~50~10~2321575~44109931~5~0~2205497~4527072~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104169 ngµy 28/07/2009~83656871~~92694594~50~10~4634730~88059864~5~0~4402993~9037723~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104171 ngµy 28/07/2009~8798748~~9749305~50~10~487465~9261840~5~0~463092~950557~PricewaterhouseCoopers ABAS Lt"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000011026363~5~0~10379018~21304300~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29103120 Ngµy 03/06/09~303191579~~335946348~50~10~16797317~319149031~5~0~15957452~32754769~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29103389 Ngµy 22/06/09~34667860~~38413141~50~10~1920657~36492484~5~0~1824624~3745281~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29103390 Ngµy 22/06/09~55447966~~61438189~50~10~3071909~58366280~5~0~2918314~5990223~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29103391 Ngµy 22/06/09~212209552~~235135238~50~10~11756762~223378476~5~0~11168924~229256"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000014026ssional  fee~~H„a Æ¨n sË BKK29105138 ngµy 22/09/2009~22313744~~24724370~50~10~1236219~23488152~5~0~1174408~2410627~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104934 ngµy 07/09/2009~49042319~~54340520~50~10~2717026~51623494~5~0~2581175~5298201~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29105178 ngµy 24/09/2009~152618247~~169106091~50~10~8455305~160650786~5~0~8032539~16487844~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29105244 ngµy 30/09/2009~264158628~~292696541~50~10~14634827~278061714~5~0~13903086~28537913~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000013026d Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104836 ngµy 25/08/2009~17577877~~19476872~50~10~973844~18503028~5~0~925151~1898995~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104835 ngµy 25/08/2009~54930940~~60865307~50~10~3043265~57822042~5~0~2891102~5934367~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104831 ngµy 25/08/2009~134965066~~149545780~50~10~7477289~142068491~5~0~7103425~14580714~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29104834 ngµy 25/08/2009~112058992~~124165088~50~10~6208254~117956834~5~0~5897842~12106096~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Profe"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0320120000000160266066 ngµy 18/11/2009~8523886~~9444749~50~10~472237~8972512~5~0~448626~920863~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29106248 ngµy 26/11/2009~111926416~~124018189~50~10~6200909~117817280~5~0~5890864~12091773~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29106801 ngµy 18/12/2009~55896623~~61935316~50~10~3096766~58838551~5~0~2941928~6038694~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29106803 ngµy 18/12/2009~84587081~~93725298~50~10~4686265~89039033~5~0~4451952~9138217~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË D/N 0911013 Ngµy 27/05/09~58635179~~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000015026 Æ¨n sË BKK29105430 ngµy 20/10/2009~328137218~~363586945~50~10~18179347~345407598~5~0~17270380~35449727~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29105520 ngµy 26/10/2009~34996624~~38777423~50~10~1938871~36838552~5~0~1841928~3780799~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29105324 ngµy 12/10/2009~19786680~~21924299~50~10~1096215~20828084~5~0~1041404~2137619~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29106072 ngµy 19/11/2009~342532882~~379537819~50~10~18976891~360560928~5~0~18028046~37004937~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK2910"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0320120000000180267445619~5~0~1372281~2816787~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1007011 nga`y 15/01/10 ~34718027~~38468728~50~10~1923436~36545292~5~0~1827265~3750701~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1008018 nga`y 25/02/10 ~5625866~~6233647~50~10~311682~5921965~5~0~296098~607780~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1008012 nga`y 25/02/10 ~17084368~~18930048~50~10~946502~17983546~5~0~899177~1845679~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1008017 nga`y 25/02/10 ~4546835~~5038044~50~10~251902~4786142~5~0"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   03201200000001702664969727~50~10~3248486~61721241~5~0~3086062~6334548~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË D/N 0911014 Ngµy 27/05/09~21321935~~23625413~50~10~1181271~22444142~5~0~1122207~2303478~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK30100269 nga`y 29/01/10~136045237~~150742645~50~10~7537132~143205513~5~0~7160276~14697408~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK30100324 nga`y 28/02/10~16407690~~18180266~50~10~909013~17271253~5~0~863563~1772576~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK30101265 nga`y 30/04/10~26073338~~28890125~50~10~1444506~2"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000020026cewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1011010 nga`y 26/05/10~601678~~666679~50~10~33334~633345~5~0~31667~65001~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1012026 nga`y 30/06/11~2761103~~3059394~50~10~152970~2906424~5~0~145321~298291~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1012019 nga`y 30/06/11~7252427~~8035930~50~10~401797~7634133~5~0~381707~783504~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK30106457 nga`y 30/06/11~14289444~~15833179~50~10~791659~15041520~5~0~752076~1543735~PricewaterhouseCoopers FAS Ltd. Ph› d"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000019026~239307~491209~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1009011 nga`y 08/03/10 ~3161619~~3503179~50~10~175159~3328020~5~0~166401~341560~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1009006 nga`y 08/03/10~10712540~~11869850~50~10~593493~11276357~5~0~563818~1157311~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1010013 nga`y 23/04/10~15540636~~17219541~50~10~860977~16358564~5~0~817928~1678905~PricewaterhouseCoopers ABAS Ltd Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N ABAS 1011009 nga`y 24/05/10~8532444~~9454232~50~10~472712~8981520~5~0~449076~921788~Pri"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000022026- Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600051 nga`y 28/02/11~563057647~~623886589~50~10~31194329~592692260~5~0~29634613~60828942~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600134 nga`y 31/03/11~612820908~~679025937~50~10~33951297~645074640~5~0~32253732~66205029~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600144 nga`y 31/05/11~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600141 nga`y 31/05/11~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professio"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000021026ﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK29600241 Ngµy 02/06/09~584377568~~647509771~50~10~32375489~615134282~5~0~30756714~63132203~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N FAS 1012025 nga`y 30/06/10~270144730~~299329341~50~10~14966467~284362874~5~0~14218144~29184611~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ D/N FAS 1004009 nga`y 21/10/09 ~1537778~~1703909~50~10~85195~1618714~5~0~80936~166131~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600049 nga`y 28/02/11~821326861~~910057464~50~10~45502873~864554591~5~0~43227730~88730603~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000024026Æ¨n s´¥ BKK31600554 nga`y 30/06/11~495425000~~548947368~50~10~27447368~521500000~5~0~26075000~53522368~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600556 nga`y 30/06/11~98094150~~108691579~50~10~5434579~103257000~5~0~5162850~10597429~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31690016 nga`y 30/09/11~34362083~~38074330~50~10~1903717~36170613~5~0~1808531~3712248~PricewaterhouseCoopers WMS Bangkok Limited Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK30900199 Ngµy 06/10/2010~71621016~~79358466~50~10~3967923~75390543~5~0~3769527~7737450~PricewaterhouseCoopers WMS Bangkok Limited Ph› dﬁch vÙ - Professional  fee"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000023026nal  fee~~Ho¥a Æ¨n s´¥ BKK31600142 nga`y 31/05/11~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600143 nga`y 31/05/11~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600111 nga`y 31/05/11~121127449~~134213240~50~10~6710662~127502578~5~0~6375129~13085791~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a Æ¨n s´¥ BKK31600360 nga`y 30/06/11~158536000~~175663158~50~10~8783158~166880000~5~0~8344000~17127158~PricewaterhouseCoopers FAS Ltd. Ph› dﬁch vÙ - Professional  fee~~Ho¥a"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000026026Floor, 1780 Bnagna-Trad Road, Bangna, Bangkok 10260~~Ho¥a Æ¨n s´¥ TJ321109013 nga`y 29/09/2011~138343893~24/02/2012~0~0~0~0~139741306~1~0~1397413~1397413~Chubb (Thailand) Limited Teo Hong Bangna Building, 7th Floor, 1780 Bnagna-Trad Road, Bangna, Bangkok 10260~~Ho¥a Æ¨n s´¥ TJ321109012 nga`y 29/09/2011~83654461~24/02/2012~92691924~50~10~4634596~88057328~5~0~4402866~9037462~Bloomberg Finance L.P. 731 Lexington Avenue NewYork, NY 10022, USA (Ph› dﬁch vÙ)~~H„a Æ¨n sË 5601452275 - 15/02/12~119405595~28/02/2012~132305368~50~10~6615268~125690100~5~0~6284505~12899773</S><S>11704810447~585240517~11259311239~0~557375913~1142616430</S><S>X~</S><S>~~~26/02/2013~1~~~</S></S01>"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   032012000000025026~~H„a Æ¨n sË BKK30900198 Ngµy 06/10/2010~105546729~~116949284~50~10~5847464~111101820~5~0~5555091~11402555~PricewaterhouseCoopers WMS Bangkok Limited Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK30900234 Ngµy 10/11/2010~12911965~~14306886~50~10~715344~13591542~5~0~679577~1394921~PricewaterhouseCoopers WMS Bangkok Limited Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK30900233 Ngµy 09/11/2010~54750606~~60665491~50~10~3033275~57632217~5~0~2881611~5914886~PricewaterhouseCoopers WMS Bangkok Limited Ph› dﬁch vÙ - Professional  fee~~H„a Æ¨n sË BKK31900036 Ngµy 14/03/2011~52352019~~58007777~50~10~2900389~55107388~5~0~2755369~5655758~Chubb (Thailand) Limited Teo Hong Bangna Building, 7th"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'' thang 5
'
'str2 = "aa316700100157406   0520120000000020349 ngµy 26/08/2011~341417175~~378301579~50~10~18915079~359386500~5~0~17969325~36884404~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31227708 ngµy 26/08/2011~202871075~~224787895~50~10~11239395~213548500~5~0~10677425~21916820~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 K"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   05201200000000103401/0101/01/1900<S01><S></S><S>PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31227709 ngµy 26/08/2011~202871075~~224787895~50~10~11239395~213548500~5~0~10677425~21916820~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL3122604"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000004034avers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31226068 ngµy 29/06/2011~232559525~~257683684~50~10~12884184~244799500~5~0~12239975~25124159~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31227705 ngµy 26/08/2011~202871075~~224787895~50~10~11239395~213548500~5~0~10677425~21916820~PricewaterhouseCoopers Taxation S"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000003034uala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31225865 ngµy 28/06/2011~424544835~~470409789~50~10~23520489~446889300~5~0~22344465~45864954~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31227703 ngµy 26/08/2011~94013425~~104170000~50~10~5208500~98961500~5~0~4948075~10156575~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Tr"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000006034~21273638~43666941~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ KUL31225518 nga`y 22/06/2011~164673713~~182463948~50~10~9123197~173340751~5~0~8667038~17790235~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31225807 ngµy 28/0"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000005034ervices Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31380342 ngµy 25/03/2011~64819783~~71822474~50~10~3591124~68231351~5~0~3411568~7002692~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ KUL31226065 nga`y 30/06/2011~404199113~~447866053~50~10~22393303~425472751~5~0"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000008034alaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31229063 ngµy 30/09/2011~195645318~~216781516~50~10~10839076~205942440~5~0~10297122~21136198~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31227505 ngµy 26/08/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kual"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0520120000000070346/2011~124901250~~138394737~50~10~6919737~131475000~5~0~6573750~13493487~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31226651 ngµy 27/07/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, M"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000010034Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31231190 ngµy 20/12/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31231189 ngµy 20/11/2011~124901250~~138394737~50~10~6919737~131475000~5~0~6573750"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000009034a Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31230617 ngµy 30/11/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31229042 ngµy 30/09/2011~124901250~~138394737~50~10~6919737~131475000~5~0~6573750~13493487~PricewaterhouseCoopers Taxation Services Sdn"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000012034184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31229765 ngµy 31/10/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysi"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000011034~13493487~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31229766 ngµy 31/10/2011~254798550~~282325263~50~10~14116263~268209000~5~0~13410450~27526713~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31229764 ngµy 31/10/2011~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000014034pur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË   KUL31229760 ngµy 31/10/2011~124901250~~138394737~50~10~6919737~131475000~5~0~6573750~13493487~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË   KUL31229759 ngµy 31/10/2011~144885450~~160537895~50~10~8026895~152511000~5~0~7625550~15652445~PricewaterhouseCoopers Taxation Services Sdn Bhd"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000013034a Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31230616 ngµy 30/11/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË  KUL31229762 ngµy 31/10/2011~184853850~~204824211~50~10~10241211~194583000~5~0~9729150~19970361~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lum"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000016034vate Limited The Millenia, 4th &amp; 7th Floor, Tower &apos;D&apos;, 1&amp;2, Murphy Road, Ulsoor, Bangalore-560008~~H„a Æ¨n sË NAT31100342 ngµy 29/06/2011~35971560~~39857684~50~10~1992884~37864800~5~0~1893240~3886124~PricewaterhouseCoopers Private Limited The Millenia, 4th &amp; 7th Floor, Tower &apos;D&apos;, 1&amp;2, Murphy Road, Ulsoor, Bangalore-560008~~H„a Æ¨n sË NAT31172222 ngµy 21/11/2011~381718204~~422956459~50~10~21147823~401808636~5~0~20090432~41238255~PricewaterhouseCoopers Private Limited The"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000015034(464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË   KUL29228796 ngµy 30/09/2009~116920300~~129551579~50~10~6477579~123074000~5~0~6153700~12631279~PricewaterhouseCoopers Private Limited The Millenia, 4th &amp; 7th Floor, Tower &apos;D&apos;, 1&amp;2, Murphy Road, Ulsoor, Bangalore-560008~~H„a Æ¨n sË NAT31100304 ngµy 31/05/2011~303260235~~336022421~50~10~16801121~319221300~5~0~15961065~32762186~PricewaterhouseCoopers Pri"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0520120000000180344th &amp; 7th Floor, Tower &apos;D&apos;, 1&amp;2, Murphy Road, Ulsoor, Bangalore-560008~~H„a Æ¨n sË NAT31100191 ngµy 31/03/ 2011~26927538~~29836607~50~10~1491830~28344777~5~0~1417239~2909069~PricewaterhouseCoopers Global Licensing Services Corporation-PwC Tower, 18 York Street, Suite 2600, Toronto ON, M5J 0B2, Canada~~H„a Æ¨n sË 9053 - 11 Ngµy 25/10/2011~278022294~~0~0~0~0~308913660~10~0~30891366~30891366~PricewaterhouseCoopers Global Licensing Services Corporation-PwC Tower, 18 York Street, Suite 2600,"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000017034 Millenia, 4th &amp; 7th Floor, Tower &apos;D&apos;, 1&amp;2, Murphy Road, Ulsoor, Bangalore-560008~~H„a Æ¨n sË NAT31172296 ngµy 30/11/2011~312393014~~346141844~50~10~17307092~328834752~5~0~16441738~33748830~PricewaterhouseCoopers Private Limited The Millenia, 4th &amp; 7th Floor, Tower &apos;D&apos;, 1&amp;2, Murphy Road, Ulsoor, Bangalore-560008~~H„a Æ¨n sË NAT31172296 ngµy 30/11/ 2011~154572600~~171271579~50~10~8563579~162708000~5~0~8135400~16698979~PricewaterhouseCoopers Private Limited The Millenia,"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000020034~~0~0~0~0~72641213~10~0~7264121~7264121~PricewaterhouseCoopers Global Licensing Services Corporation-PwC Tower, 18 York Street, Suite 2600, Toronto ON, M5J 0B2, Canada~~H„a Æ¨n sË 8940 - 11 Ngµy 14/09/2011~22709060~~0~0~0~0~25232289~10~0~2523229~2523229~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  BV140/12 ngµy 14/09/2011~1059961968~~1174473095~50~10~58723655~1115749440~5~0~55787472~114511127~PricewaterhouseCoopers Services BV Fascinato Boule"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000019034Toronto ON, M5J 0B2, Canada~~H„a Æ¨n sË 9010 - 11 Ngµy 25/10/2011~3375731064~~0~0~0~0~3750812293~10~0~375081229~375081229~PricewaterhouseCoopers Global Licensing Services Corporation-PwC Tower, 18 York Street, Suite 2600, Toronto ON, M5J 0B2, Canada~~H„a Æ¨n sË 8863 - 11 Ngµy 18/05/2011~35417800~~0~0~0~0~39353111~10~0~3935311~3935311~PricewaterhouseCoopers Global Licensing Services Corporation-PwC Tower, 18 York Street, Suite 2600, Toronto ON, M5J 0B2, Canada~~H„a Æ¨n sË 8891 - 11 Ngµy 18/05/2011~65377092"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000022034522~230175912~5~0~11508796~23623318~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  BSS076/12/ngµy 14/09/2011~96441646~~0~0~0~0~107157384~10~0~10715738~10715738~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  BV101/12 ngµy 7/09/2011~135852592~~150529187~50~10~7526459~143002728~5~0~7150136~14676595~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Nethe"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000021034vard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  ATP019/11 ngµy 06/10/2011~287545291~~0~0~0~0~319494768~10~0~31949477~31949477~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  BV292/12 ngµy 09/11/2011~1059961968~~1174473095~50~10~58723655~1115749440~5~0~55787472~114511127~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  ECC013/12 ngµy 09/11/2011~218667116~~242290434~50~10~12114"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000024034waterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  BV379/12 ngµy 17/01/2012~130138239~~144197495~50~10~7209875~136987620~5~0~6849381~14059256~PricewaterhouseCoopers Advisory Services Sdn Bhd (573259-K) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31380388 ngµy 08/04/2011~325424997~~360581714~50~10~18029086~342552628~5~0~17127631~35156717~PricewaterhouseCoopers Advisory"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000023034rlands~~H„a Æ¨n sË  IFA130/11 ngµy 14/09/2011~85372502~~94595571~50~10~4729779~89865792~5~0~4493290~9223069~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  BV252/12 ngµy 09/11/2011~134715966~~149269768~50~10~7463488~141806280~5~0~7090314~14553802~PricewaterhouseCoopers Services BV Fascinato Boulevard 350, 3065 WWB Rotterdam, The Netherlands~~H„a Æ¨n sË  ECC011/12 ngµy 09/11/2011~42249844~~46814232~50~10~2340712~44473520~5~0~2223676~4564388~Price"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000026034040~769819~PricewaterhouseCoopers Advisory Services Sdn Bhd (573259-K) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31380982 ngµy 14/07/2011~73821635~~81796825~50~10~4089841~77706984~5~0~3885349~7975190~PricewaterhouseCoopers Advisory Services Sdn Bhd (573259-K) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31381379 ngµy 3/11/2011~17026538"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000025034Services Sdn Bhd (573259-K) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31380341 ngµy 25/03/2011~105476608~~116871588~50~10~5843579~111028009~5~0~5551400~11394979~PricewaterhouseCoopers Advisory Services Sdn Bhd (573259-K) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31380343 ngµy 25/03/ 2011~7125766~~7895586~50~10~394779~7500807~5~0~375"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000028034L31035676 nga`y 30/06/2011~3389284~~3755439~50~10~187772~3567667~5~0~178383~366155~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ KUL31042418 nga`y 29/11/2011~11810598~~13086535~50~10~654327~12432208~5~0~621610~1275937~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Mala"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000027034~~18865970~50~10~943299~17922671~5~0~896134~1839433~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ KUL3135670 nga`y 30/06/2011~10078031~~11166793~50~10~558340~10608454~5~0~530423~1088763~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ KU"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000030034O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31042479 ngµy 05/12/2011~370706910~~410755579~50~10~20537779~390217800~5~0~19510890~40048669~PwC International Assignment Services Sdn Bhd (777150-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL31730316 ngµy 16/05/2011~75707923~~83886895~50~10~4194345~79692551~5~0~3984628~8178973~PricewaterhouseCoopers PO Box 258, Strathvale House, Grand Cayman KY1-110"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000029034ysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ KUL31042420 nga`y 29/11/2011~12088049~~13393960~50~10~669698~12724262~5~0~636213~1305911~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n s´¥ PEN31038896 nga`y 11/08/2011~75051292~~83159326~50~10~4157966~79001360~5~0~3950068~8108034~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0520120000000320340~12361038~234859716~5~0~11742986~24104024~Advantech Peripherals Singapore Pte. Ltd. Blk 1026 Tai Seng Avenue11 07-3542/3546 Tai Seng Indutrial Estate Singapore 534413, Singapore~~H„a Æ¨n sË 040811PWVngµy 08/04/2011~23061767~~25553204~50~10~1277660~24275544~5~0~1213777~2491437~IRON MOUNTAIN DIGITAL P.O. Box 27128 New York, NY 10087-7128 (Ph› ph©`n m™`m)~~H„a Æ¨n sË 040811PWVngµy 08/04/2011~74764053~~0~0~0~0~83071170~10~0~8307117~8307117~PROXY NETWORKS 320 Congress Street, 3rd Floor, Boston MA02210, Un"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0520120000000310344, Cayman Islands~~H„a Æ¨n sË 22067 ngµy 09/05/2011~105505708~~116903832~50~10~5845192~111058640~5~0~5552932~11398124~PricewaterhouseCoopers PO Box 258, Strathvale House, Grand Cayman KY1-1104, Cayman Islands~~H„a Æ¨n sË 22374 ngµy 09/06/2011~79268000~~87831579~50~10~4391579~83440000~5~0~4172000~8563579~MFEC Public Company Limited 333 LaoPeng Nguan Towers, Soi Chosipuang, Vibhavadi-Rangsit Rd, Chompol, Chatuchak, Bangkok 10900 (Ph› di?ch vu?)~~H„a Æ¨n sË 221110098ngµy 06/06/2011~223116730~~247220753~50~1"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000034034ed. 30 Lavant Rd, Stone Cross, Pevensey, East Sussex, BN24 5EZ, England~~H„a Æ¨n sË 4780 - 24/11/11~20299319~~22493011~50~10~1124651~21368360~5~0~1068418~2193069~Applied Intergrators Sdn. Ghd. Unit 2-1, Level 2, Tower 3, Avenue 3, Bangsar South, No.8, Japan Kerinchi 59200 Kuaka Lumpur, Malaysia~~H„a Æ¨n sË APP2012-15 - 16/01/12~93938421~~104084821~50~10~5204241~98880580~5~0~4944029~10148270</S><S>11981872961~599093655~16145731995~0~1045434235~1644527890</S><S>X~</S><S>~~~27/02/2013~1~~~</S></S01>"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   052012000000033034ited States (617) 453-2700 (Ph› ph©`n m™`m)~~H„a Æ¨n sË 1107012ngµy 07/07/2011~39880350~~0~0~0~0~44311500~10~0~4431150~4431150~Dell Marketing L.P. PO BOX 676021 C/O Dell USA L.P. Dallats, TX 75267-6021~~H„a Æ¨n sË XFNMP2JP5 - 17/02/2012~10768349~~0~0~0~0~11965296~10~0~1196530~1196530~Centillion Group Limited 11C Eastern Commercial Centre, 83 Nam On Street, Shaukeiwan, Hong Kong~~H„a Æ¨n sË 201111004 - 11/11/11~73538500~~81389979~50~10~4069499~77320480~5~0~3866024~7935523~Bibby Factors International Limit"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'' thang 6
'str2 = "aa316700100157406   062012000000002086ravers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ KUL32032208 nga`y 09/03/2012~85014930~~94199368~50~10~4709968~89489400~5~0~4474470~9184438~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ KUL32032935 nga`y 2"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000000108601/0101/01/1900<S01><S></S><S>PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ KUL32032491 nga`y 12/03/2012~640913091~~710153009~50~10~35507650~674645359~5~0~33732268~69239918~PricewaterhouseCoopers, (AF 1146) Chartered Accountants Level 10, 1 Sentral, Jalan T"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000004086ation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL31222577 ngµy 31/03/11~119301000~~132189474~50~10~6609474~125580000~5~0~6279000~12888474~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000030863/03/2012~346797500~~384263158~50~10~19213158~365050000~5~0~18252500~37465658~PricewaterhouseCoopers Advisory Services Sdn Bhd (573259-K) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~Ho¥a Æ¨n s´¥ KUL32380025 nga`y 11/11/2012~2021334000~~2239705263~50~10~111985263~2127720000~5~0~106386000~218371263~PricewaterhouseCoopers Tax"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000006086~10203375~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL31230743 ngµy 30/11/11~164038875~~181760526~50~10~9088026~172672500~5~0~8633625~17721651~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur S"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000005086Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL31227628 ngµy 26/08/11~203805875~~225823684~50~10~11291184~214532500~5~0~10726625~22017809~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL31230742 ngµy 30/11/11~94446625~~104650000~50~10~5232500~99417500~5~0~4970875"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000000808616~50~10~20103816~381972500~5~0~19098625~39202441~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32220622 ngµy 31/01/12~223689375~~247855263~50~10~12392763~235462500~5~0~11773125~24165888~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000007086entral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32220334 ngµy 31/01/12~231841610~~256888211~50~10~12844411~244043800~5~0~12202190~25046601~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32220621 ngµy 31/01/12~362873875~~4020763"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000010086244 ngµy 29/02/12~99417500~~110157895~50~10~5507895~104650000~5~0~5232500~10740395~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32223010 ngµy 30/03/12~119301000~~132189474~50~10~6609474~125580000~5~0~6279000~12888474~PricewaterhouseCoopers Taxation Servi"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000009086, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32220900 ngµy 23/02/12~43743700~~48469474~50~10~2423474~46046000~5~0~2302300~4725774~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32222"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000012086aysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32223012 ngµy 30/03/12~278369000~~308442105~50~10~15422105~293020000~5~0~14651000~30073105~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32223013 ngµy 30/03/12~99417500~~110157895~50~10~5507895~104650000~5~0~5232500~10740395~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000011086ces Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32223011 ngµy 30/03/12~119301000~~132189474~50~10~6609474~125580000~5~0~6279000~12888474~PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Mal"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000014086P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n sË KUL30731131 ngµy 16/08/2010~43283036~~47959042~50~10~2397952~45561090~5~0~2278055~4676007~PricewaterhouseCoopers ( Macau) Ltd Ph› dﬁch vÙ~~H„a Æ¨n sË HKG29015407 ngµy 24/08/2009~82566409~~91486326~50~10~4574316~86912010~5~0~4345601~8919917~Advokatfirmaet PricewaterhouseCoopers AS Postboks 748 Sentrum, NO-0106 Oslo Ph› dﬁch vÙ"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000013086PricewaterhouseCoopers Taxation Services Sdn Bhd (464731-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral, P O Box 101192, 50706 Kuala Lumpur, Malaysia Ph› dﬁch vÙ~~H„a Æ¨n KUL32223898 ngµy 27/04/12~400453690~~443716000~50~10~22185800~421530200~5~0~21076510~43262310~PricewaterhouseCoopers Taxation Services Sdn Bhd (777150-M) Level 10, 1 Sentral, Jalan Travers, Kuala Lumpur Sentral,"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000016086e Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG31017724 ngµy 24/08/2011~167735250~~185856232~50~10~9292812~176563421~5~0~8828171~18120983~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG31017821 ngµy 24/08/2011~189241847~~209686257~50~10~10484313~199201944~5~0"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000015086~~H„a Æ¨n sË OSL31710318 ngµy 09/12/2011~187739917~~208022068~50~10~10401103~197621300~5~0~9881065~20282168~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG31007894 ngµy 06/05/2011~300196437~~332627631~50~10~16631382~315996250~5~0~15799813~32431195~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulif"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000180861023667 ngµy 10/11/2011~111517591~~123565198~50~10~6178260~117386938~5~0~5869347~12047607~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG31023659 ngµy 10/11/2011~750960063~~832088712~50~10~41604436~790484276~5~0~39524214~81128650~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000017086~9960097~20444410~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG31017793 ngµy 24/08/2011~750960063~~832088712~50~10~41604436~790484276~5~0~39524214~81128650~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG3"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000020086 Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË FNB - 110340 ngµy 25/10/2011~5801796~~0~0~0~0~5860400~1~0~58604~58604~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË FNB - 110283 ngµy 11/10/2011~1243242~~0~0~0~0~1255800~1~0~12558~12558~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000019086, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË HKG31023637 ngµy 10/11/2011~167735250~~185856232~50~10~9292812~176563421~5~0~8828171~18120983~PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË FNB - 110301 ngµy 25/10/2011~9945936~~0~0~0~0~10046400~1~0~100464~100464~PricewaterhouseCoopers"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000220860~4395300~9021932~PricewaterhouseCoopers 33rd Floor Cheung Kong Center 2 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG30018846 ngµy 31/08/2010~515837641~~571565253~50~10~28578263~542986990~5~0~27149350~55727613~PricewaterhouseCoopers 21st Floor Edinburgh Tower, 15 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG31014277 ngµy 27/06/2011~134720446~~149274733~50~1"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000021086PricewaterhouseCoopers Ltd 20th Floor Tower A, Maulife Financial Centre, 223-231 Wai Yip Street, Kwun Tong, Kowloon, Hong Kong~~H„a Æ¨n sË FNB - 120416 ngµy 02/05/2012~4144140~~0~0~0~0~4186000~1~0~41860~41860~PricewaterhouseCoopers 33rd Floor Cheung Kong Center 2 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG30025965 ngµy 22/12/2010~83510700~~92532632~50~10~4626632~87906000~5~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000024086648592~~112630018~50~10~5631501~106998518~5~0~5349926~10981427~PricewaterhouseCoopers 21st Floor Edinburgh Tower, 15 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG32002981 ngµy 14/02/2012~123558995~~136907474~50~10~6845374~130062100~5~0~6503105~13348479~PricewaterhouseCoopers 21st Floor Edinburgh Tower, 15 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG30011756"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000230860~7463737~141810996~5~0~7090550~14554287~PricewaterhouseCoopers 21st Floor Edinburgh Tower, 15 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG30026406 ngµy 15/12/2010~677145647~~750299886~50~10~37514994~712784892~5~0~35639245~73154239~PricewaterhouseCoopers 21st Floor Edinburgh Tower, 15 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG31016391 ngµy 19/08/2011~101"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000026086/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK31401000 ngµy 24/06/2011~42408380~~46989895~50~10~2349495~44640400~5~0~2232020~4581515~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK32900002 ngµy"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'
'str2 = "aa316700100157406   062012000000025086 ngµy 14/04/2010~15720889~~17419268~50~10~870963~16548305~5~0~827415~1698378~PricewaterhouseCoopers WMS Asia Pacific Ltd 21st Floor Edinburgh Tower, 15 Queen&apos;s Road Central Hong Kong Ph› dﬁch vÙ~~H„a Æ¨n sË HKG30680042 ngµy 12/04/2010~248076688~~274877217~50~10~13743861~261133356~5~0~13056668~26800529~PwC International Assignment Services (Thailand) Ltd. 15th Floor Bangkok City Tower, 179"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000028086ok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK32900003 ngµy 18/01/2012~15705171~~17401851~50~10~870093~16531759~5~0~826588~1696681~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'str2 = "aa316700100157406   062012000000027086 18/01/2012~52851939~~58561705~50~10~2928085~55633620~5~0~2781681~5709766~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK32900004 ngµy 18/01/2012~5567388~~6168851~50~10~308443~5860408~5~0~293020~601463~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangk"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'str2 = "aa316700100157406   062012000000030086ed. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK31900283 ngµy 14/11/2011~28932225~~32057868~50~10~1602893~30454974~5~0~1522749~3125642~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120,"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000029086BKK32900005 ngµy 18/01/2012~12131571~~13442184~50~10~672109~12770075~5~0~638504~1310613~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK31900282 ngµy 14/11/2011~43789427~~48520140~50~10~2426007~46094133~5~0~2304707~4730714~PricewaterhouseCoopers WMS LtdLimit"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'str2 = "aa316700100157406   062012000000032086useCoopers Legal &amp; Tax Consultants Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31300110 ngµy 17/01/2011~69948363~~77505112~50~10~3875256~73629856~5~0~3681493~7556749~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathor"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000031086 Thailand~~H„a Æ¨n sË BKK31900284 ngµy 14/11/2011~25217925~~27942299~50~10~1397115~26545184~5~0~1327259~2724374~PricewaterhouseCoopers WMS LtdLimited. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand~~H„a Æ¨n sË BKK31600722 ngµy 02/09/2011~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~Pricewaterho"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000034086ces- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31101803 ngµy 25/04/2011~14417066~~15974588~50~10~798729~15175859~5~0~758793~1557522~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31100235 ngµy 24/01/2011~19279761~~21362616~50~10~1068131~20294485"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000033086n Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30101808 ngµy 21/04/2010~52208679~~57848952~50~10~2892448~54956504~5~0~2747825~5640273~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional servi"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000036086uth Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31100637 ngµy 23/02/2011~84509200~~93639003~50~10~4681950~88957053~5~0~4447853~9129803~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professi"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000035086~5~0~1014724~2082855~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31101802 ngµy 25/04/2011~22017876~~24396539~50~10~1219827~23176712~5~0~1158836~2378663~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 So"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000380860~12660790~240555017~5~0~12027751~24688541~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31100635 ngµy 23/02/2011~88913132~~98518706~50~10~4925935~93592771~5~0~4679639~9605574~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok Ci"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000037086onal services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30106449 ngµy 20/12/2010~117998741~~130746528~50~10~6537326~124209201~5~0~6210460~12747786~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31100216 ngµy 24/01/2011~228527266~~253215807~50~1"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000040086 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30105177 ngµy 19/10/2010~447951593~~496345255~50~10~24817263~471527993~5~0~23576400~48393663~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31101144 ngµy 25/03/20"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000039086ty Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30105904 ngµy 18/11/2010~130860965~~144998299~50~10~7249915~137748384~5~0~6887419~14137334~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok,"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000042086.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32100317 ngµy 27/01/2012~37940303~~42039117~50~10~2101956~39937161~5~0~1996858~4098814~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000004108611~56952473~~63105233~50~10~3155262~59949971~5~0~2997499~6152761~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31101814 ngµy 20/06/2011~67239279~~74503356~50~10~3725168~70778189~5~0~3538909~7264077~PricewaterhouseCoopers ABAS Ltd"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000004408668 ngµy 18/11/2011~378635525~~419540748~50~10~20977037~398563711~5~0~19928186~40905223~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600555 ngµy 27/06/11~356706000~~395242105~50~10~19762105~375480000~5~0~18774000~38536105~Price"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000043086District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32100319 ngµy 27/01/2012~149768089~~165948021~50~10~8297401~157650620~5~0~7882531~16179932~PricewaterhouseCoopers ABAS Ltd.15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK311063"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000046086mahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600372 ngµy 07/06/11~253409888~~280786579~50~10~14039329~266747250~5~0~13337363~27376692~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph›"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000045086waterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600428 ngµy 10/06/11~852621669~~944733151~50~10~47236658~897496494~5~0~44874825~92111483~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thung"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000480860160906~20856597~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600604 ngµy 12/07/11~224574567~~248836085~50~10~12441804~236394281~5~0~11819714~24261518~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 So"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000047086 dﬁch vÙ~~H„a Æ¨n sË BKK31600652 ngµy 05/08/11~26619582~~29495381~50~10~1474769~28020612~5~0~1401031~2875800~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600830 ngµy 14/10/11~193057214~~213913811~50~10~10695691~203218120~5~0~1"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000050086essional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600924 ngµy 02/12/2011~269511200~~298627368~50~10~14931368~283696000~5~0~14184800~29116168~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600553 ngµy11/07/2011 ~55586685~~61591895~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000049086uth Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600669 ngµy 15/08/11~353456012~~391641011~50~10~19582051~372058960~5~0~18602948~38184999~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Prof"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000052086ty Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32600027 ngµy 18/01/2012~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 101"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000005108650~10~3079595~58512300~5~0~2925615~6005210~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32600026 ngµy 19/01/2012~39634000~~43915789~50~10~2195789~41720000~5~0~2086000~4281789~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok Ci"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000540860~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600928 ngµy 30/11/2011~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floo"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000005308620, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32600029 ngµy 18/01/2012~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600929 ngµy 30/11/2011~3976700"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000056086Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600583 ngµy 07/07/2011~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600584 ngµy 07/07/"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000055086r Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600927 ngµy 30/11/2011~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District,"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000058086d. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600908 ngµy 29/11/2011~141726208~~157037349~50~10~7851867~149185482~5~0~7459274~15311141~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sa"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000570862011~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600593 ngµy 08/07/2011~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Lt"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   06201200000006008630600324 ngµy 10/06/2010~89172924~~98806564~50~10~4940328~93866236~5~0~4693312~9633640~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30600724 ngµy 26/10/2010~185882292~~205963758~50~10~10298188~195665570~5~0~9783279~20081467~Pric"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000059086thorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK30600325 ngµy 10/06/2010~102731284~~113829678~50~10~5691484~108138194~5~0~5406910~11098394~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000062086amek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31690057 ngµy 26/09/2011~13094279~~14508896~50~10~725445~13783452~5~0~689173~1414618~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000061086ewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32600070 ngµy 07/02/2012~39767000~~44063158~50~10~2203158~41860000~5~0~2093000~4296158~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmah"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000064086128895026~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600620 ngµy 20/07/2011~302592472~~335282517~50~10~16764126~318518391~5~0~15925920~32690046~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South S"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000063086~H„a Æ¨n sË BKK31690056 ngµy 29/09/2012~8303151~~9200167~50~10~460008~8740159~5~0~437008~897016~PricewaterhouseCoopers FAS Ltd. 15th Floor Bangkok City Tower, 179/74-80 South Sathorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK32600077 ngµy 14/02/2011~1193105242~~1322000268~50~10~66100013~1255900255~5~0~62795013~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000660860477437~~44850346~50~10~2242517~42606380~5~0~2130319~4372836~PricewaterhouseCoopers LLP 8 Cross Street 1117-00, PWC Building, Singapore 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP32000715 ngµy 31/01/2012~223755786~~247928849~50~10~12396442~235532406~5~0~11776620~24173062~PricewaterhouseCoopers Services LLP 8 Cross Street 1117-00, PWC Building, Singapore 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP31706173"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000065086athorn Road Thungmahamek Sub district, Sathorn District, Bangkok, 10120, Thailand. Professional services- Ph› dﬁch vÙ~~H„a Æ¨n sË BKK31600695 ngµy 15/09/2011~10071390~~11159435~50~10~557972~10601464~5~0~530073~1088045~PricewaterhouseCoopers Asia Acturial Services (Singapore) Pte Ltd  8 Cross Street 1117-00, PWC Building, Singapore 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP32500032 ngµy 04/04/2012~4"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000068086re 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP31753182 ngµy 24/11/2011~13647446~~15121824~50~10~756091~14365733~5~0~718287~1474378~PwC International Assignment Services (Singapore) Pte Ltd 8 Cross Street 1117-00, PWC Building, Singapore 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP31752024 ngµy 26/08/2011~108993500~~120768421~50~10~6038421~114730000~5~0~5736500~11774921~PwC International Assignment Services ("
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000067086ngµy 12/10/2011~363167385~~402401535~50~10~20120077~382280360~5~0~19114018~39234095~PricewaterhouseCoopers WMS Pte Ltd 8 Cross Street 1117-00, PWC Building, Singapore 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP31450124 ngµy 31/08/2011~61348496~~67976173~50~10~3398809~64576320~5~0~3228816~6627625~PwC International Assignment Services (Singapore) Pte Ltd 8 Cross Street 1117-00, PWC Building, Singapo"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000070086041~PricewaterhouseCoopers Services BV. Fasinato Boulevard 350, 3065 WB Rotterdam, The Nethrlands-Ph› dﬁch vÙ~~H„a Æ¨n sË BV244/10 ngµy 21/10/2009, IFA125/09~106391725~~117885569~50~10~5894278~111991290~5~0~5599565~11493843~PricewaterhouseCoopers Services BV. Fasinato Boulevard 350, 3065 WB Rotterdam, The Nethrlands-Ph› dﬁch vÙ~~H„a Æ¨n sË ECC030/09 ngµy 21/10/2009 &amp; ECC029/10~34134260~~37"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000069086Singapore) Pte Ltd 8 Cross Street 1117-00, PWC Building, Singapore 0458424 Ph› dﬁch vÙ~~H„a Æ¨n sË SPP31752392 ngµy 26/09/2011~32074641~~35539768~50~10~1776988~33762780~5~0~1688139~3465127~PricewaterhouseCoopers Services BV. Fasinato Boulevard 350, 3065 WB Rotterdam, The Nethrlands-Ph› dﬁch vÙ~~H„a Æ¨n sË BV087/10 ngµy 01/09/2009~120065275~~133036316~50~10~6651816~126384500~5~0~6319225~12971"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000720867831579~50~10~4391579~83440000~5~0~4172000~8563579~PricewaterhouseCoopers ABN 52 780 433 757 Darling Park Tower 2, 201 Sussex Street, GPO BOX 2650, SYDNEY NSW 1171, DX 77 Sydney Australia Ph› dﬁch vÙ~~H„a Æ¨n sË 31057214 ngµy 15/06/2011~198863595~~220347474~50~10~11017374~209330100~5~0~10466505~21483879~PricewaterhouseCoopers ABN 52 780 433 757 Darling Park Tower 2, 201 Sussex Street, GPO BOX"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000071086821895~50~10~1891095~35930800~5~0~1796540~3687635~PricewaterhouseCoopers SA. Avenue Giuseppe-Motta 50, Case postale, CH-1211 Geneve 2 Ph› dﬁch vÙ~~H„a Æ¨n sË CHD32105586 ngµy 28/02/2012~228660250~~253363158~50~10~12668158~240695000~5~0~12034750~24702908~PricewaterhouseCoopers LLP P.O. Box 7247-8001, Philadelphia, PA 19170 8001, USA Ph› dﬁch vÙ~~H„a Æ¨n sË 1031828085-5 ngµy 27/06/11~79268000~~8"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000074086~0~11663348~23940556~PricewaterhouseCoopers ABN 52 780 433 757 Darling Park Tower 2, 201 Sussex Street, GPO BOX 2650, SYDNEY NSW 1171, DX 77 Sydney Australia Ph› dﬁch vÙ~~H„a Æ¨n sË 31116272 ngµy 16/11/2011~307906638~~341170789~50~10~17058539~324112250~5~0~16205613~33264152~PricewaterhouseCoopers ABN 52 780 433 757 Darling Park Tower 2, 201 Sussex Street, GPO BOX 2650, SYDNEY NSW 1171, DX 77 S"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000730862650, SYDNEY NSW 1171, DX 77 Sydney Australia Ph› dﬁch vÙ~~H„a Æ¨n sË 31085131 ngµy 19/08/2011~270848848~~300109526~50~10~15005476~285104050~5~0~14255203~29260679~PricewaterhouseCoopers ABN 52 780 433 757 Darling Park Tower 2, 201 Sussex Street, GPO BOX 2650, SYDNEY NSW 1171, DX 77 Sydney Australia Ph› dﬁch vÙ~~H„a Æ¨n sË 31093815 ngµy 14/09/2011~221603603~~245544158~50~10~12277208~233266950~5"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000076086rhouseCoopers 188 Quay Street, Private Bag 92162, Auckland 1142, New Zealand-Ph› dﬁch vÙ~~H„a Æ¨n sË NZD31033309 ngµy 30/09/2011~357461404~~396079118~50~10~19803956~376275162~5~0~18813758~38617714~Price Waterhouse &amp; Co. 32 Khadar Nawaz Khan Road, PWC Centre, Nungambakkam Chennai 600006- Ph› dﬁch vÙ~~H„a Æ¨n sË CHN31455025 ngµy 11/05/2011~12963957~~14364495~50~10~718225~13646270~5~0~682314~"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000075086ydney Australia Ph› dﬁch vÙ~~H„a Æ¨n sË 31126944 ngµy 13/12/2011~251923613~~279139737~50~10~13956987~265182750~5~0~13259138~27216125~PricewaterhouseCoopers ABN 52 780 433 757 Darling Park Tower 2, 201 Sussex Street, GPO BOX 2650, SYDNEY NSW 1171, DX 77 Sydney Australia Ph› dﬁch vÙ~~H„a Æ¨n sË 32014590 ngµy 21/02/2012~312266378~~346001526~50~10~17300076~328701450~5~0~16435073~33735149~Pricewate"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000078086011~24750317~~27424174~50~10~1371209~26053373~5~0~1302669~2673878~The Association of Chartered Certified Accountants 2 Central Quay, 89 Hydepark Street, Glasgow G38 BW UK Ph› Æµo tπo ACCA~~H„a Æ¨n sË VP 033784 Ngµy 12/12/2011~735983759~~815494470~50~10~40774724~774720045~5~0~38736002~79510726~ANT Office express Co., Ltd 69/15 Moo 2 Srinakarin Road, Nongbon, Pravet, Bangkok 10250 Ph› dﬁch vÙ in"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   0620120000000770861400539~The Association of Chartered Certified Accountants 2 Central Quay, 89 Hydepark Street, Glasgow G38 BW UK Ph› Æµo tπo ACCA~~H„a Æ¨n sË VP 033784 Ngµy 12/12/2011~1281177174~~1419586897~50~10~70979345~1348607552~5~0~67430378~138409723~The Association of Chartered Certified Accountants 2 Central Quay, 89 Hydepark Street, Glasgow G38 BW UK Ph› Æµo tπo ACCA~~H„a Æ¨n sË VP 033784 Ngµy 12/12/2"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000080086United Kingdom Ph› dﬁch vÙ~~H„a Æ¨n sË PwC/PVN/01 ngµy 02/02/2012~617299550~~683988421~50~10~34199421~649789000~5~0~32489450~66688871~Ewart Consultancy Associates 4, St. Catherine&apos;s Gardens, Edinburgh EH12 7AZ United Kingdom Ph› dﬁch vÙ~~H„a Æ¨n sË PwC/PVN/03 ngµy 06/04/2012~111601000~~123657618~50~10~6182881~117474737~5~0~5873737~12056618~Fujitsu System Business (Thailand) Ltd. Exchange"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000079086~~H„a Æ¨n sË IV0005147 ngµy 19/03/2012~18873617~~20912595~50~10~1045630~19866760~5~0~993338~2038968~Ewart Consultancy Associates 4, St. Catherine&apos;s Gardens, Edinburgh EH12 7AZ United Kingdom Ph› dﬁch vÙ~~H„a Æ¨n sË PwC/PVN/01 ngµy 02/02/2012~554876000~~614821053~50~10~30741053~584080000~5~0~29204000~59945053~Ewart Consultancy Associates 4, St. Catherine&apos;s Gardens, Edinburgh EH12 7AZ"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000082086tr◊ ph«n m“m~~H„a Æ¨n sË 311075708 ngµy 16/09/2011~195339586~~216442754~50~10~10822138~205620617~5~0~10281031~21103169~Networkers International (UK) PLC Hanover Place, 8 Ravensbourne Road, Bromley, BR1 1HP United Kingdom -Ph› dﬁch vÙ tuy”n dÙng~~H„a Æ¨n sË 400020784 ngµy 21/04/2012~803711948~~890539554~50~10~44526978~846011540~5~0~42300577~86827555~APPLIED INTEGRATORS SDN. BHD. Unit 2-1, Level"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000081086Tower, 22nd - 23nd Floor, No. 388 Sukhumvit Road, Kwaeng Klongtoey, Khet Klongtoey, Bangkok 10110, Thailand. Ph› b∂o tr◊ ph«n m“m~~H„a Æ¨n sË 311077343 ngµy 02/11/2011~46230369~~51224785~50~10~2561239~48663546~5~0~2433177~4994416~Fujitsu System Business (Thailand) Ltd. Exchange Tower, 22nd - 23nd Floor, No. 388 Sukhumvit Road, Kwaeng Klongtoey, Khet Klongtoey, Bangkok 10110, Thailand. Ph› b∂o"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000084086~94644247~10~0~9464425~9464425~APPLIED INTEGRATORS SDN. BHD. Unit 2-1, Level 2, Tower 3, Avenue 3, Bangsar South, No. 8, Jalan Kerinchi, 59200 Kuala Lumpur, Malaysia -Ph› b∂o tr◊~~H„a Æ¨n sË APP 2010-74 ngµy 06/10/2010~56419919~~0~0~0~0~62688799~10~0~6268880~6268880~QUALYS INC, 1600 Bridge Parkway, 2nd Floor, Redwood Shores CA 94065, USA. Ph› dﬁch vÙ- Subscription fee Ph› dﬁch vÙ - Subcriptio"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'
'str2 = "aa316700100157406   062012000000083086 2, Tower 3, Avenue 3, Bangsar South, No. 8, Jalan Kerinchi, 59200 Kuala Lumpur, Malaysia -Ph› b∂o tr◊~~H„a Æ¨n sË APP 2010-92 ngµy 29/10/2010~28496804~~0~0~0~0~31663115~10~0~3166312~3166312~APPLIED INTEGRATORS SDN. BHD. Unit 2-1, Level 2, Tower 3, Avenue 3, Bangsar South, No. 8, Jalan Kerinchi, 59200 Kuala Lumpur, Malaysia -Ph› b∂o tr◊~~H„a Æ¨n sË APP 2010-58 ngµy 11/08/2010~85179822~~0~0~0~0"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000086086 27/04/12~45261753~~0~0~0~0~46182050~2~0~923641~923641~Bloomberg Finance L.P. 731 Lexington Avenue New York, NY 10022 USA~~H„a Æ¨n s´ 5601550086 ngµy 09/05/2012~117809738~~130537105~50~10~6526855~124010250~5~0~6200513~12727368</S><S>28180404131~1409020210~27084469549~0~1364262023~2773282233</S><S>X~</S><S>~~~27/02/2013~1~~~</S></S01>"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)
'str2 = "aa316700100157406   062012000000085086n fee~~H„a Æ¨n sË 35978 ngµy 31/03/2011~50906340~~0~0~0~0~56562600~10~0~5656260~5656260~Dell Marketing L.P. PO BOX 676021 C/O Dell USA L.P. Dallas, TX 752671-0621. Ph› b∂o tr◊~~H„a Æ¨n sË 4348488 ngµy 27/01/2011~10091437~~11181647~50~10~559082~10622565~5~0~531128~1090210~Inter Grace Movers (M) Sdn. Bhd. Lot 116, Jlan Semangat, 46300, Petaling Jaya, Selangor, Malaysia. ~~H„a Æ¨n sË EM 1576 ngµy"
'Barcode_Scaned TAX_Utilities_Srv_New.Convert(str2, TCVN, UNICODE)



End Sub

Private Sub Form_Activate()
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
    frmSystem.chkQuetBangKe.Visible = False
    frmSystem.chkQuetBangKe.Value = False

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
        Case comEvReceive                                       ' Received RThreshold # of chars.
            varBuff = MSComm1.Input
            lByte = varBuff
            For i = 0 To UBound(lByte)
                If Chr$(lByte(i)) <> "#" Then
                    strTemp = strTemp & Chr$(lByte(i))
                Else
                    'Debug.Print "a1: " & strTemp
                    'sua loi convert font don vi tinh to khai TTDB
                    'nvanhai sua ngay 01\07\2010
                    
                    strTemp = TAX_Utilities_Srv_New.Convert(strTemp, TCVN, UNICODE)
                    
                    Barcode_Scaned strTemp
                    'Debug.Print "a2: " & strTemp
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
'Input:
'       strBarcode: the scanned barcode string
'Output:
'Return:
'**********************************************
Private Sub Barcode_Scaned(strBarcode As String)
Dim intBarcodeCount As Long, intBarcodeNo As Long
Dim strPrefix As String, strBarcodeCount As String, strData As String
Dim idToKhai As String

On Error GoTo ErrHandle

    'Convert from TCVN to UNICODE format
    strBarcode = TrimString(strBarcode)
    'Debug.Print strBarcode
    strBarcode = Replace(strBarcode, "&", "", 1)
    'strBarcode = TAX_Utilities_Srv_New.Convert(strBarcode, TCVN, UNICODE)
    
    If Left$(strBarcode, 1) <> "0" Then
        'Version 1.2.0 and later
        ' Kiem tra neu version in to khai lon hon max_verion thi khong cho phep nhan
        'If Val(Left$(strBarcode, 3)) > Val(Replace$(APP_VERSION, ".", "")) Then
        If Val(Left$(strBarcode, 3)) > Val(Replace$(HTKK_LAST_VERSION, ".", "")) Then
            'Version tai doanh nghiep lon hon tai co quan thue
            DisplayMessage "0074", msOKOnly, miCriticalError
            Exit Sub
        ElseIf Val(Left$(strBarcode, 3)) < 200 Then ' Truong hop cac to khai thue TNCN duoc in tu phien ban 1.3.1 hieu luc trong nam 2008 thi ko cho nhan
            If Val(Mid$(strBarcode, 4, 2)) = 15 Or Val(Mid$(strBarcode, 4, 2)) = 16 Or Val(Mid$(strBarcode, 4, 2)) = 22 Or Val(Mid$(strBarcode, 4, 2)) = 23 Then
                DisplayMessage "0089", msOKOnly, miCriticalError
                Exit Sub
            End If
        End If
        strPrefix = Left$(strBarcode, 36)
        strBarcodeCount = Right$(strPrefix, 6)
        strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 6)
        
        ' nvhai xu ly nhan BCTC cho phien ban truoc 2.5.0
        ' BCTC in bang HTKK 2.1.0 in tung to rieng biet , tu phien ban 2.5.0 in gop 4 to thanh 1 bo
        ' begin
        ' 1. To khai CDKT phien ban 2.1.0 co ID la 18 (bo 15) -> chuyen thanh ID la 55
        ' 2. To khai CDKT phien ban 2.0.0 co ID la 18 (bo 15) -> chuyen thanh ID la 55
        If Left$(strPrefix, 3) = "210" Or Left$(strPrefix, 3) = "200" Or Left$(strPrefix, 3) = "131" Or Left$(strPrefix, 3) = "130" Then
            idToKhai = Mid(strPrefix, 4, 2)
            If Trim(idToKhai) = "18" Then
                strPrefix = Left$(strPrefix, 3) & "55" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If
            If Trim(idToKhai) = "19" Then
                strPrefix = Left$(strPrefix, 3) & "56" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If
            If Trim(idToKhai) = "20" Then
                strPrefix = Left$(strPrefix, 3) & "57" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If
            If Trim(idToKhai) = "21" Then
                strPrefix = Left$(strPrefix, 3) & "58" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If
        End If
        
        ' end
        
        ' nvhai
        ' 22/07/2010 chan to khai 01/TAIN, 02/TAIN ke khai tu thang 7 va 03/TAIN ke khai tu 2010 in bang HTKK 2.5.2 tro xuong
        If (Trim(Mid(strPrefix, 4, 2)) = "06" And Val(Mid(strPrefix, 19, 2)) > 6 And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Or (Trim(Mid(strPrefix, 4, 2)) = "09" And Val(Mid(strPrefix, 19, 2)) > 6 And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Then
                DisplayMessage "0102", msOKOnly, miInformation
                Exit Sub
        End If
        
        If Trim(Mid(strPrefix, 4, 2)) = "08" And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252 Then
                DisplayMessage "0103", msOKOnly, miInformation
                Exit Sub
        End If
        
        ' end
        
        
        ' Bat dau
        ' To khai 04/TNCN bat dau thu thang 2 se ko nhan nua
        If Left$(strPrefix, 3) = "250" Or Left$(strPrefix, 3) = "210" Then
            idToKhai = Mid(strPrefix, 4, 2)
            ' Neu la to khai 04AB/TNCN thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "39" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "40" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0093", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' To khai 07/TNCN phien ban 2.1.0 bat dau thu thang 2 se ko nhan nua
        If Left$(strPrefix, 3) = "210" Then
            idToKhai = Mid(strPrefix, 4, 2)
            ' Neu la to khai 07/TNCN thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "36" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0097", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        
        ' Lay ID cua to khai de xem co hien thi luon hay ko (Cac to khai quyet toan TNCN)
        If Left(strPrefix, 3) = "250" Then
            ' Kiem tra neu la phien ban 250 va la cac to khai quyet toan TNCN thi phai dat lai tong so ma vach la 1
            idToKhai = Mid(strPrefix, 4, 2)
            If Trim(idToKhai) = "17" Or Trim(idToKhai) = "41" Or Trim(idToKhai) = "42" Or Trim(idToKhai) = "43" Then
                strBarcodeCount = Left(strBarcodeCount, Len(strBarcodeCount) - 1) & "1"
            End If
        End If
        
        ' to khai 06, 02_TNCN_SX, 02_TNCN_BH khong quet bang ke
        'If isIHTKK = True Then
        idToKhai = Mid(strPrefix, 4, 2)
        If Trim(idToKhai) = "59" Or Trim(idToKhai) = "43" Or Trim(idToKhai) = "42" Then
            strBarcodeCount = Left(strBarcodeCount, Len(strBarcodeCount) - 1) & "1"
        End If
        'End If
        If Trim(idToKhai) = "17" Then
            strBarcodeCount = Left(strBarcodeCount, Len(strBarcodeCount) - 1) & "2"
        End If
        
        
        ' Doi voi cac to khai thang quy/TNCN nay da bi thay doi ID giua version 210 va 250
        ' Dat lai cho ID cua 210 dung voi 250 de nhan vao QLT_NTK
        If Left$(strPrefix, 3) = "210" Or Left$(strPrefix, 3) = "200" Then
            idToKhai = Mid(strPrefix, 4, 2)
            ' Neu la to khai 02/TNCN thang cua nam 2009 co ID = 15 thi phai set lai gia tri moi co ID = 53
            If Trim(idToKhai) = "15" And UBound(Split(Mid$(strBarcode, 37), "~")) <> 11 Then
                strPrefix = Left$(strPrefix, 3) & "53" & Mid(strPrefix, 6, Len(strPrefix) - 5)
            End If
            ' Neu la to khai 03/TNCN thang cua nam 2009 co ID = 16 thi phai set lai gia tri moi co ID = 54
            If Trim(idToKhai) = "16" And UBound(Split(Mid$(strBarcode, 37), "~")) <> 11 Then
                strPrefix = Left$(strPrefix, 3) & "54" & Mid(strPrefix, 6, Len(strPrefix) - 5)
            End If
        End If
        
        ' To khai 02/TNCN, 03/TNCN bat dau tu thang 2 se ko nhan theo TT84 nua
        'dhdang sua cho nhan to khai tu ban HTKK 2.5.0
        'date:21/04/2010
        'If (Left$(strPrefix, 3) = "250") Or (Left$(strPrefix, 3) = "210") Then
        If (Left$(strPrefix, 3) = "210") Then
            idToKhai = Mid(strPrefix, 4, 2)
            ' Neu la to khai 02AB/TNCN, 03AB/TNCN  thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "53" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "37" And Val(Mid(strPrefix, 21, 4)) > 2009) _
                Or (Trim(idToKhai) = "54" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "38" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0094", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '07072011 TT28
        ' Khong nhan cac to khai theo mau cua
        idToKhai = Mid(strPrefix, 4, 2)
        If (Val(Left$(strPrefix, 3)) < 300) Then
            If Trim(idToKhai) = "01" Or Trim(idToKhai) = "02" Or Trim(idToKhai) = "04" Or Trim(idToKhai) = "11" Or Trim(idToKhai) = "12" Or Trim(idToKhai) = "46" Or Trim(idToKhai) = "47" Or Trim(idToKhai) = "48" Or Trim(idToKhai) = "49" Or Trim(idToKhai) = "15" Or Trim(idToKhai) = "16" Or Trim(idToKhai) = "50" Or Trim(idToKhai) = "51" _
            Or Trim(idToKhai) = "36" Or Trim(idToKhai) = "70" Or Trim(idToKhai) = "06" Or Trim(idToKhai) = "05" Then
                DisplayMessage "0113", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '06012012 TT28
        ' Khong nhan cac to khai theo mau cu GD2
        If (Val(Left$(strPrefix, 3)) < 310) Then
            If Trim$(idToKhai) = "71" Or Trim$(idToKhai) = "72" Or Trim$(idToKhai) = "73" Or Trim$(idToKhai) = "03" Or Trim$(idToKhai) = "74" Or Trim$(idToKhai) = "75" Or Trim$(idToKhai) = "80" Or Trim$(idToKhai) = "81" Or Trim$(idToKhai) = "82" Or Trim$(idToKhai) = "17" Or Trim$(idToKhai) = "42" Or Trim$(idToKhai) = "43" _
            Or Trim$(idToKhai) = "59" Or Trim$(idToKhai) = "76" Or Trim$(idToKhai) = "41" Or Trim$(idToKhai) = "77" Or Trim$(idToKhai) = "86" Or Trim$(idToKhai) = "87" Or Trim$(idToKhai) = "89" Then
                DisplayMessage "0126", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '17102011 khong nhan cac mau an chi in ra bang HTKK phien ban nho hon 302
        If (Val(Left$(strPrefix, 3)) < 302) Then
            If Trim(idToKhai) = "64" Or Trim(idToKhai) = "65" Or Trim(idToKhai) = "66" Or Trim(idToKhai) = "67" Or Trim(idToKhai) = "68" Then
                DisplayMessage "0122", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        ' Ket thuc
        ' Khong nhan cac to khai 02/TAIN, 05/TNDN
        If Trim(idToKhai) = "08" Or Trim(idToKhai) = "24" Then
                DisplayMessage "0120", msOKOnly, miInformation
                Exit Sub
        End If
        ' end
        
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
            ' hlnam End
            If IsCompleteData(strData) Then
                lblLoading.Visible = False
                lblConnecting.Visible = True
                frmInterfaces.Refresh
                If Not LoadForm(strData) Then
                    StartReceiveForm
                End If
                ' Free memory
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

'
''****************************
''Description: GetCell function get cell string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetCell(pCellID As String, pFirstCell As String, pValue As String, _
'                        bolPrefix As Boolean, bolSuffix As Boolean, pOneCell As Boolean, xmlTemplateCells As MSXML.IXMLDOMNode, lCellCount As Long)
'
'On Error GoTo ErrHandle
'    If (bolPrefix = True And bolSuffix = False) Or (pOneCell = True) Then GetCell = "<Cells>" & vbCrLf
'    If bolSuffix = True And bolPrefix = True And pOneCell = False Then
'        GetCell = GetCell & "</Cells>" & vbCrLf
'        GetCell = GetCell & "<Cells>" & vbCrLf
'    End If
'    GetCell = GetCell & "<Cell CellID=""" & pCellID & """ "
'    If pFirstCell <> vbNullString Then
'        GetCell = GetCell & "FirstCell=""" & pFirstCell & """ Value=""" & pValue & """ "
'    Else
'        GetCell = GetCell & "Value=""" & pValue & """ "
'    End If
'
'    If GetAttribute(xmlTemplateCells.childNodes(lCellCount), "MCT") <> "" Then
'        GetCell = GetCell & "MCT=""" & GetAttribute(xmlTemplateCells.childNodes(lCellCount), "MCT") & """/>" & vbCrLf
'    Else
'        GetCell = GetCell & "/>" & vbCrLf
'    End If
'    If (bolSuffix = True And bolPrefix = False) Or (pOneCell = True) Then GetCell = GetCell & "</Cells>" & vbCrLf
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetCell", Err.Number, Err.Description
'End Function
'
''****************************
''Description: GetCell function get cells string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetCells(ByRef lIncreaseRow As Long, strCellsString As String, xmlTemplateCells As MSXML.IXMLDOMNode) As String
'
'On Error GoTo ErrHandle
'    Dim i As Long, lCol As Long, lRow As Long, lCellCount As Long
'    Dim strCellArr() As String
'    Dim strCellID As String, strFirstCell As String
'    Dim bolPrefix As Boolean, bolSuffix As Boolean
'    Dim blnBeginOfData As Boolean, blnEndOfData As Boolean
'    Dim bolOneCell As Boolean
'
'    strCellArr = Split(strCellsString, "~")
'
'    blnBeginOfData = True
'    blnEndOfData = False
'
'    i = 0
'    While i <= UBound(strCellArr) Or lCellCount < xmlTemplateCells.childNodes.length
'        If UBound(strCellArr) = 0 Then
'            bolOneCell = True
'        Else
'            bolOneCell = False
'        End If
'        bolPrefix = blnBeginOfData
'        bolSuffix = blnEndOfData
'
'        strFirstCell = GetAttribute(xmlTemplateCells.childNodes(lCellCount), "FirstCell")
'        If lCellCount >= xmlTemplateCells.childNodes.length Then
'            lCellCount = 0
'            lIncreaseRow = lIncreaseRow + GetDynRowCount(fpSpread1, xmlTemplateCells)
'            strFirstCell = "1"
'            bolPrefix = True
'            bolSuffix = True
'        End If
'        ParserCellID fpSpread1, GetAttribute(xmlTemplateCells.childNodes(lCellCount), "CellID"), lCol, lRow
'        lRow = lRow + lIncreaseRow
'        strCellID = GetCellID(fpSpread1, lCol, lRow)
'
'        If GetAttribute(xmlTemplateCells.childNodes(lCellCount), "Encode") <> "0" And i <= UBound(strCellArr) Then
'            GetCells = GetCells & GetCell(strCellID, strFirstCell, Replace(strCellArr(i), Chr$(20), "~"), bolPrefix, bolSuffix, bolOneCell, xmlTemplateCells, lCellCount)
'        Else
'            GetCells = GetCells & GetCell(strCellID, strFirstCell, "", bolPrefix, bolSuffix, bolOneCell, xmlTemplateCells, lCellCount)
'            i = i - 1
'        End If
'        lCellCount = lCellCount + 1
'
'        If i <= 0 Then
'            blnBeginOfData = False
'        End If
'
'        If i >= UBound(strCellArr) - 1 And lCellCount = xmlTemplateCells.childNodes.length - 1 Then
'            blnEndOfData = True
'        End If
'
'        i = i + 1
'    Wend
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetCells", Err.Number, Err.Description
'End Function
'
''****************************
''Description: GetCell function get Section string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetSection(ByRef lIncreaseRow As Long, xmlNodeDataSection As MSXML.IXMLDOMNode, xmlNodeTemplateSection As MSXML.IXMLDOMNode) As String
'    Dim xmlNodeCells As MSXML.IXMLDOMNode
'    Dim strDynamic As String, strMaxRow As String
'
'On Error GoTo ErrHandle
'    strDynamic = GetAttribute(xmlNodeTemplateSection.parentNode, "Dynamic")
'    strMaxRow = GetAttribute(xmlNodeTemplateSection.parentNode, "MaxRows")
'
'    GetSection = "<Section Dynamic=""" & strDynamic & """ MaxRows=""" & strMaxRow & """>" & vbCrLf
'
'    If Not xmlNodeDataSection Is Nothing Then
'        GetSection = GetSection & GetCells(lIncreaseRow, xmlNodeDataSection.Text, xmlNodeTemplateSection)
'    Else
'        GetSection = GetSection & GetCells(lIncreaseRow, " ", xmlNodeTemplateSection)
'    End If
'    GetSection = GetSection & "</Section>" & vbCrLf
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetSection", Err.Number, Err.Description
'End Function
'
''****************************
''Description: GetCell function get Sections string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetSections(xmlNodeDataSections As MSXML.IXMLDOMNode, xmlNodeTemplateSections As MSXML.IXMLDOMNode) As String ', rsTaxInfor As ADODB.Recordset) As String
'    Dim i As Long, j As Long, lIncreaseRow As Long
'
'On Error GoTo ErrHandle
'    GetSections = "<!-- edited with XML Spy v4.1 U (http://www.xmlspy.com) by tuyends (FSS) -->" & vbCrLf
'    GetSections = GetSections & "<!DOCTYPE Sections SYSTEM ""Schema.dtd"">" & vbCrLf
'    GetSections = GetSections & "<Sections"
'
'    'Get all attr of  Section node
'    For i = 0 To xmlNodeTemplateSections.Attributes.length - 1
'        GetSections = GetSections & " " & xmlNodeTemplateSections.Attributes(i).xml
'    Next i
'    GetSections = GetSections & "> " & vbCrLf
'
'    lIncreaseRow = 0
'
'    GetSections = GetSections & xmlNodeTemplateSections.childNodes(0).xml & vbCrLf
'
'    For i = 1 To xmlNodeTemplateSections.childNodes.length - 1
'        If xmlNodeTemplateSections.childNodes(i).baseName = "Section" Then
'            GetSections = GetSections & GetSection(lIncreaseRow, xmlNodeDataSections.childNodes(i - 1).childNodes(0), xmlNodeTemplateSections.childNodes(i).childNodes(0))
'        End If
'    Next
'    GetSections = GetSections & "</Sections>" & vbCrLf
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetSections", Err.Number, Err.Description
'End Function
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
    ' Khong check doi voi cac BCTC
    idToKhaiCheck = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    If (idToKhaiCheck >= 24 And idToKhaiCheck <= 35) Or (idToKhaiCheck >= 55 And idToKhaiCheck <= 58) Or (idToKhaiCheck >= 18 And idToKhaiCheck <= 21) Or idToKhaiCheck = 69 Then
        isSheetTk = False
    End If
    
    lElementsNo = GetElementsNo(xmlSectionTemplate.childNodes(0))
    'Get array of data units
    arrStrValue = Split(xmlSectionData.Text, "~")
    ' Lay ve so chi tieu cua chuoi ma vach
    lDataNo = UBound(arrStrValue)
    If lDataNo = -1 Then
        lDataNo = 0
    End If
    ' End
    
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
        ElseIf idToKhaiCheck = 11 Then
             If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
        ElseIf idToKhaiCheck = 12 Then
             If ((lDataNo + 1 > lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 6)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 6)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
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

Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
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
    Dim strDataRestore As String, strFileName As String
    Dim lIndex As Long, lCtrl As Long, arrStrData() As String
    Dim xmlData As New MSXML.DOMDocument, xmlTemplate As New MSXML.DOMDocument
    Dim fso As New FileSystemObject, tstFile As TextStream
    Dim blnValidData As Boolean
    
On Error GoTo ErrHandle
    arrStrData = GetSheetDatas(strBarcodeData)
    
    If UBound(arrStrData) < TAX_Utilities_Srv_New.NodeValidity.childNodes.length Then
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
        xmlTemplate.Load GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "TemplateFolder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & ".xml"
        
        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_00" & TAX_Utilities_Srv_New.Year & ".xml"
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
            TAX_Utilities_Srv_New.NodeMenu = xmlNode
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
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        TAX_Utilities_Srv_New.Month = Left$(strValue, 2)
        TAX_Utilities_Srv_New.ThreeMonths = ""
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = 1 Then
        TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
        TAX_Utilities_Srv_New.Month = ""
    End If
    
    TAX_Utilities_Srv_New.Year = Right$(strValue, 4)
    
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
    
    ' Xu ly cho cac to khai lan phat sinh
    Dim arrCT() As String
    Dim strTemp As String
    
On Error GoTo ErrHandle
    
    TAX_Utilities_Srv_New.Month = ""
    TAX_Utilities_Srv_New.ThreeMonths = ""
    TAX_Utilities_Srv_New.Year = ""
    TAX_Utilities_Srv_New.FinanceStartDate = ""
    
    
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
'
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
'    'Version 2.5.0
'        'Get version of application
'        lblVersion.caption = "2.5.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "251" Then
'    'Version 2.5.0
'        'Get version of application
'        lblVersion.caption = "2.5.1"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "252" Then
'    'Version 2.5.2
'    'Get version of application
'    lblVersion.caption = "2.5.2"
'    strTaxReportVersion = Left$(strData, 3)
'    strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "253" Then
'        'Version 2.5.2
'        'Get version of application
'        lblVersion.caption = "2.5.3"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    End If
    
    ' 03122010 - sua lai doan lay version cua ung dung in ma vach
    strTaxReportVersion = Left$(strData, 3)
    strData = Mid$(strData, 4)
    lblVersion.caption = Left$(strTaxReportVersion, 1) & "." & Mid$(strTaxReportVersion, 2, 1) & "." & Right$(strTaxReportVersion, 1)
    ' end doan lay version
    
    'Get info of barcode string --25 characters
    strTaxReportInfo = Left$(strData, 21)
    
    If xmlSQL.url = "" Then
        xmlSQL.Load App.path & "\SQL.xml"
    End If
    
    'Get Tax id
    strTaxID = Trim(Mid$(strTaxReportInfo, 3, 13))
    If Len(strTaxID) = 13 Then
        If Trim(Right(strTaxID, 3)) = "000" Then
            strTaxID = Mid$(strTaxID, 1, 10)
        Else
            strTaxID = Mid$(strTaxID, 1, 10) & "-" & Mid$(strTaxID, 11, 13)
        End If
    End If
    
    'Connect DB and get informations
    ' nvhai
    ' Yeu cau nhan BCTC khong can check co quan thue cua user dang nhap
    ' 09-06-2010
    ' begin
    Dim strIDBCTC As String
    strIDBCTC = Left$(strTaxReportInfo, 2)
     If (Val(strIDBCTC) = 24 Or Val(strIDBCTC) = 25 Or Val(strIDBCTC) = 26 Or Val(strIDBCTC) = 27 Or Val(strIDBCTC) = 28 Or Val(strIDBCTC) = 29 _
            Or Val(strIDBCTC) = 30 Or Val(strIDBCTC) = 31 Or Val(strIDBCTC) = 32 Or Val(strIDBCTC) = 33 Or Val(strIDBCTC) = 34 Or Val(strIDBCTC) = 35 _
            Or Val(strIDBCTC) = 55 Or Val(strIDBCTC) = 56 Or Val(strIDBCTC) = 57 Or Val(strIDBCTC) = 58 Or Val(strIDBCTC) = 18 Or Val(strIDBCTC) = 19 _
            Or Val(strIDBCTC) = 20 Or Val(strIDBCTC) = 21 Or Val(strIDBCTC) = 69) Then
        Set rsTaxInfor = GetTaxInfoBCTC(strTaxID, blnConnected)
    Else
        Set rsTaxInfor = GetTaxInfo(strTaxID, blnConnected)
    End If
    ' end
    
     'Connect DB fail
    If Not blnConnected Then _
        Exit Function
    
    If (rsTaxInfor Is Nothing Or rsTaxInfor.Fields.Count = 0) And Len(strTaxID) = 14 Then
        strTaxID = Replace(strTaxID, "-", " ")
        Set rsTaxInfor = GetTaxInfo(strTaxID, blnConnected)
    End If
    'Tax id is not exist
    If rsTaxInfor Is Nothing Or rsTaxInfor.Fields.Count = 0 Then
        InitParameters = False
        MessageBox "0041", msOKOnly, miCriticalError
        Exit Function
    End If
    
    'Tax id is closed
    
    If rsTaxInfor.Fields(0) = "01" Then
        InitParameters = False
        MessageBox "0042", msOKOnly, miCriticalError
        Exit Function
    End If
    
    'DTNT chuyen di noi khac
    ' reset trang thai kiem tra
    checkTT = 0
    If rsTaxInfor.Fields(0) = "02" Then
        InitParameters = False
        'checkTT = 1  ' Trang thai chuyen di noi khac
        MessageBox "0043", msOKOnly, miCriticalError
        Exit Function
    End If
    
    'DTNT hien dang bi mat tich
    If rsTaxInfor.Fields(0) = "03" Then
        'InitParameters = False
        checkTT = 2  ' Trang thai DTNT mat tich
        'MessageBox "0087", msOKOnly, miCriticalError
        'Exit Function
    End If

    'DTNT hien dang tam dung kinh doanh co thoi han
    'sua theo mail cua ptly yeu cau khong kiem tra MST tam ngung kinh doanh van duoc nop
    'ngay 01.07.2010
    
'    If rsTaxInfor.Fields(0) = "05" Then
'        InitParameters = False
'        MessageBox "0088", msOKOnly, miCriticalError
'        Exit Function
'    End If

    strMST = CStr(rsTaxInfor.Fields(1))
    
    If InStr(1, strData, "<S") < 35 Then
        'Ver 1.0
        ' Get NgayTaiChinh and ThangTaiChinh
        iNgayTaiChinh = 0
        iThangTaiChinh = 0
    Else
        'Ver 1.1.0 and later
        ' Get NgayTaiChinh and ThangTaiChinh
        strTempDate = Mid$(strData, 22, 5)
        iNgayTaiChinh = GetNgayTaiChinh(strTempDate)
        iThangTaiChinh = GetThangTaiChinh(strTempDate)
        TAX_Utilities_Srv_New.FinanceStartDate = strTempDate
    End If
    
    strID = Left$(strTaxReportInfo, 2)
    SetNodeMenu strID
    SetPeriod Right$(strTaxReportInfo, 6)
    TAX_Utilities_Srv_New.NodeValidity = GetValidityNode
    
    '*******************************
'Date: 13/02/2006
    'Gan gia tri tu ngay, den ngay.
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" Then
        TAX_Utilities_Srv_New.FirstDay = Mid$(strData, 37, 10)
        TAX_Utilities_Srv_New.LastDay = Mid$(strData, 47, 10)
    End If
'*******************************
'*******************************
'Date: 16/02/2006
    'Danh sach to khai can kiem tra ngay bat dau nam tai chinh
    On Error GoTo ThamSoErrHandle
    
    Set rsParams = clsDAO.Execute("select gia_tri from rcv_thamso where ten ='LOAI_TK_TAICHINH'")
    
    On Error GoTo ErrHandle
    'Kiem tra ngay bat dau nam tai chinh doi voi cac loai to
    '   khai co kiem tra ngay bat dau nam tai chinh
    If InStr(1, "," & rsParams.Fields(0) & ",", "," & GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") & ",") <> 0 Then
        If Not IsNull(rsTaxInfor.Fields("ngay_tchinh")) Then
            If Mid$(rsTaxInfor("ngay_tchinh"), 1, 5) <> Mid$(strTempDate, 1, 5) Then
                DisplayMessage "0065", msOKOnly, miCriticalError
                Exit Function
            End If
            'Kiem tra ngay bat dau kinh doanh
        Else 'Trong DB chua co gia tri ngay bat dau kinh doanh
            DisplayMessage "0066", msOKOnly, miCriticalError
            Exit Function
        End If
    End If
    
    'To khai ke khai tu ngay ... den ngay ...
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1" Then
        If Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) <> 17 Then
            If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" Then
                If Mid$(rsTaxInfor("ngay_tchinh"), 1, 5) <> Mid$(TAX_Utilities_Srv_New.FirstDay, 1, 5) Then
                   'Tu ngay phai bang ngay bat dau nam tai chinh
                   ' hoac ngay bat dau kinh doanh
                    DisplayMessage "0068", msOKOnly, miCriticalError
                    Exit Function
                End If
                ''Ky ke khai lon hon ngay bat dau kinh doanh
                'If CInt(Mid$(rsTaxInfor("ngay_kdoanh"), 7, 4)) > CInt(Mid$(TAX_Utilities_Srv_New.FirstDay, 7, 4)) Then
                '    DisplayMessage "0069", msOKOnly, miCriticalError
                '    Exit Function
                'End If
            End If
        End If
    End If
    
    'Kiem tra cach thuc tinh ky ke khai la tinh theo nam duong lich hay nam tai chinh
    On Error GoTo ThamSoErrHandle
    
    Set rsParams = clsDAO.Execute("select gia_tri from rcv_thamso where ten ='THEO_NAM_TAICHINH'")
    blnTinhTheoNamTaiChinh = IIf(rsParams.Fields(0) = 0 Or IsNull(rsParams.Fields(0)), False, True)
    
    ' Doi voi to khai TNDN quy 01A/TNDN va 01B/TNDN thi phai lay dung theo ky ke khai cua nam tai chinh
    ' Vi du MST co ngay bat dau Nam tai chinh la 01/04/2009 thi quy 1 se bat dau la 01/04/2009 va quy 4 se bat dau la ngay 01/01/2010
    ' Dat lai blnTinhTheoNamTaiChinh = True => Ham GetNgayDauQuy se tra lai dung ket qua
    ' ID = 11 => To khai 01A/TNDN, ID = 12 => To khai 01B/TNDN
    
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1" Then
        If Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) = 11 Or Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) = 12 Then
            blnTinhTheoNamTaiChinh = True
        End If
    End If
    
    
    On Error GoTo ErrHandle
    'Gan gia tri ngay dau ky
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        dNgayDauKy = DateSerial(CInt(TAX_Utilities_Srv_New.Year), CInt(TAX_Utilities_Srv_New.Month), 1)
        dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
        dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
        dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
        dNgayDauKy = GetNgayDauNam(CInt(TAX_Utilities_Srv_New.Year), iThangTaiChinh, iNgayTaiChinh)
        dNgayCuoiKy = DateAdd("m", 12, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    End If
'*******************************
'*******************************
'Date: 11/01/2006
    'Check validity of start date.
    
    If InStr(1, strData, "<S") < 35 Then
        'Ver 1.0
        strTempDate = Mid$(strData, 22, 8)
        strValidDate = GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "StartDate")
        If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))) _
                , DateSerial(CInt(Mid$(strTempDate, 5, 4)), CInt(Mid$(strTempDate, 3, 2)), CInt(Mid$(strTempDate, 1, 2)))) = 0 Then
            DisplayMessage "0064", msOKOnly, miInformation
            Exit Function
        End If
    '*******************************
        'Get main content
        strData = Mid$(strData, 30)
    Else
        'Ver 1.1.0 and later
        strTempDate = Mid$(strData, 27, 10)
        
        ' Neu la to khai 02A-02B/TNCN, 03A-03B/TNCN ke khai trong phien ban 2.5.0 thi set lai ngay bat dau hieu luc la 01/01/2010
        ' Sau khi co phien ban 2.5.2 se bat rang buoc cac dieu kien ve phien ban tu 2.5.2 thi bo doan check nay di
        If (Trim(strID) = "15") Or (Trim(strID) = "16") Or (Trim(strID) = "50") Or (Trim(strID) = "51") Or (Trim(strID) = "46") Or (Trim(strID) = "47") _
        Or (Trim(strID) = "48") Or (Trim(strID) = "49") Or (Trim(strID) = "36") Then
            strTempDate = "01/01/2010"
        End If
        ' End
        
        strValidDate = GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "StartDate")
        If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))) _
                , DateSerial(CInt(Mid$(strTempDate, 7, 4)), CInt(Mid$(strTempDate, 4, 2)), CInt(Mid$(strTempDate, 1, 2)))) = 0 Then
            ' Truong hop dang bi map nham ID cua cac BCTC giua version 2.5.1 ngay 17/03/1020 voi ID cua version 2.1.0
            ' Thi ko hien thi message nay ma map lai version, id cho dung voi to khai 2.5.1 roi quet lai
            ' Begin
            If Trim(strID) = "55" Or Trim(strID) = "56" Or Trim(strID) = "57" Or Trim(strID) = "58" Then
                Exit Function
            ' end
            ElseIf Trim(strID) = "70" Then
            Else
                DisplayMessage "0064", msOKOnly, miInformation
                Exit Function
            End If
        End If
        
    '*******************************
        'Get main content
        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") <> "0" Then
            If Trim(strID) = "70" Or Trim(strID) = "81" Then
                strData = Mid$(strData, 37)
            Else
                strData = Mid$(strData, 57)
            End If
        Else
            strData = Mid$(strData, 37)
        End If
    End If
    'RestoreDataFile (strData)
    If Not RestoreDataFile(strData) Then  ', rsTaxInfor
        'So chi tieu tren ma vach nhieu hon so chi tieu tren to khai
        If checkSoCT = 1 Then
            MessageBox "0104", msOKOnly, miCriticalError
            Exit Function
        End If
        ' So chi tieu tren ma vach it hon so chi tieu tren to khai
        If checkSoCT = 2 Then
            MessageBox "0105", msOKOnly, miCriticalError
            Exit Function
        End If
        ' Kiem tra cac to khai co so dong dong (chi kiem tra duoc khac chu khong phan biet duoc truong hop thieu hoac thua)
        If checkSoCT = 3 Then
            MessageBox "0106", msOKOnly, miCriticalError
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
    '***********************************
    'Date 26/05/06
    
    'Gan thong tin Header vao mang
    If Not GetHeaderData(rsTaxInfor, arrStrHeaderData) Then
        DisplayMessage "0080", msOKOnly, miCriticalError
        Exit Function
    End If
    
    ' Truoc khi lay thong tin ve tep, chuyen cac ma quy uoc cho to khai cu ve ma quy uoc cho to khai moi
    ' Lay thong tin ma so tep va so thu tu to khai.
'    If Not GetThongTinTep(changeMaToKhai(strID), arrStrHeaderData) Then
'        DisplayMessage "0079", msOKOnly, miCriticalError
'        Exit Function
'    End If
    
    ' Lay so thu tu cua to khai da dua vao RCV_TKHAI_HDR
    ' So thu tu nay phai lay theo cung Nguoi nop thue, ky ke khai, va cung loai to khai
    ' An chi
    If Val(strID) >= 64 And Val(strID) <= 68 Then
            ' An chi
            If Not getSoTTTK_AC(changeMaToKhai(strID), arrStrHeaderData, strData) Then
                DisplayMessage "0079", msOKOnly, miCriticalError
                Exit Function
            End If
    Else
        ' cac to khai binh thuong
        isTKLanPS = False
        isTKThang = False
        ngayPS = ""
        ' 02/TNDN
        If Val(strID) = 73 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(32)) <> "" Then
                ngayPS = arrCT(32)
                isTKLanPS = True
            End If
        End If
        ' 01/TTDB
        If Val(strID) = 5 Then
            strTemp = Left$(strData, InStr(1, strData, "</S></S01>") + 9)
            arrCT = Split(strTemp, "~")
            If Trim(arrCT(UBound(arrCT) - 1)) <> "" Then
                ngayPS = arrCT(UBound(arrCT) - 1)
                isTKLanPS = True
            End If
        End If

        ' Xy ly to khai 08, 08A/TNCN
        If Val(strID) = 74 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(2)) <> "" Then
                TuNgay = arrCT(2)
                DenNgay = arrCT(3)
                isTKThang = True
            End If
            
        End If
' 08A/TNCN
        If Val(strID) = 75 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(1)) <> "" Then
                TuNgay = Right$(arrCT(0), 7)
                DenNgay = arrCT(1)
                isTKThang = True
            End If
            
        End If
        
        ' To khai 01/NTNN
        If Val(strID) = 70 Then
            strTemp = Left$(strData, InStr(1, strData, "</S></S01>") + 9)
            arrCT = Split(strTemp, "~")
            If Trim(arrCT(UBound(arrCT))) <> "" And Left$(Trim(arrCT(UBound(arrCT))), 10) <> "</S></S01>" Then
                ngayPS = Left$(Trim(arrCT(UBound(arrCT))), 10)
                isTKLanPS = True
            End If
        End If
        ' To khai 03/NTNN
        If Val(strID) = 81 Then
            strTemp = Left$(strData, InStr(1, strData, "</S></S01>") + 9)
            arrCT = Split(strTemp, "~")
            If Trim(arrCT(UBound(arrCT))) <> "" And Left$(Trim(arrCT(UBound(arrCT))), 10) <> "</S></S01>" Then
                ngayPS = Left$(Trim(arrCT(UBound(arrCT))), 10)
                isTKLanPS = True
            End If
        End If
        
            
        If Not getSoTTTK(changeMaToKhai(strID), arrStrHeaderData) Then
            DisplayMessage "0079", msOKOnly, miCriticalError
            Exit Function
       End If
       
        ' 18122012
        ' to khai lan phat sinh trog ngay chi nhan 1 to khai
        If (Val(strID) = 70 Or Val(strID) = 73 Or Val(strID) = 81 Or Val(strID) = 5) And isTKLanPS = True Then
            If isToKhaiPsDaNhanTN = True Then
                DisplayMessage "0129", msOKOnly, miCriticalError
                Exit Function
            End If
        End If
            
    End If
    '***********************************
    
    ' Kiem tra to khai ton tai theo mau cu QLT
    isTKDA30 = isDA30(strID, arrStrHeaderData)
        
        
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
                .SheetName = GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(.Sheet - 1), "Caption")
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

    Dim rsHeaderData As ADODB.Recordset
    Dim arrStrHeaderData() As String
    Dim LoaiTk As String
    Dim strMST As String
    
    Dim dsTK_DLT As String
    
    Dim blnDLConnected As Boolean
    Dim strTaxDLID As String
    Dim rsTaxDLInfor As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    frmSystem.MousePointer = vbHourglass
    
    LoaiTk = Mid(strData, 4, 2)
    
    'If InitParameters(strData, rsHeaderData) = False Then
    If InitParameters(strData, arrStrHeaderData) = False Then
        ' Truong hop bi map ID sai BCTC giua phien ban 2.5.1 ban ngay 17/03/2010 voi phien ban 2.1.0
        ' Thi conver lai cho chuan ID BCTC va Init lai cac Parameter
        If Trim(LoaiTk) = "55" Then
            InitParameters "25118" & Mid(strData, 6), arrStrHeaderData
        ElseIf Trim(LoaiTk) = "56" Then
            InitParameters "25119" & Mid(strData, 6), arrStrHeaderData
        ElseIf Trim(LoaiTk) = "57" Then
            InitParameters "25120" & Mid(strData, 6), arrStrHeaderData
        ElseIf Trim(LoaiTk) = "58" Then
            InitParameters "25121" & Mid(strData, 6), arrStrHeaderData
        Else
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
            Exit Function
        End If
    End If
          
    mOnLoad = True
    fpSpread1.EventEnabled(EventAllEvents) = False
    LoadTemplate fpSpread1
    SetupSpread
    FormatGrid
    'LoadInitFiles
    
    
        ' Set cac thong tin cua DL thue
     'Get Tax id
    strMST = Trim(Mid$(strTaxReportInfo, 3, 13))
    
    strTaxDLID = Mid(strData, InStr(1, strData, "<S>") + 3, InStr(1, strData, "</S>") - InStr(1, strData, "<S>") - 3)
    
    
    If Len(Trim(strMST)) = 13 Then
        strMST = Left(strMST, 10) & "-" & Right(strMST, 3)
    End If
    strMaSoThue = strMST
    
    
    If Len(Trim(strTaxDLID)) = 13 Then
        strTaxDLID = Left(strTaxDLID, 10) & "-" & Right(strTaxDLID, 3)
    End If
    strMaDaiLyThue = strTaxDLID
    
    Set rsTaxDLInfor = GetTaxDLInfo(strMST, strTaxDLID, blnDLConnected)
        
    
    If Trim(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "Class")) <> vbNullString Then
        Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "Class"))
        Set objTaxBusiness.fps = fpSpread1
        'objTaxBusiness.strMaSoTep = strMaSoTep
        objTaxBusiness.strPhongXuLy = strMaPhongXuLy
        objTaxBusiness.strNgayNhanToKhai = strNgayNhanToKhai
        objTaxBusiness.strNguoiSuDung = strUserID
        
            ' set thong tin DL thue
        ' danh sach cac to khai se set thong tin dai ly thue TT28
        dsTK_DLT = "~01~02~03~04~05~06~11~12~46~47~48~49~15~16~50~51~36~70~71~72~73~74~75~80~81~82~77~86~87~89~17~42~43~59~76~41~"
'        If Trim(LoaiTk) = "01" Or Trim(LoaiTk) = "02" Or Trim(LoaiTk) = "04" Or Trim(LoaiTk) = "05" Or Trim(LoaiTk) = "06" Or Trim(LoaiTk) = "11" _
'        Or Trim(LoaiTk) = "12" Or Trim(LoaiTk) = "46" Or Trim(LoaiTk) = "47" Or Trim(LoaiTk) = "48" Or Trim(LoaiTk) = "49" Or Trim(LoaiTk) = "15" _
'        Or Trim(LoaiTk) = "16" Or Trim(LoaiTk) = "50" Or Trim(LoaiTk) = "51" Or Trim(LoaiTk) = "36" Or Trim(LoaiTk) = "70" Or Trim(LoaiTk) = "71" _
'        Or Trim(LoaiTk) = "72" Then
         If InStr(1, dsTK_DLT, "~" & Trim(LoaiTk) & "~", vbTextCompare) > 0 Then
            If Trim(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "Class")) <> vbNullString Then
                If Not (rsTaxDLInfor Is Nothing Or rsTaxDLInfor.Fields.Count = 0) Then
                    If Not objTaxBusiness Is Nothing Then
                        objTaxBusiness.strTenDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(0).Value), "", rsTaxDLInfor.Fields(0).Value), TCVN, UNICODE)
                        objTaxBusiness.strDiaChiDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(1).Value), "", rsTaxDLInfor.Fields(1).Value), TCVN, UNICODE)
                         objTaxBusiness.strDienThoaiDL = IIf(IsNull(rsTaxDLInfor.Fields(2).Value), "", rsTaxDLInfor.Fields(2).Value)
                        objTaxBusiness.strFaxDL = IIf(IsNull(rsTaxDLInfor.Fields(3).Value), "", rsTaxDLInfor.Fields(3).Value)
                        objTaxBusiness.strEmailDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(4).Value), "", rsTaxDLInfor.Fields(4).Value), TCVN, UNICODE)
                        objTaxBusiness.strSoHopDongDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(5).Value), "", rsTaxDLInfor.Fields(5).Value), TCVN, UNICODE)
                        objTaxBusiness.strNgayHopDongDL = IIf(IsNull(rsTaxDLInfor.Fields(6).Value), "", rsTaxDLInfor.Fields(6).Value)
                    End If
                End If
            End If
        End If

        
        ' set ngay bat dau nam tai chinh cho to khai 01ATNDN va 01BTNDN
        If LoaiTk = "11" Or LoaiTk = "12" Or LoaiTk = "03" Then
            objTaxBusiness.dNgayTC = dNgayDauKy
        End If
        ' end
        If Not objTaxBusiness.Prepared1 Then Exit Function
    End If
            
    SetupData fpSpread1

    If Not objTaxBusiness Is Nothing Then
        If Not objTaxBusiness.Prepared2(rsPXL) Then Exit Function
    End If
    
    ' set ma CQT
    If Not objTaxBusiness Is Nothing Then
        If Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) >= 64 And Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) <= 68 Then
            objTaxBusiness.strMaCQT = strTaxOfficeId
            ' lay ma phong quan ly
            'Get Tax id
            strMST = Trim(Mid$(Left$(strData, 21), 6, 13))
            If Len(strMST) = 13 Then
                strMST = Mid$(strMST, 1, 10) & "-" & Mid$(strMST, 11, 13)
            End If
            GetPhongQuanLy (strMST)
            objTaxBusiness.strMaPQL = strMaPhongQuanLy
            objTaxBusiness.strTenPQL = strTenPhongQuanLy
        End If
    End If
    
    ' Set Phong quan ly
    If Not objTaxBusiness Is Nothing Then
        If Val(LoaiTk) = 70 Or Val(LoaiTk) = 71 Or Val(LoaiTk) = 72 Or Val(LoaiTk) = 73 Or Val(LoaiTk) = 74 Or Val(LoaiTk) = 77 Or Val(LoaiTk) = 3 Or Val(LoaiTk) = 75 _
        Or Val(LoaiTk) = 80 Or Val(LoaiTk) = 81 Or Val(LoaiTk) = 82 Or Val(LoaiTk) = 86 Or Val(LoaiTk) = 87 Or Val(LoaiTk) = 89 Or Val(LoaiTk) = 17 Or Val(LoaiTk) = 42 Or Val(LoaiTk) = 43 _
        Or Val(LoaiTk) = 59 Or Val(LoaiTk) = 76 Or Val(LoaiTk) = 41 Then
            ' lay ma phong quan ly
            'Get Tax id
            strMST = Trim(Mid$(Left$(strData, 21), 6, 13))
            If Len(strMST) = 13 Then
                strMST = Mid$(strMST, 1, 10) & "-" & Mid$(strMST, 11, 13)
            End If
            GetPhongQuanLy (strMST)
            objTaxBusiness.strMaPQL = strMaPhongQuanLy
            objTaxBusiness.strTenPQL = strTenPhongQuanLy
        End If
    End If
    
    

    'Setup header data
    'SetupHeaderData rsHeaderData
    SetupHeaderData arrStrHeaderData

    If Not objTaxBusiness Is Nothing Then
        If Not objTaxBusiness.Prepared3 Then Exit Function
    End If
    fpSpread1.EventEnabled(EventAllEvents) = True
    cmdClear.Enabled = True
    cmdSave.Enabled = True
    cmdViewNow.Enabled = False
    fpSpread1.Visible = True
    
    lblLabelVersion.Left = 4460
    lblVersion.Left = 8650
    
    'If CLng(Replace$(strTaxReportVersion, ".", "")) < CLng(Replace$(APP_VERSION, ".", "")) Then
    If CLng(Replace$(strTaxReportVersion, ".", "")) < CLng(Replace$(HTKK_LAST_VERSION, ".", "")) Then
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
    ' Loai to khai la GTGT khau tru mau 01/GTGT thi cho hien check quet du lieu bang ke
    If (verToKhai = 0 And LoaiTk = "01") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Or _
                TAX_Utilities_Srv_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.Value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    ' Loai to khai la GTGT hoa hong dai ly mau 02/GTGT thi cho hien check quet du lieu bang ke
    ElseIf (verToKhai = 0 And LoaiTk = "02") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.Value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    ' Loai to khai la Quyet toan thue TNCN mau 04/TNCN thi cho hien check quet du lieu bang ke
    ElseIf (verToKhai = 0 And LoaiTk = "17") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.Value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    Else
        frmSystem.chkQuetBangKe.Value = False
        frmSystem.chkQuetBangKe.Visible = False
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
    
    '*********************************
    'Date: 10/04/06
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
    For Each xmlNode In TAX_Utilities_Srv_New.NodeValidity.childNodes
        SetAttribute xmlNode, "Active", "0"
    Next
    
    ReDim arrStrData(0)
    
    For Each xmlNode In TAX_Utilities_Srv_New.NodeValidity.childNodes
         Dim i As Integer
        intLoc1 = InStr(1, strBarcodeData, "<S" & GetAttribute(xmlNode, "ID") & ">")
        i = Len(GetAttribute(xmlNode, "ID"))
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
        If UBound(arrStrData) < TAX_Utilities_Srv_New.NodeValidity.childNodes.length Then
            ReDim Preserve arrStrData(TAX_Utilities_Srv_New.NodeValidity.childNodes.length)
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
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    

    ' Get SQL statement from DOM
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlMST")
    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
    
    Set rsReturn = clsDAO.Execute(strSQL)
    
    Set GetTaxInfo = rsReturn
    
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

' Lay thong tin DL thue 05072011
Private Function GetTaxDLInfo(ByVal strTaxIDString As String, ByVal strTaxIDDLString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    

    ' Get SQL statement from DOM
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlMSTDL")
    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
    strSQL = Replace(strSQL, "ma_dai_ly", "'" & strTaxIDDLString & "'")
    
    Set rsReturn = clsDAO.Execute(strSQL)
    
    Set GetTaxDLInfo = rsReturn
    
    Set rsReturn = Nothing
    
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetTaxDLInfo", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function


' nvhai
' Lay ve thong tin cua doi tuong NT nhung khong check CQT
' Phuc vu viec quet cac BCTC cua chi Cuc nhung quet tren Cuc
' 09-06-2010
' begin
'Input:
'       strTaxIDString: Id string
'Output:
'Return: a Recordset contain query data.
'****************************
Private Function GetTaxInfoBCTC(ByVal strTaxIDString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    

    strSQL = "SELECT trang_thai,tin, ten_dtnt, dia_chi,dien_thoai,fax, email,to_char(sysdate,'mm/rrrr') ky_lapbo "
    strSQL = strSQL & ", to_char(sysdate,'dd/mm/rrrr') ngay_nop,to_char(sysdate,'dd/mm/rrrr') ngay_nhap,to_char(ngay_tchinh,'dd/mm') ngay_tchinh,to_char(ngay_kdoanh,'dd/mm/yyyy')   ngay_kdoanh "
    strSQL = strSQL & "FROM rcv_v_dtnt where tin =" & "'" & strTaxIDString & "'"
    
    
    
    Set rsReturn = clsDAO.Execute(strSQL)
    
    Set GetTaxInfoBCTC = rsReturn
    
    Set rsReturn = Nothing
    
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetTaxInfoBCTC", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

'end

Private Sub GetPhongQuanLy(ByVal strTaxIDString As String)
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    Dim strPQLString As String
On Error GoTo ErrHandle

    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    ' Lay ma phong quan ly cua MST
    strSQL = "select ma_phong from rcv_v_dtnt where tin = '" & Trim(strTaxIDString) & "'"
    Set rsReturn = clsDAO.Execute(strSQL)
    If Not (rsReturn Is Nothing) And rsReturn.Fields.Count > 0 Then
        strPQLString = Trim(rsReturn.Fields(0).Value)
    End If
    

    ' Get SQL statement from DOM
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlPhongQuanLy")
    
    '*************************************
    'Date: 30/05/06
    strSQL = Replace$(strSQL, "MA_PQL", strPQLString)
    '*************************************
'    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
'    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
    
    Set rsReturn = clsDAO.Execute(strSQL)
    If Not (rsReturn Is Nothing) And rsReturn.Fields.Count > 0 Then
        strMaPhongQuanLy = rsReturn.Fields(0).Value
        strTenPhongQuanLy = rsReturn.Fields(1).Value
    End If
    
    
    Set rsReturn = Nothing
   
    Exit Sub
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "GetPQL", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Sub


Private Function GetPhongXuLy(ByVal strPXLString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    

    ' Get SQL statement from DOM
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlPhongXuLy")
    
    '*************************************
    'Date: 30/05/06
    strSQL = Replace$(strSQL, "MA_CQT", strTaxOfficeId)
    '*************************************
'    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
'    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
    
    Set rsReturn = clsDAO.Execute(strSQL)
    
    Set GetPhongXuLy = rsReturn
    
    Set rsReturn = Nothing
    
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetPXL", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

'****************************
'Description: SetupHeaderData procedure get data from DB
'             and fill data to screen and DOM.
'Author:ThanhDX
'Date:23/11/2005
'Input:rsTaxInfor: A recordset contain data which get from DB
'Output:
'Return:
'****************************
'Private Sub SetupHeaderData(ByRef rsTaxInfor As ADODB.Recordset)
'    Dim lIndex As Long, lCtrl As Long
'    Dim lCol As Long, lRow As Long
'
'On Error GoTo ErrHandle
'        fpSpread1.Sheet = lCtrl + 1
'        For lIndex = 1 To TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes.length
'            If lIndex < rsTaxInfor.Fields.Count Then
'                If Not rsTaxInfor.Fields(lIndex) = vbNullString Then
'                    TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex - 1) _
'                        .Attributes.getNamedItem("Value").nodeValue = TAX_Utilities_Srv_New.Convert(rsTaxInfor.Fields(lIndex).Value, TCVN, UNICODE)
'                    ParserCellID fpSpread1, GetAttribute(TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex - 1), "CellID"), lCol, lRow
'                    fpSpread1.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsTaxInfor.Fields(lIndex).Value, TCVN, UNICODE)
'                    fpSpread1.RowHeight(lRow) = fpSpread1.MaxTextRowHeight(lRow)
'                End If
'            Else
'                Exit For
'            End If
'        Next lIndex
'    Exit Sub
'ErrHandle:
'    SaveErrorLog Me.Name, "SetupHeaderData", Err.Number, Err.Description
Private Sub SetupHeaderData(arrStrHeaderData() As String)
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    
On Error GoTo ErrHandle
        fpSpread1.Sheet = lCtrl + 1
        For lIndex = 0 To UBound(arrStrHeaderData) 'TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes.length
            'If lIndex < UBound(arrStrHeaderData) Then
                If Not arrStrHeaderData(lIndex) = vbNullString Then
                    SetAttribute TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex) _
                        , "Value", TAX_Utilities_Srv_New.Convert(arrStrHeaderData(lIndex), TCVN, UNICODE)
                    ParserCellID fpSpread1, GetAttribute(TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex), "CellID"), lCol, lRow
                    fpSpread1.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(arrStrHeaderData(lIndex), TCVN, UNICODE)
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
    Dim i As Long, j As Long, strLoaiDL As String
    
On Error GoTo ErrHandle
    Set xmlListSection = xmlDomData.getElementsByTagName("Section")
    For Each xmlNodeSection In xmlListSection
        If Trim(xmlNodeSection.Attributes.getNamedItem("Dynamic").nodeValue) = "1" Then
            iRowID = 0
            For i = 0 To xmlNodeSection.childNodes.length - 1
                iRowID = iRowID + 1
                For j = 0 To xmlNodeSection.childNodes(i).childNodes.length - 1
                    Set xmlAttribute = xmlDomData.createAttribute("RowID")
                    xmlAttribute.Value = iRowID
                    Set xmlNode = xmlNodeSection.childNodes(i).childNodes(j).Attributes.setNamedItem(xmlAttribute)
                    Set xmlAttribute = Nothing
                Next
            Next
        End If
    Next
        
    strLoaiDL = Trim(TAX_Utilities_Srv_New.NodeValidity.childNodes(lPos).Attributes.getNamedItem("DataFile").nodeValue)
    Set xmlList = xmlDomData.getElementsByTagName("Cell")
    If xmlList.length > 0 Then GenerateSQL_Details = "begin"
    ' Them tham so nay de tinh cu 50 dong se bat dau ghi thanh mot block du lieu
    Dim rC As Integer
    For Each xmlNode In xmlList
        If Not xmlNode.Attributes.getNamedItem("MCT") Is Nothing Then
             If Trim(xmlNode.Attributes.getNamedItem("MCT").nodeValue) <> "" Then
                rC = rC + 1
                strSQL = strSQL_DTL
                strSQL = strSQL & "'" & vHdrID & "',"
                strSQL = strSQL & "'" & strLoaiDL & "',"
                strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("MCT").nodeValue & "',"
                strSQL = strSQL & "'" & Trim(Replace(TAX_Utilities_Srv_New.Convert(xmlNode.Attributes.getNamedItem("Value").nodeValue, UNICODE, TCVN), "'", "''")) & "',"
                If Not xmlNode.Attributes.getNamedItem("RowID") Is Nothing Then
                    strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("RowID").nodeValue & "');"
                Else
                    'strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("MCT").nodeValue & "');"
                    strSQL = strSQL & "'');"
                End If
                GenerateSQL_Details = GenerateSQL_Details & vbCrLf & strSQL
                If rC = 1 Then
                    'clsDAO.BeginTrans
                ElseIf rC = 30 Then
                    rC = 0
                    GenerateSQL_Details = GenerateSQL_Details & vbCrLf & "end;"
                    clsDAO.Execute GenerateSQL_Details
                    'clsDAO.CommitTrans
                    GenerateSQL_Details = vbNullString
                    GenerateSQL_Details = "begin"
                End If
             End If
        End If
    Next
    If Trim(GenerateSQL_Details) <> "begin" Then
        GenerateSQL_Details = GenerateSQL_Details & vbCrLf & "end;"
        clsDAO.Execute GenerateSQL_Details
        'clsDAO.CommitTrans
    End If
    Set xmlDomData = Nothing
    Set xmlList = Nothing
    Set xmlListSection = Nothing
    Exit Function

ErrHandle:
    SaveErrorLog Me.Name, "GenerateSQL_Details", Err.Number, Err.Description
    Err.Raise Err.Number
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
    TAX_Utilities_Srv_New.xmlDataReDim (0)
    cmdViewNow.Enabled = False
    cmdClear.Enabled = False
    cmdSave.Enabled = False
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
    If UCase(Left(strValue, 2)) = "AA" Then
        verToKhai = 0
    ElseIf UCase(Left(strValue, 2)) = "TT" Then
        verToKhai = 1
    ElseIf UCase(Left(strValue, 2)) = "BS" Then
        verToKhai = 2
    Else
        verToKhai = 0
    End If
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
    GetDataFormFile = TAX_Utilities_Srv_New.Convert(GetDataFormFile, TCVN, UNICODE)
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
    
    Set xmlNodeCell = TAX_Utilities_Srv_New.Data(fpSpread1.ActiveSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow))
    
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
Private Function MessageBox(strMsgId As String, intMsgStyle As MsgBoxStyle, intMsgIcon As MsgBoxIcon, Optional msType As Byte) As MsgBoxResult
    Dim intReturn As Integer
    
On Error GoTo ErrHandle
    If blnReceiveByBarcode Then StopBarcodeReader
    
    MessageBox = DisplayMessage(strMsgId, intMsgStyle, intMsgIcon, , msType)
    
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
        xmlDocumentInit(i).Load GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(i), "Ini"))
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

'Private Function GetThongTinTep(ByVal strID As String, arrStrHeaderData() As String) As Boolean
'    Dim lngIndex As Long
'    Dim rsResult As ADODB.Recordset
'    Dim strSQL As String, strMaTkhaiQLT As String
'    Dim strPrefixMaTep As String, strMatep As String
'    Dim strSTT As String
'
'    On Error GoTo ErrHandle
'
'    lngIndex = UBound(arrStrHeaderData)
'
'    On Error GoTo ConnectErrHandle
'    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'    'Lay ma to khai theo QLT
'    strSQL = "Select hso.loai_hoso " & _
'            "From rcv_map_tkhai tkhai," & _
'            "qlt_map_hoso_tkhai hso " & _
'            "Where (tkhai.nhom_hso = hso.nhom) " & _
'            "And (tkhai.ma_tkhai_qlt = hso.loai_tkhai) " & _
'            "And (tkhai.ma_tkhai = '" & strID & "')"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strMaTkhaiQLT = rsResult.Fields(0).Value
'
'    'La^'y chuo^~i tie^`n to^' cu?a ma~ te^.p
'    strSQL = "Select To_Char(Sysdate,'RRMM')||'" & strMaTkhaiQLT & _
'            "' From Dual"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strPrefixMaTep = rsResult.Fields(0).Value
'
'    'Lay so thu tu lon nhat cua tep (hau to)
'    strSQL = "Select nvl(max(To_Number(Substr(So_Hieu_tep,8,3))),1) " & _
'            "From rcv_tkhai_hdr " & _
'            "Where So_Hieu_Tep Like '" & strPrefixMaTep & "' || '%'"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strMatep = strPrefixMaTep & "-" & rsResult.Fields(0).Value
'
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
'    ReDim Preserve arrStrHeaderData(lngIndex + 1)
'    arrStrHeaderData(lngIndex + 1) = strMatep
'
'    'Ghep so thu tu to khai vao chuoi
'    ReDim Preserve arrStrHeaderData(lngIndex + 2)
'    arrStrHeaderData(lngIndex + 2) = "" 'strSTT
'
'    Set rsResult = Nothing
'    GetThongTinTep = True
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetThongTinTep", Err.Number, Err.Description
'    Exit Function
'ConnectErrHandle:
'    SaveErrorLog Me.Name, "GetThongTinTep", Err.Number, Err.Description
'End Function

Private Function getSoTTTK(ByVal strID As String, arrStrHeaderData() As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String
    Dim strMatep As String
    Dim strSTT As Integer
    
    On Error GoTo ErrHandle
    
    lngIndex = UBound(arrStrHeaderData)
    
    On Error GoTo ConnectErrHandle
    'connect to database QLT_TNK
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If

    'Lay so TT to khai trong RCV
    If strID = "02_TNDN11" And isTKLanPS = True Then
        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.loai_tkhai = '" & strID & "' " & _
                " And tkhai.ngay_ps = to_date('" & ngayPS & "','dd/mm/yyyy')"
    ElseIf (strID = "01_NTNN" Or strID = "01_TTDB11" Or strID = "03_NTNN11") And isTKLanPS = True Then
        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.loai_tkhai = '" & strID & "' " & _
                " And tkhai.ngay_ps = to_date('" & ngayPS & "','dd/mm/yyyy')"
    ElseIf (strID = "08_TNCN11" Or strID = "08A_TNCN11") And isTKThang = True Then
        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.loai_tkhai = '" & strID & "' " & _
                "And tkhai.kykk_tu_ngay = To_Date('" & "01/" & TuNgay & "','DD/MM/RRRR')" & _
                "And tkhai.kykk_den_ngay = To_Date('" & "01/" & DenNgay & "','DD/MM/RRRR')"
    Else
        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.loai_tkhai = '" & strID & "' " & _
                "And tkhai.kykk_tu_ngay = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
                "And tkhai.kykk_den_ngay = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
    End If
    
    Set rsResult = clsDAO.Execute(strSQL)
    If rsResult Is Nothing Or IsNull(rsResult.Fields(0)) Then
        strSTT = 0
        isTKTonTai = False
        ' Doi voi cac to khai 01_NTNN, 03_NTNN, 01_TTDB, 02_TNDN
        If (strID = "01_NTNN" Or strID = "01_TTDB11" Or strID = "03_NTNN11" Or strID = "02_TNDN11") And isTKLanPS = True Then
            isToKhaiPsDaNhanTN = False
        End If
        
    Else
        strSTT = rsResult.Fields(0).Value + 1
        isTKTonTai = True
        ' Doi voi cac to khai 01_NTNN, 03_NTNN, 01_TTDB, 02_TNDN trong 1 ngay chi nhan 1 to khai
        If (strID = "01_NTNN" Or strID = "01_TTDB11" Or strID = "03_NTNN11" Or strID = "02_TNDN11") And isTKLanPS = True Then
            isToKhaiPsDaNhanTN = True
        End If
    End If
    
    ' Kiem tra to khai chinh thuc
    If strSTT = 0 Then
        isToKhaiCT = False
    Else
        isToKhaiCT = True
    End If
    
    
    'Ghep ma so tep vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 1)
    arrStrHeaderData(lngIndex + 1) = strSTT
    
    'Ghep so thu tu to khai vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 2)
    arrStrHeaderData(lngIndex + 2) = strSTT
    
    Set rsResult = Nothing
    getSoTTTK = True
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK", Err.Number, Err.Description
End Function


' Kiem tra to khai theo DA30
Private Function isDA30(ByVal strID As String, arrStrHeaderData() As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    On Error GoTo ConnectErrHandle
    'connect to database QLT_TNK
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If

    strSQL = "select 1 from qlt_tkhai_hdr tkhai " & _
            "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
            "And tkhai.DTK_MA_LOAI_TKHAI = '" & changeMaToKhaiQLT(strID) & "' " & _
            "And tkhai.kykk_tu_ngay = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
            "And tkhai.kykk_den_ngay = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
            "And tkhai.YN_DA30 is null "

    Set rsResult = clsDAO.Execute(strSQL)
    If rsResult Is Nothing Then
        isDA30 = False
    Else
        isDA30 = True
    End If
       
    Set rsResult = Nothing
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "isDA30", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "isDA30", Err.Number, Err.Description




    isDA30 = False
End Function
' end

Private Function GetHeaderData(ByVal rsTaxInfor As ADODB.Recordset, arrStrHeaderData() As String) As Boolean
    Dim arrStrData() As String
    Dim lCtrl As Long
    Dim clsConvert  As New clsUnicodeConvert
    On Error GoTo ErrHandle
    
    If rsTaxInfor Is Nothing Then
        Exit Function
    End If
    
    If rsTaxInfor.Fields.Count = 0 Then
        Exit Function
    End If
    
    For lCtrl = 0 To rsTaxInfor.Fields.Count - 2
        ReDim Preserve arrStrData(lCtrl)
        If Not IsNull(rsTaxInfor.Fields(lCtrl + 1).Value) Then
            'arrStrData(lCtrl) = clsConvert.Convert(rsTaxInfor.Fields(lCtrl + 1).Value, UNICODE, TCVN)
            arrStrData(lCtrl) = rsTaxInfor.Fields(lCtrl + 1).Value
            
        End If
    Next lCtrl
           
    'Loai bo gia tri Ngay bat dau nam TC va Ngay bat dau KD
    ReDim Preserve arrStrData(UBound(arrStrData) - 1)
    
    arrStrHeaderData = arrStrData
    GetHeaderData = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetHeaderData", Err.Number, Err.Description
End Function

' Ham lay ve so tt quet An chi
Private Function getSoTTTK_AC(ByVal strID As String, arrStrHeaderData() As String, strData As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String
    Dim strMatep As String
    Dim strSTT As Integer
    
    Dim arrDeltail() As String
    
    On Error GoTo ErrHandle
    
    lngIndex = UBound(arrStrHeaderData)
    
    On Error GoTo ConnectErrHandle
    'connect to database QLT_TNK
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    ' Tach ma so thue 13 thanh ma so thue 14
    If Len(Trim(arrStrHeaderData(0))) = 13 Then
        arrStrHeaderData(0) = Left(Trim(arrStrHeaderData(0)), 10) & "-" & Right(Trim(arrStrHeaderData(0)), 3)
    End If
    
    'Lay so TT to khai trong RCV
    If strID = "01_TBAC" Then
        arrDeltail = Split(strData, "~")
        If Len(Trim(arrDeltail(UBound(arrDeltail) - 3))) = 13 Then
            arrDeltail(UBound(arrDeltail) - 3) = Left(arrDeltail(UBound(arrDeltail) - 3), 10) & "-" & Right(arrDeltail(UBound(arrDeltail) - 3), 3)
        End If
        
        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
        "And tkhai.LOAI_BC = '" & strID & "' " & _
        " And tkhai.NGAY_BC=to_date('" & arrDeltail(UBound(arrDeltail) - 1) & "','dd/mm/rrrr')" & _
        " And tkhai.TIN_DV_CQ='" & Trim(arrDeltail(UBound(arrDeltail) - 3)) & "'"
    ElseIf strID = "03_TBAC" Then
        arrDeltail = Split(strData, "~")
        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
        "And tkhai.LOAI_BC = '" & strID & "' " & _
        " And tkhai.NGAY_BC=to_date('" & Left$(arrDeltail(UBound(arrDeltail)), 10) & "','dd/mm/rrrr')"
    ElseIf strID = "BC21_AC" Then
        arrDeltail = Split(strData, "~")
        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
        "And tkhai.LOAI_BC = '" & strID & "' " & _
        " And tkhai.NGAY_BC=to_date('" & Left$(arrDeltail(UBound(arrDeltail)), 10) & "','dd/mm/rrrr')"
    ElseIf strID = "01_AC" Then
        arrDeltail = Split(strData, "~")
        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
        "And tkhai.LOAI_BC = '" & strID & "' " & _
        "And tkhai.KYBC_TU_NGAY = to_date('" & arrDeltail(1) & "','dd/mm/rrrr')" & _
        "And tkhai.KYBC_DEN_NGAY = to_date('" & Left$(arrDeltail(2), 10) & "','dd/mm/rrrr')"
    Else
        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.LOAI_BC = '" & strID & "' " & _
                "And tkhai.KYBC_TU_NGAY = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
                "And tkhai.KYBC_DEN_NGAY = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
    End If
    
    Set rsResult = clsDAO.Execute(strSQL)
    If rsResult Is Nothing Or IsNull(rsResult.Fields(0)) Then
        strSTT = 0
        isTonTaiAC = False
    Else
        strSTT = rsResult.Fields(0).Value + 1
        isTonTaiAC = True
    End If
    
    'Ghep ma so tep vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 1)
    arrStrHeaderData(lngIndex + 1) = strSTT
    
    'Ghep so thu tu to khai vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 2)
    arrStrHeaderData(lngIndex + 2) = strSTT
    
    Set rsResult = Nothing
    getSoTTTK_AC = True
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK_AC", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK_AC", Err.Number, Err.Description
End Function


' Kiem tra xem co to khai chinh thuc chua
'Private Function isToKhaiCT(ByVal strID As String, arrStrHeaderData() As String) As Boolean
'    Dim lngIndex As Long
'    Dim rsResult As ADODB.Recordset
'    Dim strSQL As String
'    Dim strMatep As String
'    Dim strSTT As Integer
'
'    On Error GoTo ErrHandle
'
'    lngIndex = UBound(arrStrHeaderData)
'
'    On Error GoTo ConnectErrHandle
'    'connect to database QLT_TNK
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'    'Lay so TT to khai trong RCV
'
'    strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
'            "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'            "And tkhai.loai_tkhai = '" & strID & "' " & _
'            "And tkhai.kkbs='1' " & _
'            "And tkhai.kykk_tu_ngay = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
'            "And tkhai.kykk_den_ngay = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    If rsResult Is Nothing Or IsNull(rsResult.Fields(0)) Then
'        isToKhaiCT = False
'    Else
'        isToKhaiCT = True
'    End If
'
'    Set rsResult = Nothing
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "isToKhaiCT", Err.Number, Err.Description
'    Exit Function
'ConnectErrHandle:
'    SaveErrorLog Me.Name, "isToKhaiCT", Err.Number, Err.Description
'End Function

' Kiem tra mst dai ly co phai cua NNT hay khong?
' Lay thong tin DL thue 05072011
Private Function isMaDLT(ByVal strTaxIDString As String, ByVal strTaxIDDLString As String) As Boolean
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    

    ' Get SQL statement from DOM
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlMSTDL")
    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
    strSQL = Replace(strSQL, "ma_dai_ly", "'" & strTaxIDDLString & "'")
    
    Set rsReturn = clsDAO.Execute(strSQL)
    
    If rsReturn Is Nothing Or rsReturn.Fields.Count = 0 Then
        If Trim(strTaxIDDLString) = "" Or strTaxIDDLString = vbNullString Then
            isMaDLT = True
        Else
            isMaDLT = False
        End If
    Else
        isMaDLT = True
    End If
    
    Set rsReturn = Nothing
    
    Exit Function
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "isMaDLT", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function
