VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
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
         Caption         =   "X�a"
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
         Caption         =   "Xem t� khai"
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
         Caption         =   "Ghi l�i"
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
         Caption         =   "Tho�t"
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

    Dim strSQL         As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String
    Dim HdrID          As Variant, strDate() As String, dDate As Date
    Dim rs             As New ADODB.Recordset, i As Long
    Dim qBoSung        As Variant
    Dim msgRs          As MsgBoxResult
    Dim idToKhai       As Integer
    
    Dim mTemp          As Integer
    
    Dim dsTK_DLT       As String
    'dntai them bien de luu ngay dau nam tai chinh va ngay cuoi nam tai chinh
    Dim dNgayDauNamTC  As Date
    Dim dNgayCuoiNamTC As Date
    Dim varDate1       As String
    Dim varDate2       As String

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
    
    dsTK_DLT = "~1~2~3~4~5~6~11~12~46~47~48~49~15~16~50~51~36~70~71~72~73~74~75~80~81~82~77~86~87~89~42~43~17~59~41~76~90~"
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
    If (Val(idToKhai) <= 68 And Val(idToKhai) >= 64) Or Val(idToKhai) = 91 Then

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
            If (Val(idToKhai) = 1 Or Val(idToKhai) = 2 Or Val(idToKhai) = 4 Or Val(idToKhai) = 71 Or Val(idToKhai) = 36) And LoaiKyKK = True Then
                strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
                dDate = DateAdd("m", 3, dDate)
                dDate = DateAdd("d", -1, dDate)
                strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy')"

            Else
            
                strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
                dDate = DateAdd("m", 1, dDate)
                dDate = DateAdd("d", -1, dDate)
                strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
            End If

        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then

            If Val(idToKhai) = 68 And LoaiKyKK = False Then
                strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
                dDate = DateAdd("m", 1, dDate)
                dDate = DateAdd("d", -1, dDate)
                strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "

            Else
                strSQL = strSQL & " and KYKK_TU_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy') "
                dDate = DateAdd("m", 3, dDate)
                dDate = DateAdd("d", -1, dDate)
                strSQL = strSQL & " and KYKK_DEN_NGAY=To_date('" & format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy')"

            End If

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
            If (Val(idToKhai) = 24 Or Val(idToKhai) = 25 Or Val(idToKhai) = 26 Or Val(idToKhai) = 27 Or Val(idToKhai) = 28 Or Val(idToKhai) = 29 Or Val(idToKhai) = 30 Or Val(idToKhai) = 31 Or Val(idToKhai) = 32 Or Val(idToKhai) = 33 Or Val(idToKhai) = 34 Or Val(idToKhai) = 35 Or Val(idToKhai) = 55 Or Val(idToKhai) = 56 Or Val(idToKhai) = 57 Or Val(idToKhai) = 58) Then
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

    
    If idToKhai = 2 Or idToKhai = 4 Or idToKhai = 46 Or idToKhai = 47 Or idToKhai = 48 Or idToKhai = 49 Or idToKhai = 15 Or idToKhai = 16 Or idToKhai = 50 Or idToKhai = 51 Or idToKhai = 36 Or idToKhai = 6 Or idToKhai = 72 Or idToKhai = 87 Or idToKhai = 86 Or idToKhai = 77 Or idToKhai = 71 Or idToKhai = 74 Or idToKhai = 89 Or idToKhai = 42 Or idToKhai = 43 Or idToKhai = 17 Or idToKhai = 59 Or idToKhai = 41 Or idToKhai = 76 Or idToKhai = 90 Or idToKhai = 95 Or idToKhai = 92 Or idToKhai = 93 Or idToKhai = 94 Then
        strSQL_HDR = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlHdrTT28").nodeValue)
    ElseIf idToKhai = 1 Or idToKhai = 11 Or idToKhai = 12 Or idToKhai = 5 Or idToKhai = 70 Or idToKhai = 80 Or idToKhai = 81 Or idToKhai = 82 Or idToKhai = 3 Or idToKhai = 73 Then
        strSQL_HDR = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlHdrTT28_NNKD").nodeValue)
    Else
        strSQL_HDR = CStr(xmlSQL.getElementsByTagName("SQLs")(0).Attributes.getNamedItem("SqlHdr").nodeValue)
    End If

    ' xu ly de ghi cac mau an chi
    If Val(idToKhai) = 66 Or Val(idToKhai) = 68 Or Val(idToKhai) = 67 Or Val(idToKhai) = 64 Or Val(idToKhai) = 65 Or Val(idToKhai) = 91 Then
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
    Dim str1  As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String, str9 As String, str10 As String
    Dim str11 As String, str12 As String, str13 As String, str14 As String, str15 As String, str16 As String, str17 As String, str18 As String, str19 As String, str20 As String
    Dim str21 As String, str22 As String, str23 As String, str24 As String, str25 As String, str26 As String, str27 As String, str28 As String, str29 As String, str30 As String
    Dim str31 As String, str32 As String, str33 As String, str34 As String, str35 As String, str36 As String, str37 As String, str38 As String, str39 As String, str40 As String
    Dim str41 As String, str42 As String, str43 As String, str44 As String, str45 As String, str46 As String, str47 As String, str48 As String, str49 As String, str50 As String
    Dim str51 As String, str52 As String, str53 As String

    'str2 = "aa315703600247325   12201200100100100101/0101/01/1900<S01><S></S><S>dfsdf~2222222222~~0~~100000000~10~10~1000000~10000000~5~0~500000~1500000</S><S>100000000~1000000~10000000~0~500000~1500000</S><S>~X</S><S>~~Nguy�n Phi Phong~18/12/2012~1~~~18/12/2012</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa315023600247325   12201200100100100201/0114/06/2006<S01><S></S><S>0~10000000~1000000~0~0~0~0~1000000~0~1000000~0~0~1000000</S><S>~Nguy�n Phi Phong~~08/01/2013~1~~</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa315023600247325   122012001001002002<S01_2_9><S>dfsdf~42423~31/12/2012~nfgdgfd~2222222222~thuoc l�~10000000~10~1000000~</S><S>10000000~1000000</S></S01_2_9>"
    'str2 = "aa315033600247325   00201200100100100201/0114/06/200601/01/201231/12/2012<S03><S></S><S>0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~"
    'Barcode_Scaned str2
    'str2 = "aa315033600247325   0020120010010020020~0~0~0~0~-100000~0~0</S><S></S><S></S><S></S><S>~Nguy�n Phi Phong~~08/01/2013~1~1~0~1052</S></S03>"
    'Barcode_Scaned str2
    'str2 = "aa315013600247325   04201200100100100101/0114/06/2006<S01><S></S><S>~2000000~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~0~2000000~0~2000000</S><S>~~Nguy�n Phi Phong~08/01/2013~1~~~1701~x~01</S></S01>"
    'Barcode_Scaned str2
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
    'str2 = "aa316013600247325   05201300300300100201/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~234~0~0~0~0~0~0~0~468~0~0~0</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   05201300400400100301/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~-23"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   0520130040040020034~0~0~~0~~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004004003003<S01_5><S>11/06/2013~23423423~10101~11/06/2013~23423423~10101~11/06/2013~23423423~10101</S></S01_5>"
    'Barcode_Scaned str2

    '' phu luc 05
    'str2 = "aa316013600247325   05201300400400100301/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~-23"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   0520130040040020034~0~0~~0~~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004004003003<S01_5><S>11/06/2013~23423423~10101~11/06/2013~23423423~10101~11/06/2013~23423423~10101</S></S01_5>"
    'Barcode_Scaned str2

    ''phu luc 06
    'str2 = "aa316013600247325   05201300400700100401/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~-234"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004007002004~0~0~0~0~0~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004007003004<S01_6><S>test~3600247325~54654~546546~60120"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   0520130040070040040~11478~0~10301~10300</S><S>~~~</S></S01_6>"
    'Barcode_Scaned str2

    ''phu luc 6B
    'str2 = "aa316013600247325   05201300400900100301/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~-234"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004009002003~0~0~0~0~0~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004009003003<S01_6B><S>test~3600247325~234~5~0~10701</S><S>~~~</S></S01_6B>"
    'Barcode_Scaned str2

    ''phu luc 07
    'str2 = "aa316013600247325   05201300401000100401/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~234"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004010002004~0~0~0~0~0~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004010003004<S01_7><S>ewru~3600247325~24342340~234234324~2585"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   05201300401000400476664~4928109~0</S><S>0~23423423~0~0</S></S01_7>"
    'Barcode_Scaned str2

    '' phu luc 01TD
    'str2 = "aa316013600247325   05201300501100100301/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~-234"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013005011002003~0~0~0~0~0~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013005011003003<S01_1_TD><S>erwrew~23432432~234324320~2342343</S></S01_1_TD>"
    'Barcode_Scaned str2

    '' phu luc 02TD
    'str2 = "aa316013600247325   05201300401500100301/0114/06/2006<S01><S></S><S>~234~0~23423~234~0~0~0~0~0~0~0~0~0~0~-234"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004015002003~0~0~0~0~0~0~0~468~0~0~468</S><S>~~~20/06/2013~1~~~1701~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316013600247325   052013004015003003<S01_2_TD><S>weewrew~55~3535353~10304~~fdgwer~43~43344343~11105~</S></S01_2_TD>"
    'Barcode_Scaned str2
    '
    '' To khai 02 GTGT
    'str2 = "aa316023600247325   06201300100100100101/0114/06/2006<S01><S></S><S>120~213123~32432~324320~432432~432~32423~432423~-107671~234~105338~34~324234~43~105261</S><S>ewrew~rwerwe~23423423~07/07/2013~1~~~1</S></S01>"
    'Barcode_Scaned str2
    '
    '' To khai 03 GTGT
    'str2 = "aa316043600247325   06201300100100100101/0114/06/2006<S01><S></S><S>3242~0~23432~23423~0~432423~-20181~0~-408991~0~0~0~26674~0</S><S>3424~234wer~234234~07/07/2013~1~~~0</S></S01>"
    'Barcode_Scaned str2

    '' To khai 03B GTGT
    'str2 = "aa316953600247325   06201300100100100101/0114/06/2006<S01><S></S><S>1~2~3~4~5~6~7~8~9~10~11~12~13~14~15~16~17~18~19~29~21~22~23</S><S>H? v� t�n~K� t�n~Ch?ng ch? h�nh ngh? s?~Ng�y k�~~1~2~1</S></S01>"
    'Barcode_Scaned str2

    '' Phu luc 03B_1_GTGT
    'str2 = "aa316953600247325   06201300300300100201/0114/06/2006<S01><S></S><S>1~2~3~6~13~20~6~5~4~7~8~9~12~11~10~-6~3~12~0~0~1~39~1</S><S>aaa~aaaa~aq334~09/07/2013~1~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316953600247325   062013003003002002<S01_1><S>11~22~01/01/2013~33~2342343243~54~65~76</S><S>10~20~01/01/2013~30~2342343243~50~60~07</S><S>8~9~01/01/2013~10~2342343243~12~13~14</S><S>15~16~01/01/2013~17~2342343243~19~20~21</S><S>45~39</S></S01_1>"
    'Barcode_Scaned str2
    '
    '' Phu luc 03B_2_GTGT
    'str2 = "aa316953600247325   06201300400400100301/0114/06/2006<S01><S></S><S>1~2~3~0~0~0~6~5~4~7~8~9~12~11~10~-12~-10~-8~0~0~0~0~0</S><S>aaa~aaaa~aq334~09/07/2013~1~~~0</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa316953600247325   062013004004002003<S01_2><S>11~22~33~44~55~66~0~77</S><S>111~222~333~444~555~666~0~777</S><S>123~321~111~222"
    'Barcode_Scaned str2
    'str2 = "aa316953600247325   062013004004003003~333~444~0~555</S><S>212~211~3232~431~43~3~0~4343</S><S>0~0</S></S01_2>"
    'Barcode_Scaned str2

    ''To khai 04/GTGT
    'str2 = "aa316713600247325   06201300100200100301/0101/01/1900<S01><S></S><S>0~0~0</S><S>0~0~0~</S><S>0~0~0~</S><S>0~0~0</S><S>0~0~0~</S><S>0~0~0</S><S>0~0~0~</S><S>0~0</S><S>~~~09/07/2013~1~~~1</S></S01>"
    'Barcode_Scaned str2
    '
    ''To khai 01A/TNDN
    'str2 = "aa317113600247325   02201300100300100101/0114/06/2006<S01><S></S><S>1~2~-1~3~4~-2~5~0~0~0~0~0~x</S><S>x~x</S><S>a~c~b~25/07/2013~1~0~~1052~02</S></S01>"
    'Barcode_Scaned str2

    '' Mau TB04/AC
    'str2 = "aa316913600247325   06201300300300100101/0101/01/2009<S01><S>1~11~01DVPH~2~22~02DCTS~3~33~03Phone</S><S>01/01/2009~0102030405001~tong cuc thue~01/01/2009~nguyen van a</S></S01>"
    'Barcode_Scaned str2
    ''str2 = "aa316913600247325   0620130030030020021/2009~0102030405001~tong cuc thue~01/01/2009~nguyen van a</S></S01>"
    ''Barcode_Scaned str2

    '' phu luc 05A/TNDN
    'str2 = "aa317113600247325   02201300100400100201/0114/06/2006<S01><S></S><S>0~0~0~0~0~0~0~0~0~0~0~0~x</S><S>~</S><S>~~~25/07/2013~1~0~~1052~02</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa317113600247325   022013001004002002<S01-1><S>2134325</S><S>test~0102030405~100~0~10101</S></S01-1>"
    'Barcode_Scaned str2

    'To khai 01B/TNDN
    'str2 = "aa317123600247325   02201300100100100101/0114/06/2006<S01><S></S><S>x~x~0~0~0~0~20~~0~0~0~0~0~0~x</S><S>arae~26/07/2013~tesdt~rewr~1~~1052~02</S></S01>"
    'Barcode_Scaned str2

    '' phu luc 05A/TNDN
    'str2 = "aa317123600247325   02201300200200100201/0114/06/2006<S01><S></S><S>x~x~0~0~0~0~20~~0~0~0~0~0~0~x</S><S>arae~26/07/2013~tesdt~rewr~1~~1052~02</S></S01>"
    'Barcode_Scaned str2
    'str2 = "aa317123600247325   022013002002002002<S01-1><S>0</S><S>wetrwe~2222222222~90~0~10301~qweqw~0102030405~10~0~10304</S></S01-1>"
    'Barcode_Scaned str2

    ''02/TNDN
    'str2 = "aa317733600247325   02201300100100100201/0114/06/2006<S02><S></S><S>124~3897~324~2340~342~423~234~234~-3773~0~-3773~23~0~54~0~4365~5466~23~1~0"
    'Barcode_Scaned str2
    'str2 = "aa317733600247325   022013001001002002</S><S>1~~test~2222222222~tes~44~01/08/2013~02/08/2013</S><S>ttest~34~tet~28/07/2013~1~~~1053~x</S></S02>"
    'Barcode_Scaned str2

    ''03A TNCN
    'str2 = "aa317503600247325   06201300100100100101/0101/01/2010<S01><S></S><S>12321321~616066~423432423~423432~43543543~2177177~43543543~4354354~54645654~0~54667576~0</S><S>test~29/07/2013~test~3242~1~~</S></S01>"
    'Barcode_Scaned str2

    ''08 KK TNCN
    'str2 = "aa317742600168262   02201300100100100101/0101/01/1900<S01><S></S><S>1~2~0~11~0~0~5~6~0~0~0~0~0</S><S>test~29/07/2013~test~213~1~</S></S01>"
    'Barcode_Scaned str2

    '07 KK TNCN
    'str2 = "aa317362600102014   02201300000100100101/0101/01/2010<S07><S></S><S>~0~4000000~4000000~0~0~0~0~0~0~0~0~0~20~0~0~1~0</S><S>~16/07/2013~~~1~~~1</S></S07>"
    'Barcode_Scaned str2

    '08A KK TNCN
    'str2 = "aa317753600247325   02201300000000100201/0101/01/1900<S01><S>0102030405002</S><S>0~0~0</S><S>~~0~0~0"
    'Barcode_Scaned str2
    'str2 = "aa317753600247325   022013000000002002~0~0~0~0~0</S><S>~30/07/2013~Nguyen Van AA~ict02~1~</S></S01>"
    'Barcode_Scaned str2

    '03_TNDN13
    '    str2 = "aa320033600247325   00201200500800100201/0114/06/200601/01/201231/12/2012<S03><S></S><S>100000000~5000000~1000000~1000000~1000000~1000000~1000000~4000000~1000000~1000000~1000000~1000000~101000000~100000000~1000000~100000000~1000000~1000000~9800000"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   0020120050080020020~1000000~97000000~1000000~970000000000~1000000~1000000~1000000~969997000000~1000000~969998000000~969997000000~1000000</S><S>x~</S><S></S><S>Tr��ng h�p ���c gia h�n:</S><S>~~~01/01/2007~1~1~0~1052~</S></S03>"
    '    Barcode_Scaned str2

    'PL 03_3A_TNDN
    
    '    str2 = "aa320033600247325   00201200601000100401/0114/06/200601/01/201231/12/2012<S03><S></S><S>100000000~5000000~1000000~1000000~1000000~1000000~1000000~4000000~1000000~1000000~1000000~1000000~101000000~100000000~1000000~100000000~1000000~1000000~9800"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   0020120060100020040000~1000000~97000000~1000000~1500000000~100000000~1000000~1000000~1398000000~1000000~1399000000~1398000000~1000000</S><S>x~</S><S></S><S>Tr��ng h�p ���c gia h�n:</S><S>~~~01/01/2007~1~1~0~1052~</S></S03>"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   002012006010003004<S03-3A><S>x~~~~~x~~~~~x~10.000~2002~2020~18~2002~2~2002~10000000000~10000000"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   00201200601000400400~2500000000~1500000000~1000000000~100~1000000000~10~100000000</S></S03-3A>"
    '    Barcode_Scaned str2
    
    '    'PL 03_3B_TNDN
    '    str2 = "aa320033600247325   00201200701200100601/0114/06/200601/01/201231/12/2012<S03><S></S><S>100000000~5000000~1000000~1000000~1000000~1000000~1000000~4000000~1000000~1000000~1000000~1000000~101000000~100000000~1000000~100000000~1000000~1000000~9800"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   0020120070120020060000~1000000~97000000~1000000~1501500000~100100000~1000000~1000000~1399400000~1000000~1400400000~1399400000~1000000</S><S>x~</S><S></S><S>Tr��ng h�p ���c gia h�n:</S><S>~~~01/01/2007~1~1~0~1052~</S></S03>"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   002012007012003006<S03-3A><S>x~~~~~x~~~~~x~10.000~2002~2020~18~2002~2~2002~10000000000~10000000"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   00201200701200400600~2500000000~1500000000~1000000000~100~1000000000~10~100000000</S></S03-3A>"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   002012007012005006<S03-3B><S>x~x~~~~x~x~dau tu 1~dau tu 2~dau tu 3~01/01/2002~10.000~2002~2020~18~2002~2~2002~100"
    '    Barcode_Scaned str2
    '    str2 = "aa320033600247325   0020120070120060060000000~100000000~100000000~10000000~2500000~1000000~1500000~10~1000000~10~100000</S></S03-3B>"
    '    Barcode_Scaned str2
    
    'PL 03_5_TNDN
'    str2 = "aa320033600247325   00201200801400100701/0114/06/200601/01/201231/12/2012<S03><S></S><S>100000000~5000000~1000000~1000000~1000000~1000000~1000000~4000000~1000000~1000000~1000000~1000000~101000000~100000000~1000000~100000000~1000000~1"
'    Barcode_Scaned str2
'    str2 = "aa320033600247325   002012008014002007000000~98000000~100000~97900000~100000000~1500001500~100000100~1000000~1000000~~1000000~1000000~~1000000</S><S>x~</S><S></S><S>Tr��ng h�p ���c gia h�n:</S><S>~~~01/01/2007~1~1~0~1052~</S></S03>"
'    Barcode_Scaned str2
'    str2 = "aa320033600247325   002012008014003007<S03-3A><S>x~~~~~x~~~~~x~10.000~2002~2020~18~2002~2~2002~10000000000~10000000"
'    Barcode_Scaned str2
'    str2 = "aa320033600247325   00201200801400400700~2500000000~1500000000~1000000000~100~1000000000~10~100000000</S></S03-3A>"
'    Barcode_Scaned str2
'    str2 = "aa320033600247325   002012008014005007<S03-3B><S>x~x~~~~x~x~dau tu 1~dau tu 2~dau tu 3~01/01/2002~10.000~2002~2020~18~2002~2~2"
'    Barcode_Scaned str2
'    str2 = "aa320033600247325   002012008014006007002~1000000000000~100000000~100000000~10000~2500~1000~1500~10~1000~10~100</S></S03-3B>"
'    Barcode_Scaned str2
'    str2 = "aa320033600247325   002012008014007007<S03-5><S>100000000000~10005000000~1000000~10000000000~1000000~1000000~1000000~1000000~89995000000~1000000~89994000000~1000000~89993000000~23~20698390000~1000000~0~20697390000</S></S03-5>"
'    Barcode_Scaned str2
    
    'To khai 01/NTNN
    'str2 = "aa320703600247325   07201302102500100101/0101/01/1900<S01><S></S><S>abc~0102030405~0001~100000000~21/07/2013~100000000~10~10~1000000~100000000~10~10~9999990~10999990</S><S>0~0~0~0~0~0</S><S>X~</S><S>~~~21/08/2013~1~~~</S></S01>"

'    str2 = "bs320703600247325   07201302002300100201/0101/01/1900<S01><S></S><S>abc~0102030405~0001~100000000~21/07/2013~100000000~10~10~1000000~1000000000~10~1000000~99000000~100000000~bca~0102030405~0001~100000000~21/07/2013~100000000~10~10~1000000~1000000000~10~2000000~98000000~99000000</S><S>0~0~1000000000~10000000~10000000~10000000</S><S>X~</S><S>~~~21/08/2013~~1~1~</S></S01>"
'    Barcode_Scaned str2
'    str2 = "bs320703600247325   072013020023002002<SKHBS><S>Thu� TNDN ph�i n�p~13~0~10000000~10000000</S><S>~~0~0~0</S><S>21/08/2013~1~5000~ok~10000000</S></SKHBS>"
'    Barcode_Scaned str2
    
    'To khai 02/NTNN
'    str2 = "bs320803600247325   00201200200200100201/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD01~06/06/2013</S><S>100~20000000~100000000~80000000~ok~80~10000000~20000000~10000000~ok~60~1000000~1000000~0~~~0~0~0~~2~0~-2~1~1~0~"
'    Barcode_Scaned str2
'    str2 = "bs320803600247325   002012002002002002-1~ok~1~1~0~-1~ok~2~2~0~1~1~1~0~ok~1~1~1~0~ok~0~-2~-2~~0~-1~-1~ok~~0~-1~-1~ok</S><S>nguyen van a~27/08/2013~ict01~giam doc 1~1~~~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
    'bo sung
'    str2 = "bs320803600247325   00201201401400100301/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD1~02/02/2013</S><S>100~1000000~100000000~99000000~~80~0~0~0~~60~0~0~0~~~0~0~0~~1000000~20000000~19000000~~1000000~2"
'    Barcode_Scaned str2
'    str2 = "bs320803600247325   0020120140140020030000000~19000000~~~0~0~0~~0~0~0~~0~0~0~~~0~0~0~~1000000~20000000~19000000~~1000000~20000000~19000000~~~0~0~0~</S><S>~28/08/2013~~~~1~1~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
'    str2 = "bs320803600247325   002012014014003003<SKHBS><S>Thu� GTGT ph�i n�p c�a nh� th�u n��c ngo�i~4a~0~20000000~20000000</S><S>~~0~0~0</S><S>28/08/2013~195~1950000~~20000000</S></SKHBS>"
'    Barcode_Scaned str2
    
    'To khai 02/NTNN - PL 01
'    str2 = "aa320803600247325   00201200300300100301/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD01~06/06/2013</S><S>100~1000000000~1000000000~0~ok~80~10000000~20000000~10000000~ok~60~1000000~1000000~0~~~0~0~0~~2~0~-2~1~1~0~-1"
'    Barcode_Scaned str2
'    str2 = "aa320803600247325   002012003003002003~ok~1~1~0~-1~ok~2~2~0~1~1~1~0~ok~1~1~1~0~ok~0~-2~-2~~0~-1~-1~ok~~0~-1~-1~ok</S><S>nguyen van a~27/08/2013~ict01~giam doc 1~1~~~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
'    str2 = "aa320803600247325   002012003003003003<S01-1><S>ibm~my~0102030405~~01~abc~ha noi~01/01/2013~1000000000~1000000000~1000000000~1000000000~100</S><S>1000000000~1000000000~100</S></S01-1>"
'    Barcode_Scaned str2
        
    'To khai 02/NTNN - PL 02
'    str2 = "aa320803600247325   00201200400400100401/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD01~06/06/2013</S><S>100~1000000000~1000000000~0~ok~80~1000000~1000000~0~ok~60~1000000~1000000~0~~~0~0~0~~2~0~-2~1~1~0~-1~ok~1"
'    Barcode_Scaned str2
'    str2 = "aa320803600247325   002012004004002004~1~0~-1~ok~2~2~0~1~1~1~0~ok~1~1~1~0~ok~0~-2~-2~~0~-1~-1~ok~~0~-1~-1~ok</S><S>nguyen van a~27/08/2013~ict01~giam doc 1~1~~~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
'    str2 = "aa320803600247325   002012004004003004<S01-1><S>ibm~my~0102030405~~01~abc~ha noi~01/01/2013~1000000000~1000000000~1000000000~1000000000~100</S><S>1000000000~1000000000~100</S></S01-1>"
'    Barcode_Scaned str2
'    str2 = "aa320803600247325   002012004004004004<S01-2><S>fpt~0102030405~~HD01~test~ha noi~01/01/2013~1000000~1000000~1000000~1000000</S><S>1000000~1000000</S></S01-2>"
'    Barcode_Scaned str2
    
    'To khai 03/NTNN
    'str2 = "aa320813600247325   07201300300300100101/0101/01/1900<S01><S></S><S>01~~~0~~1000000000~1~2000000~8000000</S><S>0~1000000000~2000000~8000000</S><S>1~</S><S>~~~28/08/2013~1~1~~</S></S01>"
    'Barcode_Scaned str2
    
    'bo sung
'    str2 = "bs320813600247325   07201300500500100201/0101/01/1900<S01><S></S><S>01~~~0~~1000000000~1~3000000~7000000</S><S>0~1000000000~3000000~7000000</S><S>1~</S><S>~~~28/08/2013~~1~1~</S></S01>"
'    Barcode_Scaned str2
'    str2 = "bs320813600247325   072013005005002002<SKHBS><S>~~0~0~0</S><S>Thu� TNDN ph�i n�p~9~8000000~7000000~-1000000</S><S>28/08/2013~8~-4000~~-1000000</S></SKHBS>"
'    Barcode_Scaned str2

    'to khai 04/NTNN
'    str2 = "aa320823600247325   00201200100100100101/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD01~06/06/2013</S><S>10~10000000~200000000~190000000~ok~~5000000~70000000~0~~~1000000~20000000~19000000~ok~~9000000~180000000~171000000~~1000000~10000000~0~-10000000~ok~2000000~20000000~0~-20000000~ok~10~-10000000~0~10000000~</S><S>dai ly~nguoiky~ict01~29/08/2013~1~~~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
    
    'Phu luc 01
'    str2 = "aa320823600247325   00201200500600100201/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD01~06/06/2013</S><S>10~10000000~200000000~190000000~ok~~0~2000000~0~~~0~2000000~2000000~ok~~10000000~198000000~188000000~~1000000~10000000~0~-10000000~ok~2000000~20000000~0~-20000000~ok~10~-10000000~0~10000000~</S><S>dai ly~nguoiky~ict01~29/08/2013~1~~~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
'    str2 = "aa320823600247325   002012005006002002<S01-1><S>cmc~0102030405~~~~~~~0~10000000~2000000</S><S>0~2000000</S></S01-1>"
'    Barcode_Scaned str2
    
    'to khai bo sung
'    str2 = "bs320823600247325   00201200300300100201/0101/01/190001/01/201231/12/2012<S01><S></S><S>HD01~06/06/2013</S><S>10~10000000~200000000~190000000~ok~~5000000~70000000~0~~~0~0~0~ok~~10000000~200000000~190000000~~1000000~10000000~0~-10000000~ok~2000000~20000000~0~-20000000~ok~10~-10000000~0~10000000~</S><S>dai ly~nguoiky~ict01~29/08/2013~~1~1~01/01/2012~31/12/2012</S></S01>"
'    Barcode_Scaned str2
'    str2 = "bs320823600247325   002012003003002002<SKHBS><S>Thu� TNDN ph�i n�p~4~180000000~200000000~20000000</S><S>~~0~0~0</S><S>29/08/2013~196~1960000~~20000000</S></SKHBS>"
'    Barcode_Scaned str2
    
    'Thue tai nguyen 01/TAIN
    'str2 = "aa320923600247325   07201300100100100201/0114/06/2006<S01><S></S><S>010102~Kg~1000.000~10.00~10.000~0~10.00</S><S>010103~Kg~"
    'Barcode_Scaned str2
    'str2 = "aa320923600247325   0720130010010020022000.000~20.00~11.000~0~20.00</S><S>nhanvien~ict01~giamdoc~29/08/2013~1~~0</S></S01>"
    'Barcode_Scaned str2
    
    'Thue tai nguyen 01/TAIN-DK
'    str2 = "aa320923600247325   08201300100100100101/0114/06/2006<S01><S></S><S>1~x~x~~0~1~0~1~~30/08/2013~HDKhiLo001</S><S>1000000~1000000~12000~12000000000~20~2400000000~21200</S><S>daily~ict01~nguoiky~30/08/2013</S></S01>"
'    Barcode_Scaned str2

    'Thue tai nguyen 02/TAIN-DK
'    str2 = "aa320933600247325   00201300100100100201/0114/06/2006<S01><S></S><S>100~1000~2000~2000~1000~1000~1000~20~10000000~10~12000~200~200~12000~5000~605"
'    Barcode_Scaned str2
'    str2 = "aa320933600247325   00201300100100200200000~60000000~500000~100000000~-39500000</S><S>hoten~ict01~nguoiky~04/09/2013~0~0~1~1~~2012~HD2</S></S01>"
'    Barcode_Scaned str2
    
    'Thue 01/TD-GTGT
'    str2 = "aa320943600247325   08201300100100100101/0114/06/2006<S01><S></S><S>1000000~1200~1200000000~10~120000000~100000000~20000000</S><S>hoten~nguoiky~ict01~05/09/2013~1~~</S></S01>"
'    Barcode_Scaned str2
    'Phu luc 1-2
    
    str2 = "aa320943600247325   08201302702700100201/0114/06/2006<S01><S>0102030405</S><S>1000000000~1500~1500000000000~10~150000000000~0~150000000000</S><S>Lenh Ho Xung~Tieu Ngao Giang Ho~ICT01~16/09/2013~1~~</S></S01>"
    Barcode_Scaned str2
    str2 = "aa320943600247325   082013027027002002<S01_2><S>0001~CMCser 01~C�c Thu� Th�nh ph� H� N�i~20~100000000~10100~0001~CMCser 02~Chi c�c Thu� Qu�n H�ng B�ng~10~50000000~10301</S><S>150000000</S></S01_2>"
    Barcode_Scaned str2

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
            
                ' Check version <= 3.1.6
                If Val(Left$(strData, 3)) <= 316 Then
                    If Mid$(strData, 4, 2) = "01" Or Mid$(strData, 4, 2) = "02" Or Mid$(strData, 4, 2) = "04" Or Mid$(strData, 4, 2) = "71" Or Mid$(strData, 4, 2) = "36" Then
                        strData = Left$(strData, Len(strData) - 10) & "~0" & Right$(strData, 10)
                    ElseIf Mid$(strData, 4, 2) = "68" Then
                        strData = Left$(strData, Len(strData) - 10) & "~1" & Right$(strData, 10)
                    ElseIf Mid$(strData, 4, 2) = "73" Then
                        strData = Left$(strData, Len(strData) - 10) & "~" & Right$(strData, 10)
                    End If
                End If

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
    Dim strID As String
    strID = Left$(strTaxReportInfo, 2)
    
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        TAX_Utilities_Srv_New.Month = Left$(strValue, 2)

        If strID = "01" Or strID = "02" Or strID = "04" Or strID = "71" Or strID = "36" Or strID = "68" Then
            TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
        Else
            TAX_Utilities_Srv_New.ThreeMonths = ""
        End If

    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = 1 Then

        If strID = "68" Then
            TAX_Utilities_Srv_New.Month = Left$(strValue, 2)
            TAX_Utilities_Srv_New.ThreeMonths = ""
        Else
            TAX_Utilities_Srv_New.Month = ""
            TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
        End If

        TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
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
    
    LoaiKyKK = LoaiToKhai(strData)
    
    'Gan gia tri ngay dau ky
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        dNgayDauKy = DateSerial(CInt(TAX_Utilities_Srv_New.Year), CInt(TAX_Utilities_Srv_New.Month), 1)
        dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)

        If Val(strIDBCTC) = 1 Or Val(strIDBCTC) = 2 Or Val(strIDBCTC) = 4 Or Val(strIDBCTC) = 71 Or Val(strIDBCTC) = 36 Then
            If LoaiKyKK = True Then
                dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
                dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
                dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
            End If

        End If

    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then

        If Val(strIDBCTC) = 68 And LoaiKyKK = False Then
            dNgayDauKy = DateSerial(CInt(TAX_Utilities_Srv_New.Year), CInt(TAX_Utilities_Srv_New.Month), 1)
            dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
            dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
        Else
            dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
            dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
            dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
        End If

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
            If Trim(strID) = "70" Or Trim(strID) = "81" Or Trim(strID) = "91" Then
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
    If (Val(strID) >= 64 And Val(strID) <= 68) Or Val(strID) = 91 Then

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
    
'     Kiem tra to khai ton tai theo mau cu QLT
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
        
    ElseIf strID = "04_TBAC" Then
        arrDeltail = Split(strData, "~")
        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
        "And tkhai.LOAI_BC = '" & strID & "' " & _
        " And tkhai.NGAY_BC = to_date('" & arrDeltail(UBound(arrDeltail) - 1) & "','dd/mm/rrrr')" & _
        " And tkhai.NGAY_TB_PH = to_date('" & Right$(arrDeltail(UBound(arrDeltail) - 5), 10) & "','dd/mm/rrrr')"

    ElseIf strID = "BC26_AC" Then

        If LoaiKyKK = False Then
                 strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.LOAI_BC = '" & strID & "' " & _
                "And tkhai.QUY_BC = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
        Else
                 strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
                "And tkhai.LOAI_BC = '" & strID & "' " & _
                "And tkhai.KYBC_TU_NGAY = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
                "And tkhai.KYBC_DEN_NGAY = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"

        End If

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

Private Function LoaiToKhai(ByVal strData As String) As Boolean
    Dim LoaiTk As String
    
On Error GoTo ErrHandle
    strData = Left$(strData, Len(strData) - 10)
    LoaiTk = Right$(strData, 1)
    If LoaiTk = "1" Then
        LoaiToKhai = True
    Else
        LoaiToKhai = False
    End If
    Exit Function
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "LoaiToKhai", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

