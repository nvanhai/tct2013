VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   0  'None
   Caption         =   "Setting"
   ClientHeight    =   11565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   11565
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReceiverCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtTranCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame FraDLT 
      Caption         =   "DLT"
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   7680
      Width           =   10575
      Begin VB.TextBox txtParamDLT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   2760
         Width           =   9135
      End
      Begin VB.TextBox txtXmlRequestDLT 
         Appearance      =   0  'Flat
         Height          =   1735
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   9135
      End
      Begin VB.TextBox txtSoapActionDLT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   9135
      End
      Begin VB.TextBox txtUrlDLT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   9135
      End
      Begin VB.Label lblParameterDLT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter"
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label lblXmlRequestDLT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Xml Request"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblSoapActionDLT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soap Action"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblUrlDLT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Url"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame FraNNT 
      Caption         =   "NNT"
      Height          =   3135
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   10575
      Begin VB.TextBox txtParamNNT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   2760
         Width           =   9135
      End
      Begin VB.TextBox txtXmlRequestNNT 
         Appearance      =   0  'Flat
         Height          =   1725
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   9135
      End
      Begin VB.TextBox txtSoapActionNNT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   9135
      End
      Begin VB.TextBox txtUrlNNT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   9135
      End
      Begin VB.Label lblParameterNNT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter"
         Height          =   195
         Left            =   360
         TabIndex        =   35
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label lblXmlRequestNNT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XmlRequest"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblSoapActionNNT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soap Action"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblUrlNNT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Url"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.TextBox txtParam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   3960
      Width           =   9135
   End
   Begin VB.TextBox txtXmlRequest 
      Appearance      =   0  'Flat
      Height          =   1725
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   9135
   End
   Begin VB.TextBox txtSoapAction 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   9135
   End
   Begin VB.TextBox txtUrlNSD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   9135
   End
   Begin VB.CommandButton cmdThoat 
      Caption         =   "Exit"
      Height          =   360
      Left            =   9120
      TabIndex        =   20
      Top             =   11040
      Width           =   1470
   End
   Begin VB.CommandButton cmdDongY 
      Caption         =   "Ok"
      Height          =   360
      Left            =   7440
      TabIndex        =   19
      Top             =   11040
      Width           =   1455
   End
   Begin VB.TextBox txtIDQLAC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtIDBCTC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtQueuename 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtQueueMgr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame FraNSD 
      Caption         =   "NSD"
      Height          =   3135
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   10575
      Begin VB.Label lblParameter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label lblXmlRequest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XmlRequest"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblSoapAction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soap Action"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Url"
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Label lblReceiverCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver Code"
      Height          =   195
      Left            =   7560
      TabIndex        =   40
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label lblTranCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tran Code"
      Height          =   195
      Left            =   8040
      TabIndex        =   39
      Top             =   240
      Width           =   750
   End
   Begin VB.Label lblIDQLAC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID QLAC"
      Height          =   195
      Left            =   4200
      TabIndex        =   26
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblIDBCTC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID BCTC"
      Height          =   195
      Left            =   4200
      TabIndex        =   25
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblQueueName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Name"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   930
   End
   Begin VB.Label lblQueueManager 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Manager"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Base64Unicode As New Base64Unicode
Private Sub cmdDongY_Click()
'Save value to file config

SaveConfig

Unload Me
End Sub

Private Sub cmdThoat_Click()
Unload Me
End Sub
Private Sub SaveConfig()
On Error GoTo ErrHandle

Dim xmlTempConfig As New MSXML.DOMDocument
xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")

xmlTempConfig.getElementsByTagName("queue_manager_name")(0).Text = Base64Unicode.Base64EncodeString(txtQueueMgr.Text)
xmlTempConfig.getElementsByTagName("queue_name")(0).Text = Base64Unicode.Base64EncodeString(txtQueuename.Text)
xmlTempConfig.getElementsByTagName("BCTC")(0).Text = Base64Unicode.Base64EncodeString(txtIDBCTC.Text)
xmlTempConfig.getElementsByTagName("QLAC")(0).Text = Base64Unicode.Base64EncodeString(txtIDQLAC.Text)
xmlTempConfig.getElementsByTagName("TRAN_CODE")(0).Text = Base64Unicode.Base64EncodeString(txtTranCode.Text)
xmlTempConfig.getElementsByTagName("RECEIVER_CODE")(0).Text = Base64Unicode.Base64EncodeString(txtReceiverCode.Text)

'NSD
xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text = Base64Unicode.Base64EncodeString(txtUrlNSD.Text)
xmlTempConfig.getElementsByTagName("SoapActionNSD")(0).Text = Base64Unicode.Base64EncodeString(txtSoapAction.Text)
xmlTempConfig.getElementsByTagName("XmlRequestNSD")(0).Text = Base64Unicode.Base64EncodeString(txtXmlRequest.Text)
xmlTempConfig.getElementsByTagName("ParamNameNSD")(0).Text = Base64Unicode.Base64EncodeString(txtParam.Text)

'NNT
xmlTempConfig.getElementsByTagName("WsUrlNNT")(0).Text = Base64Unicode.Base64EncodeString(txtUrlNNT.Text)
xmlTempConfig.getElementsByTagName("SoapActionNNT")(0).Text = Base64Unicode.Base64EncodeString(txtSoapActionNNT.Text)
xmlTempConfig.getElementsByTagName("XmlRequestNNT")(0).Text = Base64Unicode.Base64EncodeString(txtXmlRequestNNT.Text)
xmlTempConfig.getElementsByTagName("ParamNameNNT")(0).Text = Base64Unicode.Base64EncodeString(txtParamNNT.Text)

'DLT
xmlTempConfig.getElementsByTagName("WsUrlDLT")(0).Text = Base64Unicode.Base64EncodeString(txtUrlDLT.Text)
xmlTempConfig.getElementsByTagName("SoapActionDLT")(0).Text = Base64Unicode.Base64EncodeString(txtSoapActionDLT.Text)
xmlTempConfig.getElementsByTagName("XmlRequestDLT")(0).Text = Base64Unicode.Base64EncodeString(txtXmlRequestDLT.Text)
xmlTempConfig.getElementsByTagName("ParamNameDLT")(0).Text = Base64Unicode.Base64EncodeString(txtParamDLT.Text)

'Set value const
xmlTempConfig.getElementsByTagName("VERSION")(0).Text = Base64Unicode.Base64EncodeString(APP_VERSION)
xmlTempConfig.getElementsByTagName("SENDER_CODE")(0).Text = Base64Unicode.Base64EncodeString("NTK")
xmlTempConfig.getElementsByTagName("SENDER_NAME")(0).Text = Base64Unicode.Base64EncodeString("He thong nhan to khai ma vach")
xmlTempConfig.getElementsByTagName("RECEIVER_NAME")(0).Text = Base64Unicode.Base64EncodeString("He thong quan ly thue tap trung")
xmlTempConfig.getElementsByTagName("ORIGINAL_CODE")(0).Text = Base64Unicode.Base64EncodeString("NTK")
xmlTempConfig.getElementsByTagName("ORIGINAL_NAME")(0).Text = Base64Unicode.Base64EncodeString("He thong nhan to khai ma vach")


Dim sFileName As String
sFileName = App.path & "\Config.xml"
xmlTempConfig.save sFileName
    
ErrHandle:
    SaveErrorLog Me.Name, "SaveConfig", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
LoadConfig
End Sub

Private Sub LoadConfig()
Dim xmlTempConfig As New MSXML.DOMDocument
On Error GoTo ErrHandle
Dim isDefaultConfig As String

xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")

txtQueueMgr.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("queue_manager_name")(0).Text)
txtQueuename.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("queue_name")(0).Text)
txtIDBCTC.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("BCTC")(0).Text)
txtIDQLAC.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("QLAC")(0).Text)
txtTranCode.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("TRAN_CODE")(0).Text)
txtReceiverCode.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("RECEIVER_CODE")(0).Text)


'NSD
txtUrlNSD.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text)
txtSoapAction.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("SoapActionNSD")(0).Text)
txtXmlRequest.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("XmlRequestNSD")(0).Text)
txtParam.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("ParamNameNSD")(0).Text)

'NNT
txtUrlNNT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("WsUrlNNT")(0).Text)
txtSoapActionNNT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("SoapActionNNT")(0).Text)
txtXmlRequestNNT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("XmlRequestNNT")(0).Text)
txtParamNNT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("ParamNameNNT")(0).Text)

'DLT
txtUrlDLT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("WsUrlDLT")(0).Text)
txtSoapActionDLT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("SoapActionDLT")(0).Text)
txtXmlRequestDLT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("XmlRequestDLT")(0).Text)
txtParamDLT.Text = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("ParamNameDLT")(0).Text)


ErrHandle:
    SaveErrorLog Me.Name, "LoadConfig", Err.Number, Err.Description
End Sub


