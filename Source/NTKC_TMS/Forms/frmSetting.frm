VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSetting 
   BorderStyle     =   0  'None
   Caption         =   "Setting"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
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
   ScaleHeight     =   9045
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdThoat 
      Caption         =   "Exit"
      Height          =   360
      Left            =   6960
      TabIndex        =   18
      Top             =   8520
      Width           =   1470
   End
   Begin VB.CommandButton cmdDongY 
      Caption         =   "Ok"
      Height          =   360
      Left            =   5280
      TabIndex        =   17
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox txtIDQLAC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   16
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtIDBCTC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtQueuename 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtQueueMgr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Frame frmWs 
      BackColor       =   &H80000005&
      Caption         =   "Connect Webservice"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.TextBox txtParamName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   3120
         Width           =   6855
      End
      Begin VB.TextBox txtRequest 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   6855
      End
      Begin VB.TextBox txtSoapAct 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   6855
      End
      Begin MSForms.OptionButton OptDLT 
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   360
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1508;450"
         Value           =   "0"
         Caption         =   "DLT"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OptNNT 
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1508;450"
         Value           =   "0"
         Caption         =   "NNT"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OptNSD 
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1508;450"
         Value           =   "1"
         Caption         =   "NSD"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblParamName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ParamName"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblRequest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XML Request"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblSoapAction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SoapAction"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Url Ws"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   465
      End
   End
   Begin MSForms.TabStrip NSD 
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   8415
      ListIndex       =   0
      Size            =   "14843;6165"
      Items           =   "Tab1;Tab2;"
      MultiRow        =   -1  'True
      TipStrings      =   ";;"
      Names           =   "Tab1;Tab2;"
      NewVersion      =   -1  'True
      TabsAllocated   =   2
      Tags            =   ";;"
      TabData         =   2
      Accelerator     =   ";;"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      TabState        =   "3;3"
   End
   Begin VB.Label lblIDQLAC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID QLAC"
      Height          =   195
      Left            =   4920
      TabIndex        =   12
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblIDBCTC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID BCTC"
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label lblQueueName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Name"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   930
   End
   Begin VB.Label lblQueueManager 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Manager"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3840
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


'CreateConfig
'Dim sBase64 As String
'sBase64 = Base64Unicode.Base64EncodeString(txtRequest.Text)
'MsgBox sBase64
'sBase64 = Base64Unicode.Base64DecodeString(sBase64)
'MsgBox sBase64
'Unload Me
End Sub

Private Sub cmdThoat_Click()
Unload Me
End Sub
Private Function CreateConfig()
Dim xmlConfig As New MSXML.DOMDocument
On Error GoTo ErrHandle

xmlConfig.Load GetAbsolutePath("..\Project\Config.xml")
txtUrl.Text = xmlConfig.getElementsByTagName("WsUrlNSD")(0).Text

ErrHandle:
    SaveErrorLog Me.Name, "CreateConfig", Err.Number, Err.Description
End Function

Private Sub Form_Load()
NSD.Tabs.Add "tabDLT", "DLT", 2
NSD.Tabs(0).caption = "NSD"
NSD.Tabs(1).caption = "NNT"



End Sub
