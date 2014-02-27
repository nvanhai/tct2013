VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C?u hình"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3795
   Begin VB.TextBox txtPortWs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Text            =   "PortWs"
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdThoat 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1470
   End
   Begin VB.CommandButton cmdDongY 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtIpServer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblPortServices 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Services"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblQueueManager 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP may chu truc ESB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1440
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
If ValidateIP(txtIpServer.Text) Then
    SaveConfig
    Unload Me
    frmLogin.Show
End If
End Sub

Private Sub cmdThoat_Click()
Unload Me
frmLogin.Show
End Sub
Private Function ValidateIP(ByVal IPAddress As String) As Boolean
        Dim count As Byte
        Dim dotcount As Byte

        'check for illegal charaters
        For count = 1 To Len(IPAddress)
            If InStr("1234567890.", LCase(Mid(IPAddress, count, 1))) > 0 Then
            Else
                MsgBox ("There are illegal characters")
                ValidateIP = False
                txtIpServer.SetFocus
                Exit Function
            End If
        Next
        'check if first character is "."

        If InStr(IPAddress, ".") = 1 Then
            MsgBox ("First Character is '.'")
            ValidateIP = False
            txtIpServer.SetFocus
            Exit Function
        End If

        'check if there are consecutive ".."

        If InStr(IPAddress, "..") > 0 Then
            MsgBox ("There are consecutive '.'")
            ValidateIP = False
            txtIpServer.SetFocus
            Exit Function
        End If


        'check for number of dots
        For count = 1 To Len(IPAddress)
            If Mid(IPAddress, count, 1) = "." Then

                dotcount = dotcount + 1
                If dotcount > 3 Then
                    MsgBox ("There are two many '.'")
                    ValidateIP = False
                    txtIpServer.SetFocus
                    Exit Function
                End If

            End If
        Next

        'check for values of ip address components
        Dim num() As String
        num = Split(IPAddress, ".")
        
        For count = 0 To 3
            If (num(count)) > 255 Then
                MsgBox ("IP address is invalid")
                ValidateIP = False
                txtIpServer.SetFocus
                Exit Function
            End If

            'checks if last split is = 255
            If num(3) = 255 Then
                MsgBox ("IP address is invalid")
                ValidateIP = False
                txtIpServer.SetFocus
                Exit Function
            End If
        Next
        
'        'Check port
'        If Not (IsNumeric(txtPortWs.Text)) Then
'            MsgBox ("Port services is invalid")
'            txtPortWs.SetFocus
'            ValidateIP = False
'            Exit Function
'        End If
        
        'MsgBox ("Valid IP address")
        ValidateIP = True
        'if all of these things are true return true

    End Function
Private Sub SaveConfig()
On Error GoTo ErrHandle

Dim xmlTempConfig As New MSXML.DOMDocument
xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")

xmlTempConfig.getElementsByTagName("queue_manager_name")(0).Text = Base64Unicode.Base64EncodeString("ESB.INT.QMGR")
xmlTempConfig.getElementsByTagName("queue_name")(0).Text = Base64Unicode.Base64EncodeString("INT.INBOX.QUEUE")
xmlTempConfig.getElementsByTagName("BCTC")(0).Text = Base64Unicode.Base64EncodeString("69;19;20;21;22")
xmlTempConfig.getElementsByTagName("QLAC")(0).Text = Base64Unicode.Base64EncodeString("64;65;66;67;68;91")
xmlTempConfig.getElementsByTagName("TRAN_CODE")(0).Text = Base64Unicode.Base64EncodeString("01000")
xmlTempConfig.getElementsByTagName("RECEIVER_CODE")(0).Text = Base64Unicode.Base64EncodeString("TMS")

'NSD
Dim param As String
param = "http://" & txtIpServer.Text & ":7080/wsCBTVerify" '10.64.112.155
xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "http://tempuri.org/verify"
xmlTempConfig.getElementsByTagName("SoapActionNSD")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:q0=""http://tempuri.org/ESB_TCT_INTERNAL_MSG"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & _
        "    <soapenv:Body>" & _
        "        <q0:UserVerifyWSRequest>string</q0:UserVerifyWSRequest>" & _
        "    </soapenv:Body>" & _
        "</soapenv:Envelope>"
xmlTempConfig.getElementsByTagName("XmlRequestNSD")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "UserVerifyWSRequest"
xmlTempConfig.getElementsByTagName("ParamNameNSD")(0).Text = Base64Unicode.Base64EncodeString(param)

'NNT
param = "http://" & txtIpServer.Text & ":7080/wsDangKyThue" 'http://10.64.112.155:7080/wsDangKyThue
xmlTempConfig.getElementsByTagName("WsUrlNNT")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "http://gdt.gov.vn/getNNT"
xmlTempConfig.getElementsByTagName("SoapActionNNT")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:q0=""http://tempuri.org/ESB_TCT_INTERNAL_MSG"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & _
        "    <soapenv:Body>" & _
        "        <q0:NNTWSRequest>string</q0:NNTWSRequest>" & _
        "    </soapenv:Body>" & _
        "</soapenv:Envelope>"
xmlTempConfig.getElementsByTagName("XmlRequestNNT")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "NNTWSRequest"
xmlTempConfig.getElementsByTagName("ParamNameNNT")(0).Text = Base64Unicode.Base64EncodeString(param)

'DLT
param = "http://" & txtIpServer.Text & ":7080/wsDangKyThue" '10.64.112.155
xmlTempConfig.getElementsByTagName("WsUrlDLT")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "http://gdt.gov.vn/getDLT"
xmlTempConfig.getElementsByTagName("SoapActionDLT")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:q0=""http://tempuri.org/ESB_TCT_INTERNAL_MSG"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & _
        "    <soapenv:Body>" & _
        "        <q0:DLTWSRequest>string</q0:DLTWSRequest>" & _
        "    </soapenv:Body>" & _
        "</soapenv:Envelope>"
xmlTempConfig.getElementsByTagName("XmlRequestDLT")(0).Text = Base64Unicode.Base64EncodeString(param)
param = "DLTWSRequest"
xmlTempConfig.getElementsByTagName("ParamNameDLT")(0).Text = Base64Unicode.Base64EncodeString(param)

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
    SetControlCaption Me, "frmSetting"
   frmSetting.Top = (frmSystem.Height - frmLogin.Height) / 2
    frmSetting.Left = (frmSystem.Width - frmLogin.Width) / 2
LoadConfig
End Sub

Private Sub LoadConfig()
    Dim xmlTempConfig As New MSXML.DOMDocument
    On Error GoTo ErrHandle
    Dim sUrl As String

    xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")
    sUrl = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text)
    sUrl = Mid$(sUrl, InStr(1, sUrl, "//", vbTextCompare) + 2)
    sUrl = Mid$(sUrl, 1, InStr(1, sUrl, ":", vbTextCompare) - 1)
    txtIpServer.Text = sUrl
    txtPortWs.Text = ""
ErrHandle:
    SaveErrorLog Me.Name, "LoadConfig", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload Me
    frmLogin.Show
End Sub
