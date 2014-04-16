VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSetting 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1785
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4470
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   4365
      Begin MSForms.TextBox txtIpServer 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   300
         Width           =   1695
         VariousPropertyBits=   746604571
         Size            =   "2990;556"
         Value           =   "10.64.85.170"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblQueueManager 
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2205
         VariousPropertyBits=   276824091
         Caption         =   "Username"
         Size            =   "3889;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
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
      Left            =   120
      TabIndex        =   1
      Text            =   "PortWs"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   330
      TabIndex        =   7
      Top             =   60
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
   Begin MSForms.CommandButton cmdDongY 
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   1305
      Caption         =   "Login"
      Size            =   "2302;661"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdThoat 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3090
      TabIndex        =   5
      Top             =   1320
      Width           =   1305
      Caption         =   "Exit"
      Size            =   "2302;661"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
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
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image imgCaption 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   3915
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
    SaveConfigBase64
    Unload Me
    frmLogin.Show
End If
End Sub

Private Sub cmdThoat_Click()
Unload Me
frmLogin.Show
End Sub
Private Function MessageBox(strMsgId As String, intMsgStyle As MsgBoxStyle, intMsgIcon As MsgBoxIcon, Optional msType As Byte) As MsgBoxResult
    Dim intReturn As Integer
    
On Error GoTo ErrHandle
    
    
    MessageBox = DisplayMessage(strMsgId, intMsgStyle, intMsgIcon, , msType)
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "MessageBox", Err.Number, Err.Description
End Function
Private Function ValidateIP(ByVal IPAddress As String) As Boolean
        Dim Count As Byte
        Dim dotcount As Byte

        'check for illegal charaters
        For Count = 1 To Len(IPAddress)
            If InStr("1234567890.", LCase(Mid(IPAddress, Count, 1))) > 0 Then
            Else
                MessageBox "0152", msOKOnly, miCriticalError
                'MsgBox ("There are illegal characters")
                ValidateIP = False
                txtIpServer.SetFocus
                Exit Function
            End If
        Next
        'check if first character is "."

        If InStr(IPAddress, ".") = 1 Then
            'MsgBox ("First Character is '.'")
            MessageBox "0152", msOKOnly, miCriticalError
            ValidateIP = False
            txtIpServer.SetFocus
            Exit Function
        End If

        'check if there are consecutive ".."

        If InStr(IPAddress, "..") > 0 Then
            'MsgBox ("There are consecutive '.'")
            MessageBox "0152", msOKOnly, miCriticalError
            ValidateIP = False
            txtIpServer.SetFocus
            Exit Function
        End If


        'check for number of dots
        For Count = 1 To Len(IPAddress)
            If Mid(IPAddress, Count, 1) = "." Then

                dotcount = dotcount + 1
                If dotcount > 3 Then
                    'MsgBox ("There are two many '.'")
                    MessageBox "0152", msOKOnly, miCriticalError
                    ValidateIP = False
                    txtIpServer.SetFocus
                    Exit Function
                End If

            End If
        Next

        'check for values of ip address components
        Dim num() As String
        num = Split(IPAddress, ".")
        
        For Count = 0 To 3
            If (num(Count)) > 255 Then
                'MsgBox ("IP address is invalid")
                MessageBox "0152", msOKOnly, miCriticalError
                ValidateIP = False
                txtIpServer.SetFocus
                Exit Function
            End If

            'checks if last split is = 255
            If num(3) = 255 Then
                'MsgBox ("IP address is invalid")
                MessageBox "0152", msOKOnly, miCriticalError
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
Private Sub SaveConfigBase64()
On Error GoTo ErrHandle

Dim xmlTempConfig As New MSXML.DOMDocument
xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")

xmlTempConfig.getElementsByTagName("queue_manager_name")(0).Text = Base64Unicode.Base64EncodeString("ESB.INT.QMGR")
xmlTempConfig.getElementsByTagName("queue_name")(0).Text = Base64Unicode.Base64EncodeString("INT.INBOX.QUEUE")
xmlTempConfig.getElementsByTagName("BCTC")(0).Text = Base64Unicode.Base64EncodeString("69;19;20;21;22")
xmlTempConfig.getElementsByTagName("QLAC")(0).Text = Base64Unicode.Base64EncodeString("64;65;66;67;68;91;07;09;10;13;14")
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
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SaveConfig", Err.Number, Err.Description
End Sub
Private Sub SaveConfig()
On Error GoTo ErrHandle

Dim xmlTempConfig As New MSXML.DOMDocument
xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")

xmlTempConfig.getElementsByTagName("queue_manager_name")(0).Text = "ESB.INT.QMGR"
xmlTempConfig.getElementsByTagName("queue_name")(0).Text = "INT.INBOX.QUEUE"
xmlTempConfig.getElementsByTagName("BCTC")(0).Text = "69;19;20;21;22"
xmlTempConfig.getElementsByTagName("QLAC")(0).Text = "64;65;66;67;68;91"
xmlTempConfig.getElementsByTagName("TRAN_CODE")(0).Text = "01000"
xmlTempConfig.getElementsByTagName("RECEIVER_CODE")(0).Text = "TMS"

'NSD
Dim param As String
param = "http://" & txtIpServer.Text & ":7080/wsCBTVerify" '10.64.112.155
xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text = param
param = "http://tempuri.org/verify"
xmlTempConfig.getElementsByTagName("SoapActionNSD")(0).Text = param
param = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:q0=""http://tempuri.org/ESB_TCT_INTERNAL_MSG"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & _
        "    <soapenv:Body>" & _
        "        <q0:UserVerifyWSRequest>string</q0:UserVerifyWSRequest>" & _
        "    </soapenv:Body>" & _
        "</soapenv:Envelope>"
xmlTempConfig.getElementsByTagName("XmlRequestNSD")(0).Text = param
param = "UserVerifyWSRequest"
xmlTempConfig.getElementsByTagName("ParamNameNSD")(0).Text = param

'NNT
param = "http://" & txtIpServer.Text & ":7080/wsDangKyThue" 'http://10.64.112.155:7080/wsDangKyThue
xmlTempConfig.getElementsByTagName("WsUrlNNT")(0).Text = param
param = "http://gdt.gov.vn/getNNT"
xmlTempConfig.getElementsByTagName("SoapActionNNT")(0).Text = param
param = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:q0=""http://tempuri.org/ESB_TCT_INTERNAL_MSG"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & _
        "    <soapenv:Body>" & _
        "        <q0:NNTWSRequest>string</q0:NNTWSRequest>" & _
        "    </soapenv:Body>" & _
        "</soapenv:Envelope>"
xmlTempConfig.getElementsByTagName("XmlRequestNNT")(0).Text = param
param = "NNTWSRequest"
xmlTempConfig.getElementsByTagName("ParamNameNNT")(0).Text = param

'DLT
param = "http://" & txtIpServer.Text & ":7080/wsDangKyThue" '10.64.112.155
xmlTempConfig.getElementsByTagName("WsUrlDLT")(0).Text = param
param = "http://gdt.gov.vn/getDLT"
xmlTempConfig.getElementsByTagName("SoapActionDLT")(0).Text = param
param = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:q0=""http://tempuri.org/ESB_TCT_INTERNAL_MSG"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & _
        "    <soapenv:Body>" & _
        "        <q0:DLTWSRequest>string</q0:DLTWSRequest>" & _
        "    </soapenv:Body>" & _
        "</soapenv:Envelope>"
xmlTempConfig.getElementsByTagName("XmlRequestDLT")(0).Text = param
param = "DLTWSRequest"
xmlTempConfig.getElementsByTagName("ParamNameDLT")(0).Text = param

'Set value const
xmlTempConfig.getElementsByTagName("VERSION")(0).Text = APP_VERSION
xmlTempConfig.getElementsByTagName("SENDER_CODE")(0).Text = "NTK"
xmlTempConfig.getElementsByTagName("SENDER_NAME")(0).Text = "He thong nhan to khai ma vach"
xmlTempConfig.getElementsByTagName("RECEIVER_NAME")(0).Text = "He thong quan ly thue tap trung"
xmlTempConfig.getElementsByTagName("ORIGINAL_CODE")(0).Text = "NTK"
xmlTempConfig.getElementsByTagName("ORIGINAL_NAME")(0).Text = "He thong nhan to khai ma vach"



Dim sFileName As String
sFileName = App.path & "\Config.xml"
xmlTempConfig.save sFileName
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SaveConfig", Err.Number, Err.Description
End Sub
Private Sub Form_Load()
    SetControlCaption Me, "frmSetting"
    frmSetting.Top = (frmSystem.Height - frmLogin.Height) / 2
    frmSetting.Left = (frmSystem.Width - frmLogin.Width) / 2
    'LoadConfig
End Sub

Private Sub LoadConfigBase64()
   
'    Dim xmlTempConfig As New MSXML.DOMDocument
'    Dim sUrl As String
'
'    xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")
'    sUrl = Base64Unicode.Base64DecodeString(xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text)
'    sUrl = Mid$(sUrl, InStr(1, sUrl, "//", vbTextCompare) + 2)
'    sUrl = Mid$(sUrl, 1, InStr(1, sUrl, ":", vbTextCompare) - 1)
'    txtIpServer.Text = IIf(sUrl <> "", sUrl, "10.64.85.167")
'    txtPortWs.Text = ""
End Sub

Private Sub LoadConfig()
   
'    Dim xmlTempConfig As New MSXML.DOMDocument
'    Dim sUrl As String
'
'    xmlTempConfig.Load GetAbsolutePath("..\Project\Config.xml")
'    sUrl = xmlTempConfig.getElementsByTagName("WsUrlNSD")(0).Text
'    sUrl = Mid$(sUrl, InStr(1, sUrl, "//", vbTextCompare) + 2)
'    sUrl = Mid$(sUrl, 1, InStr(1, sUrl, ":", vbTextCompare) - 1)
'    txtIpServer.Text = IIf(sUrl <> "", sUrl, "10.64.85.167")
'    txtPortWs.Text = ""
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload Me
    frmLogin.Show
End Sub
