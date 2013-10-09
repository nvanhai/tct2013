VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDisplayMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   -30
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSForms.Label lblCaption 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   2535
      ForeColor       =   -2147483634
      Size            =   "4471;661"
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
      Width           =   4695
   End
   Begin MSForms.CommandButton cmdButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1665
      Width           =   1305
      Size            =   "2302;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdButton2 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1665
      Width           =   1305
      Size            =   "2302;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdButton3 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1665
      Width           =   1305
      Size            =   "2302;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblMessage 
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   705
      Width           =   3915
      Caption         =   "Thông tin kê khai sai."
      Size            =   "6906;1508"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmDisplayMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : FPT Software Solution (FSS)
' Project           : Du an ho tro ke khai thue
' Package           : Interface
' Form, Module
'   or Class name   : frmDisplayMessage
' Descriptions      : Report sh
' Start date        : 11/10/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : TuanLM
' Integrate         :
' Project manager   : ThietKN
' Last modify       :
' Reason of modify  :
'******************************************************

Option Explicit

Private msgResult As MsgBoxResult ' result of button that user click
Private arrResult(3) As MsgBoxResult ' array values of buttons on the form
Private ButtonCount As Integer ' number of button on the form
Private Const SPACE_BUTTON = 50 ' space between buttons

'****************************************************
'Description:cmdButton_Click procedure return the value of button
'   that user clicked
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub cmdButton1_Click()
    msgResult = arrResult(0)
    Unload Me
End Sub

Private Sub cmdButton2_Click()
    msgResult = arrResult(1)
    Unload Me
End Sub

Private Sub cmdButton3_Click()
    msgResult = arrResult(2)
    Unload Me
End Sub

'****************************************************
'Description:DisplayMessage procedure display message
'   Step 1: Load message with id is pMsgID
'   Step 2: Show buttons
'   Step 3: Set value for buttons
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input: pMsgID - ID of message in file Message.xml
'       pMsgStyle - style of MessageBox
'       pIcon - icon of MessageBox
'       pTitle - title of MessageBox
'Output:
'Return: MsgBoxResult that user clicked

'****************************************************

Public Function DisplayMessage(pMsgID As String, Optional pMsgStyle As MsgBoxStyle, Optional pIcon As MsgBoxIcon, Optional pTitle As String) As MsgBoxResult
    On Error GoTo ErrorHandle
    Dim clsMess As New clsMessageBox
    
    SetMessage pMsgID
    
    Select Case pMsgStyle
        Case msAbortRetryIgnore
            clsMess.SetControlCaption Me, "frmAbortRetryIgnore"
            SetButtonValue mrAbort, mrRetry, mrIgnore
            ButtonCount = 3
        Case msOKCancel
            clsMess.SetControlCaption Me, "frmOKCancel"
            SetButtonValue , mrOK, mrCancel
            cmdButton1.Visible = False
            ButtonCount = 2
        Case msOKOnly
            clsMess.SetControlCaption Me, "frmOKOnly"
            SetButtonValue , , mrOK
            cmdButton1.Visible = False
            cmdButton2.Visible = False
            ButtonCount = 1
        Case msOKCancelRetry
            clsMess.SetControlCaption Me, "frmOKCancelRetry"
            SetButtonValue mrOK, mrCancel, mrRetry
            ButtonCount = 3
        Case msYesNo
            clsMess.SetControlCaption Me, "frmYesNo"
            SetButtonValue , mrYes, mrNo
            cmdButton1.Visible = False
            ButtonCount = 2
        Case msYesNoCancel
            clsMess.SetControlCaption Me, "frmYesNoCancel"
            SetButtonValue mrYes, mrNo, mrCancel
            ButtonCount = 3
    End Select
    
    With Me
        .Caption = vbNullString
        .picIcon(pIcon).Visible = True
        .picIcon(pIcon).Top = 620
        .picIcon(pIcon).Left = 200
        .picIcon(pIcon).Enabled = False
    End With
    
    ResizeMsgbox
    
    Me.Show vbModal
    
    DisplayMessage = msgResult
    
    Set clsMess = Nothing
    
    Exit Function

ErrorHandle:
End Function

'****************************************************
'Description:SetMessage procedure set value of message
'   Step 1: Load message with id is pMsgID
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input: pMsgID - ID of message in file Message.xml
'Output:
'Return:

'****************************************************

Private Sub SetMessage(pMsgID As String)
    On Error GoTo ErrorHandle
    Dim xmlNode As MSXML.IXMLDOMNode
    
    lblMessage.Caption = ""
    
    For Each xmlNode In xmlNodeListMessage
        If xmlNode.Attributes.getNamedItem("ID").nodeValue = pMsgID Then
            lblMessage.Caption = xmlNode.Attributes.getNamedItem("Msg").nodeValue
            Exit For
        End If
    Next

    ResizeMsgbox
    Set xmlNode = Nothing
    
    Exit Sub
ErrorHandle:
End Sub


'****************************************************
'Description:ResizeMsgbox procedure resize messagebox
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub ResizeMsgbox()
    On Error GoTo ErrorHandle
    
    Dim i As Integer

    If ButtonCount = 1 Then
        cmdButton3.Left = (Me.Width - cmdButton1.Width) \ 2
    ElseIf ButtonCount = 2 Then
        cmdButton2.Left = Me.Width / 2 - cmdButton1.Width - SPACE_BUTTON
        cmdButton3.Left = cmdButton2.Left + cmdButton1.Width + SPACE_BUTTON
        cmdButton2.TabIndex = 0
        cmdButton3.TabIndex = 1
        cmdButton1.TabIndex = -1
    Else
        cmdButton1.Left = Me.Width / 2 - cmdButton1.Width * 3 / 2 - SPACE_BUTTON
        cmdButton2.Left = cmdButton1.Left + cmdButton1.Width + SPACE_BUTTON
        cmdButton3.Left = cmdButton2.Left + cmdButton1.Width + SPACE_BUTTON
        cmdButton1.TabIndex = 0
        cmdButton2.TabIndex = 1
        cmdButton3.TabIndex = 2
    End If
    
    Exit Sub
ErrorHandle:
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandle
    If ButtonCount = 1 Then
        cmdButton3.SetFocus
    ElseIf ButtonCount = 2 Then
        cmdButton2.SetFocus
    ElseIf ButtonCount = 3 Then
        cmdButton1.SetFocus
    End If
ErrorHandle:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim clsMess As New clsMessageBox
    
    picIcon(0).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\information.gif"))
    picIcon(1).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\question.gif"))
    picIcon(2).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\exclamation.gif"))
    picIcon(3).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\critical.gif"))
    
    Exit Sub
    
ErrorHandle:
End Sub

'****************************************************
'Description:SetButtonValue set value for button
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input: msg1 - value of button 1
'       msg1 - value of button 2
'       msg1 - value of button 3
'Output:
'Return:

'****************************************************

Private Sub SetButtonValue(Optional msg1 As MsgBoxResult, Optional msg2 As MsgBoxResult, Optional msg3 As MsgBoxResult)
    On Error GoTo ErrorHandle
    arrResult(0) = msg1
    arrResult(1) = msg2
    arrResult(2) = msg3
    
    Exit Sub
ErrorHandle:
End Sub

Private Sub Form_Resize()
    Dim clsMess As New clsMessageBox
    clsMess.SetFormCaption Me, imgCaption, lblCaption
End Sub
