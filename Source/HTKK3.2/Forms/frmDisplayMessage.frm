VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDisplayMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Box "
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4680
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
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
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
      Index           =   2
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
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
      Index           =   1
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   600
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
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSForms.CommandButton cmdButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1560
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
      Top             =   1560
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
      Top             =   1560
      Width           =   1305
      Size            =   "2302;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblMessage 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3945
      Caption         =   "Thông tin kê khai sai."
      Size            =   "6959;1085"
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
    SetMessage pMsgID
    
    Select Case pMsgStyle
        Case msAbortRetryIgnore
            SetControlCaption Me, "frmAbortRetryIgnore"
            SetButtonValue mrAbort, mrRetry, mrIgnore
            ButtonCount = 3
        Case msOKCancel
            SetControlCaption Me, "frmOKCancel"
            SetButtonValue , mrOK, mrCancel
            cmdButton1.Visible = False
            ButtonCount = 2
        Case msOKOnly
            SetControlCaption Me, "frmOKOnly"
            SetButtonValue , , mrOK
            cmdButton1.Visible = False
            cmdButton2.Visible = False
            ButtonCount = 1
        Case msOKCancelRetry
            SetControlCaption Me, "frmOKCancelRetry"
            SetButtonValue mrOK, mrCancel, mrRetry
            ButtonCount = 3
        Case msYesNo
            SetControlCaption Me, "frmYesNo"
            SetButtonValue , mrYes, mrNo
            cmdButton1.Visible = False
            ButtonCount = 2
        Case msYesNoCancel
            SetControlCaption Me, "frmYesNoCancel"
            SetButtonValue mrYes, mrNo, mrCancel
            ButtonCount = 3
    End Select
    
    With Me
        .caption = pTitle
        .picIcon(pIcon).Visible = True
        .picIcon(pIcon).Top = 300
        .picIcon(pIcon).Left = 150
        .picIcon(pIcon).Enabled = False
    End With
    
    ResizeMsgbox
    
    Me.Show vbModal
    
    DisplayMessage = msgResult
    
    Exit Function

ErrorHandle:
    SaveErrorLog Me.Name, "DisplayMessage", Err.Number, Err.Description
    
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
    
    lblMessage.caption = ""
    
    For Each xmlNode In xmlNodeListMessage
        If xmlNode.Attributes.getNamedItem("ID").nodeValue = pMsgID Then
            lblMessage.caption = xmlNode.Attributes.getNamedItem("Msg").nodeValue
'            lblMessage.Font.Size = 0.5
 '           lblMessage.AutoSize = True
            
            Exit For
        End If
    Next

    ResizeMsgbox
    Set xmlNode = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetMessage", Err.Number, Err.Description
    
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
    
'    If lblMessage.Width < cmdButton3.Width Then
'        lblMessage.Width = cmdButton3.Width
'    End If
    
'    picBtnContainer.Width = lblMessage.Width
'    picBtnContainer.Top = lblMessage.Top + lblMessage.Height + 500
    
'    Me.Width = lblMessage.Left + lblMessage.Width + 540
'    Me.Height = picBtnContainer.Top + picBtnContainer.Height + 500


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
    
    
'    For i = 1 To cmdButton.Count - 1
'        cmdButton(i).Left = cmdButton(i - 1).Left + cmdButton(0).Width + SPACE_BUTTON
'    Next
    
'    Call MyBorder
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ResizeMsgbox", Err.Number, Err.Description
    
End Sub


'****************************************************
'Description:MyBorder procedure draw border
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************

Public Sub MyBorder()
    On Error GoTo ErrorHandle
    
'    Line1.BorderColor = vb3DHighlight
'    Line1.X1 = Me.ScaleLeft + 60
'    Line1.X2 = Me.ScaleWidth - 60
'    Line1.Y1 = Me.ScaleTop + 60
'    Line1.Y2 = Me.ScaleTop + 60
    
'    Line2.BorderColor = vb3DHighlight
'    Line2.X1 = Me.ScaleLeft + 60
'    Line2.X2 = Me.ScaleLeft + 60
'    Line2.Y1 = Me.ScaleTop + 60
'    Line2.Y2 = Me.ScaleHeight - 60
    
'    Line3.BorderColor = vb3DDKShadow
'    Line3.X1 = Me.ScaleLeft + 80
'    Line3.X2 = Me.ScaleWidth - 63
'    Line3.Y1 = Me.ScaleHeight - 60
'    Line3.Y2 = Me.ScaleHeight - 60
    
'    Line4.BorderColor = vb3DDKShadow
'    Line4.X1 = Me.ScaleWidth - 60
'    Line4.X2 = Me.ScaleWidth - 60
'    Line4.Y1 = Me.ScaleTop + 61
'    Line4.Y2 = Me.ScaleHeight - 60
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "MyBorder", Err.Number, Err.Description
   
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandle
'    If ButtonCount <> 3 Then
'        cmdButton3.SetFocus
'    Else
'        If arrResult(0) = mrAbort Then
'            cmdButton1.SetFocus
'        ElseIf arrResult(1) = mrCancel Then
'            cmdButton2.SetFocus
'        Else
'            cmdButton3.SetFocus
'        End If
'    End If
    If ButtonCount = 1 Then
        cmdButton3.SetFocus
    ElseIf ButtonCount = 2 Then
        cmdButton2.SetFocus
    ElseIf ButtonCount = 3 Then
        cmdButton1.SetFocus
    End If
    
'    If cmdButton1 Then
'        cmdButton1.SetFocus
'    End If
ErrorHandle:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    picIcon(0).Picture = LoadPicture("..\Pictures\information.ico")
    picIcon(1).Picture = LoadPicture("..\Pictures\question.ico")
    picIcon(2).Picture = LoadPicture("..\Pictures\exclamation.ico")
    picIcon(3).Picture = LoadPicture("..\Pictures\critical.ico")
    
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
        
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description

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
    SaveErrorLog Me.Name, "SetButtonValue", Err.Number, Err.Description
    
End Sub

