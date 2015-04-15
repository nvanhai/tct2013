VERSION 5.00
Begin VB.Form frmDisplayMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   -30
   ClientWidth     =   4785
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
   ScaleHeight     =   2160
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton3 
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1680
      Width           =   1305
   End
   Begin VB.CommandButton cmdButton2 
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1305
   End
   Begin VB.CommandButton cmdButton1 
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1305
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DS Sans Serif"
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
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DS Sans Serif"
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
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DS Sans Serif"
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
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DS Sans Serif"
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
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   840
      TabIndex        =   5
      Top             =   540
      Width           =   3735
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmDisplayMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Software
' Center Name       : FIS (Financial Insurance Solution)
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmDisplayMessage
' Descriptions      : Report sh
' Start date        : 10/08/2007 (dd/mm/yyyy)
' Finish date       :
' Coder             : hlnam
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************

Option Explicit

Private msgResult As MsgBoxResult           ' Result of button that user click
Private arrResult(3) As MsgBoxResult        ' Array values of buttons on the form
Private ButtonCount As Integer              ' Number of button on the form
Private udtDefaultButton As MsgBoxResult    ' Default button
Private Const SPACE_BUTTON = 50             ' space between buttons

'*****************************************************
'Description : cmdButton_Click procedure return the value of button that user clicked
'Author     : hlnam
'Modify by  : hlnam
'Date       :10/08/2007
'Input      :
'Output     :
'Return     :

'*****************************************************

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

'*****************************************************
'Description    : DisplayMessage procedure display message
'   Step 1 : Load message with id is pMsgID
'   Step 2 : Show buttons
'   Step 3 : Set value for buttons
'Author         : hlnam
'Modify by      :
'Date           : 10/08/2007
'Input          : pMsgID - ID of message in file Message.xml
'       pMsgStyle - style of MessageBox
'       pIcon - icon of MessageBox
'       pTitle - title of MessageBox
'Output         :
'Return         : MsgBoxResult that user clicked

'*****************************************************

Public Function DisplayMessage(pMsgID As String, Optional pMsgStyle As MsgBoxStyle, Optional pIcon As MsgBoxIcon, Optional pTitle As String, Optional pDefaultMsgBoxStyle As MsgBoxResult) As MsgBoxResult
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
    
    'Set default button
    udtDefaultButton = pDefaultMsgBoxStyle
    
    With Me
        .Caption = vbNullString
        .picIcon(pIcon).Visible = True
        .picIcon(pIcon).Top = 620
        .picIcon(pIcon).Left = 150
        .picIcon(pIcon).Enabled = False
    End With
    
    ResizeMsgbox
    
    Me.Show vbModal
    
    DisplayMessage = msgResult
    
    Set clsMess = Nothing
    
    Exit Function

ErrorHandle:
End Function

'*****************************************************
'Description    : SetMessage procedure set value of message
'  Step 1  : Load message with id is pMsgID
'Author         : hlnam
'Modify by      :
'Date           : 10/08/2007
'Input          : pMsgID - ID of message in file Message.xml
'Output         :
'Return         :
'*****************************************************

Private Sub SetMessage(pMsgID As String)
    On Error GoTo ErrorHandle
    Dim xmlNode As MSXML.IXMLDOMNode
    
    lblMessage.Caption = ""
    ' hien thi them message text truyen vao
    If Len(pMsgID) > 4 And Left(pMsgID, 3) = "###" Then
        lblMessage.Caption = Mid$(pMsgID, 4)
    Else
        For Each xmlNode In xmlNodeListMessage
            If xmlNode.Attributes.getNamedItem("ID").nodeValue = pMsgID Then
                lblMessage.Caption = xmlNode.Attributes.getNamedItem("Msg").nodeValue
                Exit For
            End If
        Next
    End If
        
    ResizeMsgbox
    Set xmlNode = Nothing
    
    Exit Sub
ErrorHandle:
End Sub


'*****************************************************
'Description    : ResizeMsgbox procedure resize messagebox
'Author         : hlnam
'Modify by      :
'Date           : 10/08/2007
'Input          :
'Output         :
'Return         :

'*****************************************************

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

    If udtDefaultButton = arrResult(0) Then
        cmdButton1.SetFocus
    ElseIf udtDefaultButton = arrResult(1) Then
        cmdButton2.SetFocus
    ElseIf udtDefaultButton = arrResult(2) Then
        cmdButton3.SetFocus
    Else
        If ButtonCount = 1 Then
            cmdButton3.SetFocus
        ElseIf ButtonCount = 2 Then
            cmdButton2.SetFocus
        ElseIf ButtonCount = 3 Then
            cmdButton1.SetFocus
        End If
    End If
    Exit Sub
ErrorHandle:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim clsMess As New clsMessageBox
    
    picIcon(0).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\information.ico"))
    picIcon(1).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\question.ico"))
    picIcon(2).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\exclamation.ico"))
    picIcon(3).Picture = LoadPicture(clsMess.GetAbsolutePath("..\Pictures\critical.ico"))
    
    Set clsMess = Nothing
    Exit Sub
    
ErrorHandle:
    Set clsMess = Nothing
End Sub

'*****************************************************
'Description    : SetButtonValue set value for button
'Author         : hlnam
'Modify by      :
'Date           : 10/08/2007
'Input  : msg1 - value of button 1
'         msg1 - value of button 2
'         msg1 - value of button 3
'Output         :
'Return         :

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
