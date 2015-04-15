VERSION 5.00
Begin VB.Form frmBrowser 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "§ã&ng"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   5340
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "§ån&g ý"
      Default         =   -1  'True
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
      Left            =   2970
      TabIndex        =   3
      Top             =   5340
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   4965
      Left            =   0
      TabIndex        =   0
      Top             =   270
      Width           =   5775
      Begin VB.DirListBox Dir1 
         Height          =   4365
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   5535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   5535
      End
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Chän ®­êng dÉn ..."
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
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmBrowser
' Descriptions      : Report sh
' Start date        : 21/10/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : htphuong
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************

Option Explicit
Private strPath  As String

Private Sub cmdClose_Click()
    strPath = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Err.Clear
    On Error GoTo ErrorHandle
    
    strPath = Dir1.path
    Unload Me
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
End Sub

Private Sub Drive1_Change()
    Dim strDrive As String
    On Error GoTo ErrorHandle
    
    strDrive = Left(Dir1.path, InStr(Dir1.path, "\") - 1)
    Dir1.path = Drive1.Drive
    
    Exit Sub
ErrorHandle:
    If Err.Number = 68 Then
        DisplayMessage "0031", msOKOnly, miCriticalError
        Drive1.Drive = strDrive
    ElseIf Err.Number = 419 Then
        DisplayMessage "0046", msOKOnly, miCriticalError
        Drive1.Drive = strDrive
    Else
        SaveErrorLog Me.Name, "Drive1_Change", Err.Number, Err.Description
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Me.Top = frmSystem.Top + (frmSystem.Height - Me.Height) / 2
    Me.Left = frmSystem.Left + (frmSystem.Width - Me.Width) / 2
    SetControlCaption Me, "frmBrowser"
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
    
End Sub

'****************************************************
'Description:getPath function chose the path to the
'   directory to backup or restore data
'****************************************************

Public Function getPath() As String
    On Error GoTo ErrorHandle
    
    Me.Show vbModal
    getPath = strPath
    
    Exit Function
     
ErrorHandle:
    SaveErrorLog Me.Name, "getPath", Err.Number, Err.Description
    
End Function

'****************************************************
'Description:Form_KeyUp procedure process keyup event
'       When user press Alt + F4 -> process Exit
'Input: KeyCode: vbKeyCode
'       Shift: Ctrl or Alt or Shift key
'****************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 4 Then
        cmdClose_Click
    End If
End Sub

Private Sub Form_Resize()
     SetFormCaption Me, imgCaption, lblCaption
End Sub

