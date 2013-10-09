VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmChangepass 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   75
      TabIndex        =   0
      Top             =   390
      Width           =   4380
      Begin MSForms.Label lblNew 
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1020
         Width           =   1455
         VariousPropertyBits=   276824091
         Caption         =   "hehe"
         Size            =   "2566;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtNew 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   2175
         VariousPropertyBits=   746604571
         Size            =   "3836;556"
         PasswordChar    =   42
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPassword 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   630
         Width           =   2175
         VariousPropertyBits=   746604571
         Size            =   "3836;556"
         PasswordChar    =   42
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblPassword 
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   690
         Width           =   1545
         VariousPropertyBits=   276824091
         Caption         =   "Password"
         Size            =   "2725;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtUsername 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   2175
         VariousPropertyBits=   746604571
         Size            =   "3836;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblUsername 
         Height          =   195
         Left            =   255
         TabIndex        =   1
         Top             =   345
         Width           =   1125
         VariousPropertyBits=   276824091
         Caption         =   "Username"
         Size            =   "1984;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.Label Label1 
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1380
      Width           =   705
      VariousPropertyBits=   276824091
      Caption         =   "Password"
      Size            =   "1244;344"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   315
      Left            =   1650
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
      VariousPropertyBits=   746604571
      Size            =   "3836;556"
      PasswordChar    =   42
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   90
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
   Begin VB.Image imgCaption 
      Height          =   315
      Left            =   30
      Top             =   30
      Width           =   3915
   End
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   2040
      Width           =   1305
      Caption         =   "Exit"
      Size            =   "2302;635"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   360
      Left            =   1695
      TabIndex        =   6
      Top             =   2040
      Width           =   1305
      Caption         =   "Login"
      Size            =   "2302;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmChangepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : FIS - CMC Software Solution
' Project           : Du an ho tro ke khai thue
' Package           : Interface
' Form, Module
'   or Class name   : frmTreeviewMenu
' Descriptions      : Report sh
' Start date        : 21/11/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : TuanLM
' Integrate         :
' Project manager   : ThietKN
' Last modify       :
' Reason of modify  :
'******************************************************
Option Explicit

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
            
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
            
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Sub cmdClose_Click()
    Unload Me
    'Unload frmSystem
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorHandle
   
'    If Len(txtUsername.Text) = 0 Then
'        DisplayMessage "0056", msOKOnly, miInformation
'        txtUsername.SetFocus
'        Exit Sub
'    End If
    
    'Quangtv
    
    Dim fs As FileSystemObject
     Set fs = New FileSystemObject
     
    Dim txt As String
       
    If txtPassword.Text <> strFile(2) Then
        DisplayMessage "0091", msOKOnly, miInformation
    Else
        txt = spathVat
        txt = txt & "," & txtUsername.Text
        txt = txt & "," & txtNew.Text
        Call WritePathFile(App.path & "\config.txt", txt)
        
        DisplayMessage "0090", msOKOnly, miInformation
        
    End If
    Unload Me
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
End Sub


Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_Load()
    SetControlCaption Me, "frmChangePass"
    Call ReadPathFile(App.path & "\config.txt")
    txtUsername.Text = strFile(1)
    txtUsername.Enabled = False
End Sub
Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set frmChangepass = Nothing
End Sub

Private Sub txtUsername_Change()
    txtUsername.Text = UCase(txtUsername.Text)
End Sub

Private Sub txtUsername_LostFocus()
    If Len(txtUsername.Text) > 0 Then
        txtUsername.Text = UCase(txtUsername.Text)
    End If
End Sub


