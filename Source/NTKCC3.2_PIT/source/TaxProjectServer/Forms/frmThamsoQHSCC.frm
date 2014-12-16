VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmThamsoQHSCC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   75
      TabIndex        =   0
      Top             =   390
      Width           =   4980
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   315
         Left            =   4410
         TabIndex        =   2
         Top             =   555
         Width           =   375
      End
      Begin MSForms.Label lblDir 
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   915
         VariousPropertyBits=   276824091
         Caption         =   "Path QHSCC"
         Size            =   "1614;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDir 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   4245
         VariousPropertyBits=   746604571
         Size            =   "7488;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   360
      Left            =   2280
      TabIndex        =   6
      Top             =   1455
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
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3750
      TabIndex        =   5
      Top             =   1455
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
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   360
      TabIndex        =   1
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
End
Attribute VB_Name = "frmThamsoQHSCC"
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
  lparam As Long
  iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
            
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
            
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorHandle

   
'    If Len(txtUsername.Text) = 0 Then
'        DisplayMessage "0056", msOKOnly, miInformation
'        txtUsername.SetFocus
'        Exit Sub
'    End If
    'Quangtv
    'Beep
    
    If Len(txtDir.Text) = 0 Then
        DisplayMessage "0118", msOKOnly, miInformation
        txtDir.SetFocus
        Exit Sub
    End If
    Dim fs As FileSystemObject
     Set fs = New FileSystemObject
    If Not fs.FolderExists(txtDir.Text) Then
        DisplayMessage "0084", msOKOnly, miInformation
        txtDir.SetFocus
        Exit Sub
    End If

            '********************************
            If txtDir.Text <> spathQHSCC Then
                Dim txt As String
                txt = strFile(0) & "," & strFile(1) & "," & strFile(2) & "," & txtDir.Text
                Call WritePathFile(App.path & "\config.txt", txt)
                DisplayMessage "0090", msOKOnly, miInformation
                spathQHSCC = txtDir.Text
            End If
                      
    Unload Me
    Set frmThamsoQHSCC = Nothing
    'dhdang sua canh bao ko ke noi duoc den QHS
    'ngay 07\07\2010
    If CheckConnection = False Then
        DisplayMessage "0117", msOKOnly, miInformation
    End If
    Exit Sub
    
   ' frmSystem.Show
    
ErrorHandle:
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
End Sub

Private Sub cmdOpen_Click()
    Dim path As String
    path = BrowseFolder("")
    txtDir.Text = path
End Sub
Public Function BrowseFolder(szDialogTitle As String) As String
  Dim X As Long, bi As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer
  
    With bi
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = vbNullString
    End If
End Function
Private Sub Form_Activate()
    'txtDir.SetFocus
End Sub

Private Sub Form_Load()
    SetControlCaption Me, "frmThamsoQHSCC"
    Call ReadPathFile(App.path & "\config.txt")
    txtDir.Text = strFile(3)
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Set clsDAO = Nothing
    Set frmThamsoQHSCC = Nothing
End Sub

'dhdang tao ham check connection toi QHS
'06/07/2010

Private Function CheckConnection() As Boolean
    Dim flag As Boolean

    clsDAO.CreateConnectionStringCheckSQL spathQHSCC
    clsDAO.Connect_qhs
    flag = clsDAO.Connected_qhs
    clsDAO.DisConnect_qhs
    CheckConnection = flag
'CheckConnection = True
End Function

