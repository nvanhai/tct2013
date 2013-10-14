VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin MSForms.CommandButton cmdExit 
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
      Size            =   "3413;714"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCaption 
      Height          =   315
      Left            =   570
      TabIndex        =   0
      Top             =   150
      Width           =   1995
      ForeColor       =   -2147483634
      Size            =   "3519;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   435
      Left            =   90
      Top             =   30
      Width           =   4755
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = frmSystem.Top + (frmSystem.Height - Me.Height) / 2
    Me.Left = frmSystem.Left + (frmSystem.Width - Me.Width) / 2
    SetControlCaption Me, "frmAbout"
End Sub

Private Sub Form_Resize()
     SetFormCaption Me, imgCaption, lblCaption
End Sub
