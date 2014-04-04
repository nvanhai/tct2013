VERSION 5.00
Begin VB.Form frmBarcodeScan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Barcode Scan"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
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
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBarcode 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   5775
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   4935
   End
End
Attribute VB_Name = "frmBarcodeScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    frmBarcodeScan.Top = (frmSystem.Height - frmLogin.Height) / 2
    frmBarcodeScan.Left = (frmSystem.Width - frmLogin.Width) / 2
End Sub
