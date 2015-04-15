VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "NhËn file d÷ liÖu font TCVN3 (ABC)"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "NhËn file d÷ liÖu font Unicode"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   3615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public fontSelect As String


Private Sub OKButton_Click()
    If Option1.value = True Then
        fontSelect = "Unicode"
    Else
        fontSelect = "ABC"
    End If
End Sub
