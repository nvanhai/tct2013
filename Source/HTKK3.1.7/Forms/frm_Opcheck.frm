VERSION 5.00
Begin VB.Form frm_Opcheck 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Thông tin in"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "DS Sans Serif"
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
   Moveable        =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "§ãng"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "§ång ý"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt_end 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Txt_star 
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Op_numberch 
      Caption         =   "Chän theo sè thø tù"
      Height          =   250
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.OptionButton Op_noncheck 
      Caption         =   "Bá chän tÊt c¶"
      Height          =   250
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.OptionButton Op_allcheck 
      Caption         =   "Chän tÊt c¶"
      Height          =   250
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "§Õn"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Tõ"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "frm_Opcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dhdang
'form chon thong tin in
Private options As Integer
Private star As String
Private endd As String


Public Function getOptions() As String
    Me.Show vbModal
    getOptions = options
   End Function
Public Function getStar() As String
    'Me.Show vbModal
       getStar = star
End Function
Public Function getEndd() As String
       getEndd = endd
End Function

Private Sub Command1_Click()
        If Op_numberch.value = True Then
                options = 3
                star = Trim(Txt_star.Text)
                endd = Trim(Txt_end.Text)
            Else
                If Op_noncheck.value = True Then
                    options = 2
                Else
                    options = 1
                End If
            End If
        Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Op_numberch_Click()
                Txt_star.Enabled = True
                Txt_end.Enabled = True
                Txt_star.SetFocus
End Sub
'bat loi nhap chu so
Private Sub Txt_end_KeyPress(KeyAscii As Integer)
        If KeyAscii >= 48 And KeyAscii <= 57 Then
        'Your code here
        Else
        KeyAscii = 8
        Beep
        End If
End Sub

Private Sub Txt_star_KeyPress(KeyAscii As Integer)
        If KeyAscii >= 48 And KeyAscii <= 57 Then
        'Your code here
        Else
        KeyAscii = 8
        Beep
End If
End Sub
