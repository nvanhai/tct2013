VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDuongDan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Më tÖp b¶ng kª"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
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
   ScaleHeight     =   2985
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTaiDuLieu 
      Caption         =   "&T¶i d÷ liÖu"
      Height          =   420
      Left            =   2205
      TabIndex        =   7
      Top             =   2430
      Width           =   1365
   End
   Begin VB.CommandButton cmdThoat 
      Caption         =   "&§ãng"
      Height          =   420
      Left            =   3735
      TabIndex        =   6
      Top             =   2430
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4275
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5190
      Begin VB.CommandButton cmdBrow 
         Caption         =   "..."
         Height          =   420
         Left            =   4680
         TabIndex        =   5
         Top             =   1575
         Width           =   420
      End
      Begin VB.TextBox txtDuongDan 
         Height          =   420
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1575
         Width           =   4425
      End
      Begin VB.OptionButton optThemXoa 
         Caption         =   "Thªm d÷ liÖu míi, xãa d÷ liÖu ®· cã"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   765
         Width           =   3390
      End
      Begin VB.OptionButton optThem 
         Caption         =   "Thªm d÷ liÖu vµo d÷ liÖu ®· cã"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3390
      End
      Begin VB.Label Label1 
         Caption         =   "Chän ®­êng dÉn tíi file ®Ó nhËn d÷ liÖu"
         Height          =   375
         Left            =   225
         TabIndex        =   3
         Top             =   1260
         Width           =   3210
      End
   End
End
Attribute VB_Name = "frmDuongDan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strFileName As String

Private Sub cmdBrow_Click()
    With CommonDialog1
        .CancelError = False
        .Filter = "File (UNICODE Font) (*.xls)|*.xls|File (TCVN3 Font) (*.xls)|*.xls|File (VNI Font) (*.xls)|*.xls|File (VIQR Font) (*.xls)|*.xls|File (VISCII Font) (*.xls)|*.xls"
        .FilterIndex = 1
        .DialogTitle = "Chon bang ke de import vao chuong trinh"
        .ShowOpen
        txtDuongDan.Text = .FileName
        Select Case .FilterIndex
            Case 1
                 strfileFont = "UNICODE"
            Case 2
                 strfileFont = "TCVN"
            Case 3
                 strfileFont = "VNI"
            Case 4
                 strfileFont = "VIQR"
            Case 5
                 strfileFont = "VISCII"
        End Select
    End With
End Sub

Private Sub cmdTaiDuLieu_Click()
    If Trim(txtDuongDan.Text) = vbNullString Or Trim(txtDuongDan.Text) = "" Then
        DisplayMessage "0148", msOKOnly, miInformation, "Tai du lieu"
        cmdBrow.SetFocus
        Exit Sub
    Else
        If optThem.value Then
            themDuLieu = True
        Else
            themDuLieu = False
        End If
        If optThemXoa.value Then
            themXoaDuLieu = True
        Else
            themXoaDuLieu = False
        End If
        strFileName = txtDuongDan.Text
    End If
    Unload Me
End Sub

Public Function getFileName() As String
    Me.Show vbModal
    getFileName = RTrim(LTrim(strFileName))
End Function

Private Sub cmdThoat_Click()
    Unload Me
End Sub

