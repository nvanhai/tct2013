VERSION 5.00
Begin VB.Form frmInBienBanPhatNopCham 
   Appearance      =   0  'Flat
   Caption         =   "In biªn b¶n vi ph¹m ph¸p luËt vÒ thuÕ"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInBienBanPhatNopCham.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "§¹i diÖn ng­êi nép thuÕ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   34
      Top             =   1440
      Width           =   10095
      Begin VB.TextBox txtDaiDien1 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtDaiDien2 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtChucVuDaiDien1 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   5
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtChucVuDaiDien2 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label19 
         Caption         =   "1. Hä vµ tªn:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Chøc vô:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "2. Hä vµ tªn:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Chøc vô:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   35
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   29
      Top             =   7200
      Width           =   9615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   4695
         TabIndex        =   40
         Top             =   0
         Width           =   4695
         Begin VB.CommandButton cmdIn 
            Caption         =   "&In"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   42
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton cmdThoat 
            Caption         =   "Th&o¸t"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   41
            Top             =   120
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Th«ng tin ph¹t nép chËm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   10095
      Begin VB.TextBox txtCanCuMucPhat 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmInBienBanPhatNopCham.frx":0ECA
         Top             =   2760
         Width           =   9735
      End
      Begin VB.TextBox txtKyKeKhai 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Ky ke khai"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtLoaiHoSo 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Loai ho so"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtCanCuViPham 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmInBienBanPhatNopCham.frx":0ED3
         Top             =   1800
         Width           =   9735
      End
      Begin VB.TextBox txtDiaChi 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Dia chi"
         Top             =   720
         Width           =   8655
      End
      Begin VB.TextBox txtMaTNT 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Ma so thue"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtNNT 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Don vi"
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtNgayNopHoSo 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Ngay nop ho so"
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox txtHanNopHoSo 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "han nop ho so"
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "C¨n cø møc ph¹t:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lbThoiGianChamNop 
         Caption         =   "??? Ngµy lµm viÖc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label14 
         Caption         =   "Thêi gian chËm nép:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Kú kª khai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   31
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Lo¹i hå s¬:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "C¨n cø vi ph¹m:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "§Þa chØ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "M· sè thuÕ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "§¬n vÞ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Ngµy nép hå s¬:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   3630
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "H¹n nép hå s¬:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3630
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "§¹i diÖn bé phËn mét cöa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtChucVuCanBo2 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtChucVuCanBo1 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtTenCanBo2 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtTenCanBo1 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Chøc vô:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "2. ¤ng/bµ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Chøc vô:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "1. ¤ng/bµ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmInBienBanPhatNopCham"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public IDinBB As String
Dim TEN_CQT As String
Dim TEN_CUCTHUE  As String


Private Sub CmdIn_Click()
'Ten chuc vu can bo
'rptBienBanViPhamPhapLuatThue.txtTenDaiDien1.Text = Me.txtTenCanBo1.Text
'rptBienBanViPhamPhapLuatThue.txtTenDaiDien2.Text = Me.txtTenCanBo2.Text
'rptBienBanViPhamPhapLuatThue.txtChucVuDaiDien1.Text = Me.txtChucVuCanBo1.Text
'rptBienBanViPhamPhapLuatThue.txtChucVuDaiDien2.Text = Me.txtChucVuCanBo2.Text
'
''Ten chuc vu dai dien NNT
'rptBienBanViPhamPhapLuatThue.txtTenDaiDienNNT1.Text = Me.txtDaiDien1.Text
'rptBienBanViPhamPhapLuatThue.txtTenDaiDienNNT2.Text = Me.txtDaiDien2.Text
'rptBienBanViPhamPhapLuatThue.txtChucVuDaiDienNNT1.Text = Me.txtChucVuDaiDien1.Text
'rptBienBanViPhamPhapLuatThue.txtChucVuDaiDienNNT2.Text = Me.txtChucVuDaiDien2.Text
'
'
''rptBienBanViPhamPhapLuatThue.txtDaiDienDTNT.Text = Me.txtNguoiNop.Text
'rptBienBanViPhamPhapLuatThue.txtTenDonVi.Text = Me.txtNNT.Text
'rptBienBanViPhamPhapLuatThue.txtMaDTNT.Text = Me.txtMaTNT.Text
'rptBienBanViPhamPhapLuatThue.txtDiaChi.Text = Me.txtDiaChi.Text
''rptBienBanViPhamPhapLuatThue.txtCanCu.Text = Me.txtCanCu.Text
''rptBienBanViPhamPhapLuatThue.txtPhuTrach.Text = Me.txtTenCanBo1.Text
''rptBienBanViPhamPhapLuatThue.txtNguoiLapBB.Text = Me.txtTenCanBo2.Text
''Dia diem co quan Thue
'rptBienBanViPhamPhapLuatThue.lblTenDiaDiem.caption = LayThamSo("TEN_DIADIEM")
''Can cu vi pham
'rptBienBanViPhamPhapLuatThue.txtCanCuViPham.Text = Me.txtCanCuViPham.Text
'
''Can cu muc phat
'rptBienBanViPhamPhapLuatThue.txtDieuKhoan.Text = Me.txtCanCuMucPhat.Text
'
''Ngay, so ngay cham
'Dim s As String
's = "Ngµy " & Me.txtNgayNopHoSo.Text & " Chi côc ThuÕ míi nhËn ®­îc hå s¬ khai thuÕ " & Me.txtLoaiHoSo.Text
's = s & ", kú kª khai " & Me.txtKyKeKhai.Text & " cña c¬ së kinh doanh. C¬ së ®· vi ph¹m vÒ chËm nép hå s¬ thuÕ: " & lbThoiGianChamNop.caption & "."
'
''Thêi gian nép chËm lµ: " & Me.txtThoiGianChamNop.Text & " ngµy lµm viÖc."
'
'rptBienBanViPhamPhapLuatThue.txtChiTiet.Text = s
''ten cuc thue, chi cuc thue
'rptBienBanViPhamPhapLuatThue.lblCucThue.caption = LayThamSo("Ten_CucThue")    ' "ten Cuc Thue tu DB"
'rptBienBanViPhamPhapLuatThue.lblChiCucThue.caption = LayThamSo("Ten_CQT")
'CaiDatThamSo rptBienBanViPhamPhapLuatThue, "Doc", 900, 50, 700, 700

'rptBienBanViPhamPhapLuatThue.Show vbModal
On Error GoTo ErrorHandle
    Dim tencanbo1 As String
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lRowCtrl As Long, lColCtrl As Long
    'Dim xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlTempCellNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lngCol As Long, lngRow As Long, lngColRol As String
    
    Dim arrT(15) As String
    Dim i As Integer
    
    SetupDataPrint
    Prepare_In
    arrT(0) = TAX_Utilities_Svr_New.Convert(TEN_CUCTHUE, TCVN, UNICODE)
    arrT(1) = TAX_Utilities_Svr_New.Convert(TEN_CQT, TCVN, UNICODE)
    arrT(2) = TAX_Utilities_Svr_New.Convert(Me.txtTenCanBo1.Text, TCVN, UNICODE)
    arrT(3) = TAX_Utilities_Svr_New.Convert(Me.txtChucVuCanBo1.Text, TCVN, UNICODE)
    arrT(4) = TAX_Utilities_Svr_New.Convert(Me.txtTenCanBo2.Text, TCVN, UNICODE)
    arrT(5) = TAX_Utilities_Svr_New.Convert(Me.txtChucVuCanBo2.Text, TCVN, UNICODE)
    arrT(6) = TAX_Utilities_Svr_New.Convert(Me.txtDaiDien1.Text, TCVN, UNICODE)
    arrT(7) = TAX_Utilities_Svr_New.Convert(Me.txtChucVuDaiDien1.Text, TCVN, UNICODE)
    arrT(8) = TAX_Utilities_Svr_New.Convert(Me.txtDaiDien2.Text, TCVN, UNICODE)
    arrT(9) = TAX_Utilities_Svr_New.Convert(Me.txtChucVuDaiDien2.Text, TCVN, UNICODE)
    arrT(10) = TAX_Utilities_Svr_New.Convert(Me.txtNNT.Text, TCVN, UNICODE)
    'arrT(11) = TAX_Utilities_Svr_New.Convert(Me.txtDiaChi.Text, TCVN, UNICODE)
    arrT(11) = TAX_Utilities_Svr_New.Convert(Me.txtDiaChi.Text, TCVN, UNICODE)
    arrT(12) = TAX_Utilities_Svr_New.Convert(Me.txtMaTNT.Text, TCVN, UNICODE)
    arrT(13) = TAX_Utilities_Svr_New.Convert(Me.txtCanCuViPham.Text, TCVN, UNICODE)
    arrT(14) = TAX_Utilities_Svr_New.Convert(Me.txtCanCuMucPhat.Text, TCVN, UNICODE)
    i = 0
   Set xmlNodeList = TAX_Utilities_Svr_New.Data(0).getElementsByTagName("Cell")
    
    'Set xmlTempCellNode = xmlNodeList.Item
    For Each xmlTempCellNode In xmlNodeList
         lngColRol = GetAttribute(xmlTempCellNode, "CellID")
            If arrT(i) <> vbNullString Then
                UpdateCellPrint lngColRol, arrT(i)
            End If
        i = i + 1
    Next
    frmReports.Show 1
    
ErrorHandle:
    'Restore active properties of node validit

End Sub

Private Sub cmdThoat_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    InitParametersPrint
End Sub

Private Sub txtChucVuCanBo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChucVuCanBo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChucVuDaiDien1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChucVuDaiDien2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDaiDien1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDaiDien2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHanNopHoSo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Public Sub txtHanNopHoSo_LostFocus()
''    If txtHanNopHoSo.Text = "" Then Exit Sub
''    txtHanNopHoSo.Text = XuLyNgay(txtHanNopHoSo.Text)
''    If Not isDateChuan(txtHanNopHoSo.Text) Then
''        txtHanNopHoSo.Text = ""
''        MsgBox "nhËp l¹i h¹n nép hå s¬", vbInformation, "Th«ng b¸o"
''        Exit Sub
''    End If
'
'    Dim D1 As Date
'    Dim d2 As Date
'    Dim strSQL As String
'
'    D1 = DateSerial(Int(Mid$(txtHanNopHoSo.Text, 7, 4)), Int(Mid$(txtHanNopHoSo.Text, 4, 2)), Int(Mid$(txtHanNopHoSo.Text, 1, 2))) + 1
'    d2 = DateSerial(Int(Mid$(txtNgayNopHoSo.Text, 7, 4)), Int(Mid$(txtNgayNopHoSo.Text, 4, 2)), Int(Mid$(txtNgayNopHoSo.Text, 1, 2)))
'
'   If clsDAO.Connected_qhs = False Then
'            clsDAO.Connect_qhs
'    End If
'
'        i = 0
'        j = 0
'
'        While (D1 <= d2)
'            Set rs = New Recordset
'             strSQL = "Select Count(*) from QHS_DM_NGAYNGHI  where NgayNghi = '" & format(D1, "MM/DD/YYYY") & "'"
'             Set rs = clsDAO.Execute_Qhs(strSQL)
''            rs.Open "Select Count(*) from QHS_DM_NGAYNGHI  where NgayNghi = '" & format(D1, "MM/DD/YYYY") & "'", cn
'            If ((Weekday(D1) = 1) Or (Weekday(D1) = 7)) And (rs.Fields.Item(0) <> 0) Then j = j + 1
'            If j < 0 Then j = 0
'            If Not ((Weekday(D1) = 1) Or (Weekday(D1) = 7) Or (rs.Fields.Item(0) <> 0) Or (j > 0)) Then i = i + 1
'            If (Weekday(D1) <> 1) And (Weekday(D1) <> 7) And (rs.Fields.Item(0) = 0) Then j = j - 1
'
'            D1 = D1 + 1
'        Wend
'
'    clsDAO.DisConnect_qhs
'
'    lbThoiGianChamNop.caption = str(i) + " (ngµy lµm viÖc)"
'End Sub
Public Sub txtHanNopHoSo_LostFocus()


    Dim strSQL As String
    If txtHanNopHoSo.Text = "" Then Exit Sub

    'txtHanNopHoSo.Text = DateSerial(Int(Mid$(txtHanNopHoSo.Text, 7, 4)), Int(Mid$(txtHanNopHoSo.Text, 4, 2)), Int(Mid$(txtHanNopHoSo.Text, 1, 2))) + 1

'    If Not isDateChuan(txtHanNopHoSo.Text) Then
'
'        txtHanNopHoSo.Text = ""
'
'        MsgBox "nhËp l¹i h¹n nép hå s¬", vbInformation, "Th«ng b¸o"
'
'        Exit Sub
'
'    End If

   

    Dim D1 As Date

    Dim d2 As Date

   
    D1 = DateSerial(Int(Mid$(txtHanNopHoSo.Text, 7, 4)), Int(Mid$(txtHanNopHoSo.Text, 4, 2)), Int(Mid$(txtHanNopHoSo.Text, 1, 2)))
    d2 = DateSerial(Int(Mid$(txtNgayNopHoSo.Text, 7, 4)), Int(Mid$(txtNgayNopHoSo.Text, 4, 2)), Int(Mid$(txtNgayNopHoSo.Text, 1, 2)))
'    D1 = StringToDate(txtHanNopHoSo.Text)
'
'    d2 = StringToDate(txtNgayNopHoSo.Text)

    i = 0

    If clsDAO.Connected_qhs = False Then
            clsDAO.Connect_qhs
    End If
        Set rs = New Recordset
        strSQL = "Select Count(*) as TONG from QHS_DM_NGAYNGHI  where NgayNghi >= '" & format(D1 + 1, "MM/DD/YYYY") & "' and NgayNghi <= '" & format(d2, "MM/DD/YYYY") & "'"
        Set rs = clsDAO.Execute_Qhs(strSQL)
        '    rs.Open "Select Count(*) as TONG from QHS_DM_NGAYNGHI  where NgayNghi >= '" & format(D1 + 1, "MM/DD/YYYY") & "' and NgayNghi <= '" & format(d2, "MM/DD/YYYY") & "'", cn
        i = DateDiff("D", D1, d2) - rs!TONG
    clsDAO.Disconnect
    lbThoiGianChamNop.caption = str(i) + " (ngµy lµm viÖc)"

End Sub

Private Sub txtTenCanBo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTenCanBo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Public Function InitParametersPrint() As Boolean

'ThanhDX modified
'Date: 10/04/06
    Dim strTaxID As String, strID As String
    Dim s As String
    
On Error GoTo ErrHandle
    
    strID = 88
    SetNodeMenu strID
    'SetPeriod Right$(strTaxReportInfo, 6)
    TAX_Utilities_Svr_New.NodeValidity = GetValidityNode
    Me.txtHanNopHoSo = format(frmInterfaces.HAN_NOP, "dd/mm/yyyy")
    Me.txtNgayNopHoSo = format(frmInterfaces.NGAYNOP_PRINT, "dd/mm/yyyy")
    txtHanNopHoSo_LostFocus
    s = ". Thêi gian nép chËm lµ: " & Me.lbThoiGianChamNop.caption & "."
    Me.txtNNT.Text = frmInterfaces.NNT_PRINT
    Me.txtDiaChi.Text = frmInterfaces.DIACHI_PRINT
    Me.txtMaTNT.Text = frmInterfaces.MST_PRINT
    Me.txtKyKeKhai.Text = frmInterfaces.KyKeKhai
    Me.txtCanCuViPham = TAX_Utilities_Svr_New.Convert(frmInterfaces.CAN_CU1, UNICODE, TCVN)
    Me.txtCanCuMucPhat = TAX_Utilities_Svr_New.Convert(frmInterfaces.CAN_CU2, UNICODE, TCVN) & s
    Me.txtLoaiHoSo = TAX_Utilities_Svr_New.Convert(frmInterfaces.LOAihs_PRINT, UNICODE, TCVN)
    
'     Dim NGAYNOP As String
'        fpSpread1.GetText fpSpread1.ColLetterToNumber("E"), 12, NGAYNOP
'        'NGNOP = Date
'
'        If Trim(NGAYNOP) = vbNullString Then
'            NGAYNOP = "CTOD('')"
'        Else
'            'NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGAYNOP = DateSerial(Int(Mid$(NGAYNOP, 7, 4)), Int(Mid$(NGAYNOP, 4, 2)), Int(Mid$(NGAYNOP, 1, 2)))
'            NGAYNOP_PRINT = "CTOD('" & format(NGAYNOP, "mm/dd/yyyy") & "')"
'        End If
    
    '*******************************
    'RestoreDataFile (strData)
'    If Not RestoreDataFile(strData) Then  ', rsTaxInfor
'        If blnReceiveByBarcode Then
'            MessageBox "0057", msOKOnly, miCriticalError
'        Else
'            MessageBox "0053", msOKOnly, miCriticalError
'        End If
'        Exit Function
'    End If
    
    InitParametersPrint = True

    Exit Function
ThamSoErrHandle:
    DisplayMessage "0078", msOKOnly, miCriticalError
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "InitParameters", Err.Number, Err.Description
End Function
Private Sub SetNodeMenu(strMenuId As String)
    Dim xmlMenuDom As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
On Error GoTo ErrHandle
    xmlMenuDom.Load App.path & "\Menu.xml"
    
    For Each xmlNode In xmlMenuDom.getElementsByTagName("Root")(0).childNodes
        If GetAttribute(xmlNode, "ID") = strMenuId Then
            TAX_Utilities_Svr_New.NodeMenu = xmlNode
            Exit For
        End If
    Next
    IDinBB = 88
    Set xmlNode = Nothing
    Set xmlMenuDom = Nothing
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetNodeMenu", Err.Number, Err.Description
End Sub
Private Function UpdateCellPrint(ByVal pColRow As String, ByVal pValue As String) As Boolean
    On Error GoTo ErrHandle
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    'GetCellSpan fpSpread1, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_Svr_New.Data(0).nodeFromID(pColRow)
    
    
    If GetAttribute(xmlNodeCell, "Value") <> pValue Then
        SetAttribute xmlNodeCell, "Value", pValue
        UpdateCell = True
    End If
    
    Set xmlNodeCell = Nothing
    
    Exit Function
    
ErrHandle:
    SaveErrorLog Me.Name, "UpdateCell", Err.Number, Err.Description
End Function
'dhdang
'ngay 01/09/2010
Private Function Prepare_In() As String
    Dim strSQL As String
    Dim strSQL1 As String
   Dim rs As ADODB.Recordset
   Dim NGNOP As Variant
   
    'sSQLCol = "DHS_MA, SO_HOSO_NHAN, TIN,TEN,DIA_CHI, NGAY_NHAN,NGUOI_NOP,NGAY_NHAP,NGUOI_NHAP,HAN_XULY,PHONG_XULY,PHONG_XULY_HIENTAI,GHI_CHU,TTHAI_HOSO,HTHUC_NOP,GUI_BD"
    

'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            'NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = DateSerial(Int(Mid$(NGNOP, 7, 4)), Int(Mid$(NGNOP, 4, 2)), Int(Mid$(NGNOP, 1, 2)))
'            'NGAYNOP_PRINT = "CTOD('" & format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
        
        
        strSQL = "Select top 1 GIA_TRI from QHSCC.dbo.QHS_THAMSO_HETHONG where TEN = 'TEN_CQT'"
        strSQL1 = "Select top 1 GIA_TRI from QHSCC.dbo.QHS_THAMSO_HETHONG where TEN = 'TEN_CUCTHUE'"
        
        'clsDAO.DisConnect_qhs
        If clsDAO.Connected_qhs = False Then
            clsDAO.Connect_qhs
        End If
        
        Set rs = clsDAO.Execute_Qhs(strSQL)
        Set rs1 = clsDAO.Execute_Qhs(strSQL1)
        If rs Is Nothing Then
            TEN_CQT = "Chi Cuc thue"
        Else
            TEN_CQT = rs(0)
        End If
        
        If rs Is Nothing Then
            TEN_CUCTHUE = "Cuc thue"
        Else
            TEN_CUCTHUE = rs1(0)
        End If
        clsDAO.DisConnect_qhs
   
End Function

