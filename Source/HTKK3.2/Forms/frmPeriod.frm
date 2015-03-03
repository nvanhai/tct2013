VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmPeriod 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7230
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11910
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox chkTKKy 
      Caption         =   "Tê khai kú"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   44
      Top             =   11280
      Width           =   1335
   End
   Begin VB.CheckBox chkTKhaiLanXB 
      Caption         =   "TK lÇn xuÊt b¸n"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   43
      Top             =   11640
      Width           =   1575
   End
   Begin VB.TextBox txtLanXuat 
      Height          =   315
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   42
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkKhiThien 
      Caption         =   "KhÝ thiªn nhiªn"
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
      Left            =   3360
      TabIndex        =   39
      Top             =   11640
      Width           =   1575
   End
   Begin VB.CheckBox chkCondensate 
      Caption         =   "Condensate"
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
      TabIndex        =   38
      Top             =   11640
      Width           =   1335
   End
   Begin VB.CheckBox chkDauTho 
      Caption         =   "DÇu th«"
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
      TabIndex        =   37
      Top             =   11640
      Width           =   1095
   End
   Begin VB.CheckBox chkQTNamDau 
      Caption         =   "QuyÕt to¸n hÕt vµo n¨m ®Çu"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   34
      Top             =   10920
      Width           =   2415
   End
   Begin VB.CheckBox chkQTTungNam 
      Caption         =   "QuyÕt to¸n cho riªng tõng n¨m"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   10920
      Width           =   2655
   End
   Begin VB.CheckBox chkTuThangDenThang 
      Caption         =   "Tê khai tõ th¸ng ®Õn th¸ng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CheckBox chkTKQuy 
      Caption         =   "Tê khai quý"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CheckBox chkTKLanPS 
      Caption         =   "Tê khai lÇn ph¸t sinh"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CheckBox chkTkhaiThang 
      Caption         =   "Tê khai th¸ng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   29
      Top             =   9720
      Width           =   1335
   End
   Begin VB.OptionButton OptTKLanPS 
      Caption         =   "Tê khai lÇn ph¸t sinh"
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
      Left            =   3240
      TabIndex        =   27
      Top             =   9120
      Width           =   1815
   End
   Begin VB.OptionButton OptTKThang 
      Caption         =   "Tê khai th¸ng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   9120
      Width           =   2295
   End
   Begin VB.ComboBox cboNganhKD 
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   7920
      Width           =   4575
   End
   Begin VB.OptionButton OptChinhthuc 
      Caption         =   "Tê khai lÇn ®Çu"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   6240
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.OptionButton OptBosung 
      Caption         =   "Tê khai bæ sung"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   21
      Top             =   6600
      Width           =   1785
   End
   Begin VB.TextBox txtSolan 
      Height          =   315
      Left            =   4530
      MaxLength       =   2
      TabIndex        =   20
      Text            =   "1"
      Top             =   6930
      Width           =   645
   End
   Begin VB.TextBox txtDay 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   840
      MaxLength       =   2
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ComboBox cmbQuy 
      Height          =   315
      ItemData        =   "frmPeriod.frx":0000
      Left            =   1860
      List            =   "frmPeriod.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtMonth 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   0
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox txtNgayCuoi 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3630
      MaxLength       =   10
      TabIndex        =   4
      Top             =   5190
      Width           =   1095
   End
   Begin VB.TextBox txtNgayDau 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   5220
      Width           =   1095
   End
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
      Left            =   2520
      TabIndex        =   9
      Top             =   2760
      Width           =   1275
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
      Left            =   1170
      TabIndex        =   8
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Frame frmKy 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   30
      TabIndex        =   11
      Top             =   330
      Width           =   4755
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   30
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   4755
      Begin VB.CheckBox chkSelectAll 
         Height          =   195
         HelpContextID   =   81211
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Width           =   195
      End
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   825
         HelpContextID   =   81211
         Left            =   60
         TabIndex        =   7
         Top             =   480
         Width           =   4635
         _Version        =   458752
         _ExtentX        =   8176
         _ExtentY        =   1455
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   1
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmPeriod.frx":0020
      End
      Begin VB.Label lblSelectAll 
         Caption         =   "Chän phô lôc kª khai"
         BeginProperty Font 
            Name            =   "DS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   6
         Top             =   210
         Width           =   2355
      End
   End
   Begin FPUSpreadADO.fpSpread fpsNgaykhaiBS 
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   8400
      Visible         =   0   'False
      Width           =   3015
      _Version        =   458752
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   4
      MaxRows         =   3
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "frmPeriod.frx":031A
      UserResize      =   1
      Appearance      =   1
   End
   Begin VB.Label lblLanXuat 
      AutoSize        =   -1  'True
      Caption         =   "LÇn xuÊt b¸n"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   41
      Top             =   6840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblLanXuatBan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   360
      TabIndex        =   40
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label lblDenThang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "®Õn th¸ng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   36
      Top             =   11400
      Width           =   735
   End
   Begin VB.Label lblTuThang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tõ th¸ng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   35
      Top             =   11400
      Width           =   660
   End
   Begin VB.Label lblNganhKD 
      Caption         =   "Danh môc ngµnh nghÒ"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   7440
      Width           =   4335
   End
   Begin VB.Label lblSolan 
      Caption         =   "LÇn"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   6960
      Width           =   405
   End
   Begin VB.Label lblNgay 
      Caption         =   "Ngµy"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5760
      Width           =   525
   End
   Begin VB.Label lblYear 
      Caption         =   "N¨m"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label lblQuy 
      Caption         =   "Quý"
      Height          =   255
      Left            =   1470
      TabIndex        =   16
      Top             =   4830
      Width           =   375
   End
   Begin VB.Label lblMonth 
      Caption         =   "Th¸ng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   5760
      Width           =   525
   End
   Begin VB.Label lblNgayDau 
      Caption         =   "Tõ ngµy"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   14
      Top             =   5220
      Width           =   1095
   End
   Begin VB.Label lblNgayCuoi 
      Caption         =   "§Õn ngµy"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2550
      TabIndex        =   13
      Top             =   5220
      Width           =   1095
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Chän kú tÝnh thuÕ"
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
      Left            =   360
      TabIndex        =   12
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   120
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmPeriod
' Descriptions      : Report sh
' Start date        : 10/08/2007 (dd/mm/yyyy)
' Finish date       :
' Coder             : htphuong
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************

Option Explicit

Dim sFormat As String
'Dim intMonth As Integer ' 1 - Month, 3- ThreeMonth
Dim bIsClosed As Boolean
Dim blnClick As Boolean
Dim blnValidInfo(1 To 4) As Boolean
Dim oldYear As String
Dim yChange As String
Dim oldMonth As String
Private blnFPChange As Boolean
Dim objCvt As DateUtils
Dim strDateKHBS As String

Private strLoaiSacThue As String

Private arrStrXMLFileNames() As String


Private Sub cboNganhKD_Click()
    strLoaiNNKD = cboNganhKD.ItemData(cboNganhKD.ListIndex)
    ' xu lý ten data file cho cac to khai DK
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89" Then
        If strLoaiNNKD = 1 Then
            strLoaiTkDk = "DT"
        ElseIf strLoaiNNKD = 2 Then
            strLoaiTkDk = "KTN"
        ElseIf strLoaiNNKD = 3 Then
            strLoaiTkDk = "CD"
        End If
        LoadGrid
    End If
End Sub

Private Sub chkCondensate_Click()
Dim d, m, Y As Integer
    If chkCondensate.value = 1 Then
        chkDauTho.value = 0
        chkKhiThien.value = 0
        
        strCondensate = chkCondensate.value
        strLoaiTkDk = "CD"
        
        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
            strLoaiTKThang_PS = "TK_LANPS"
        ElseIf (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93") Then
            strLoaiTKThang_PS = "TK_NAM"
        End If
        strDauTho = 0
        strKhiThienNhien = 0

        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
            lblNgay.Visible = True
            txtDay.Visible = True
            lblMonth.Left = 1360
            txtMonth.Left = 1930
            lblYear.Left = 2710
            txtYear.Left = 3130
        End If
        'fix 01/TAIN-DK, 01A/TNDN-DK
        If (GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98") Then
            d = Day(Now)
            m = month(Now)
            Y = Year(Now)
                
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y

            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If

            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
        End If
    End If

End Sub

Private Sub chkDauTho_Click()
Dim d, m, Y As Integer
    If chkDauTho.value = 1 Then
        chkCondensate.value = 0
        chkKhiThien.value = 0
        
        strDauTho = chkDauTho.value
        strLoaiTkDk = "DT"
        
        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
            strLoaiTKThang_PS = "TK_LANPS"
        ElseIf (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93") Then
            strLoaiTKThang_PS = "TK_NAM"
        End If
        
        strKhiThienNhien = 0
        strCondensate = 0
        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
            lblNgay.Visible = True
            txtDay.Visible = True
            lblMonth.Left = 1360
            txtMonth.Left = 1930
            lblYear.Left = 2710
            txtYear.Left = 3130
        End If
        'fix 01/TAIN-DK, 01A/TNDN-DK
        If (GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98") Then
            d = Day(Now)
            m = month(Now)
            Y = Year(Now)
                
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y

            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If

            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
        End If
    End If
End Sub

Private Sub chkKhiThien_Click()
Dim m, Y As Integer
    If chkKhiThien.value = 1 Then
        chkCondensate.value = 0
        chkDauTho.value = 0
        
        strKhiThienNhien = chkKhiThien.value
        strLoaiTkDk = "KTN"
        
        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
            strLoaiTKThang_PS = "TK_LANPS"
        ElseIf (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93") Then
            strLoaiTKThang_PS = "TK_NAM"
        End If
        
        strDauTho = 0
        strCondensate = 0

        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
            lblNgay.Visible = False
            txtDay.Visible = False
            lblMonth.Left = 900
            txtMonth.Left = 1500
            lblYear.Left = 2210
            txtYear.Left = 2630
        End If

        'fix 01/TAIN-DK, 01A/TNDN-DK
        If (GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98") Then
            m = CInt(txtMonth.Text)
            Y = CInt(txtYear.Text)
            If m = 1 Then
                m = 12
                Y = Y - 1
            Else
                m = m - 1
            End If
                
            'txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y

            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If

            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
        End If
    End If

End Sub

Private Sub chkQTNamDau_Click()
   If chkQTNamDau.value = 0 Then
        strLoaiTKQT = "QT_TUNG_NAM"
        chkQTTungNam.value = 1
    Else
        strLoaiTKQT = "QT_NAM_DAU"
        chkQTTungNam.value = 0
    End If
End Sub

Private Sub chkQTTungNam_Click()
    If chkQTTungNam.value = 0 Then
        strLoaiTKQT = "QT_NAM_DAU"
        chkQTNamDau.value = 1
    Else
        strLoaiTKQT = "QT_TUNG_NAM"
        chkQTNamDau.value = 0
    End If
End Sub

Private Sub chkSelectAll_Click()
Dim lCtrl As Long

If blnFPChange Then
    Exit Sub
End If
fpSpread1.Col = 1
For lCtrl = 2 To fpSpread1.MaxRows
    fpSpread1.Row = lCtrl
    If Not fpSpread1.Lock And IIf(fpSpread1.value = 1, True, False) <> chkSelectAll.value Then _
        fpSpread1.value = IIf(chkSelectAll.value, 1, 0)
Next lCtrl
End Sub

Private Sub chkSelectAll_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then fpSpread1.SetFocus
End Sub

Private Sub chkSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnClick = True
End Sub


Private Sub chkTKhaiLanXB_Click()
    Dim m, Y, d As Integer
    Dim dTem, dtem1, dtem2 As Date
    Dim varMenuId As String
    dtem2 = Date
    dTem = DateAdd("D", -1, Date)
    dtem1 = DateAdd("M", -1, Date)
    If chkTKLanPS.value = 1 Or chkTKhaiLanXB.value = 1 Then
        lblNgay.Visible = True
        txtDay.Visible = True
     Else
        lblNgay.Visible = False
        txtDay.Visible = False
    End If
    If chkTKhaiLanXB.value = 1 Then
            SetValueToListDK ("0")
            strLoaiTKThang_PS = "TK_LANPS"
            strQuy = "TK_LANXB"
            'strKieuKy = "D"
            OptChinhthuc.value = True
            lblSolan.Visible = False
            txtSolan.Visible = False
            fpsNgaykhaiBS.Visible = False
            
            
            chkTkhaiThang.value = 0
            chkTKLanPS.value = 0
            frmKy.Height = 3000
            
            cmbQuy.Visible = False
            txtMonth.Visible = True
            
            Set lblLanXuat.Container = frmKy
            lblLanXuat.Top = 1050
            lblLanXuat.Left = 120
            lblLanXuat.Visible = True
            
            Set txtLanXuat.Container = frmKy
            txtLanXuat.Top = 1050
            txtLanXuat.Left = 1200
            txtLanXuat.Visible = True
            
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 1500
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1800
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1800
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1800
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            m = month(dtem2)
            Y = Year(dtem2)
            d = Day(dtem2)
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            
            Frame2.Top = 3300
            
            Set lblNganhKD.Container = frmKy
            lblNganhKD.Top = 2100
            lblNganhKD.Left = 120
            
            Set cboNganhKD.Container = frmKy
            cboNganhKD.Top = 2500
            cboNganhKD.Left = 120
            Call Form_Resize
            
            LoadGrid
     End If
End Sub

Private Sub chkTkhaiThang_Click()
    Dim m, Y, d As Integer
    Dim dTem, dtem1 As Date
    Dim q As Quy
    
    m = month(Date)
    Y = Year(Date)
    
    If strLoaiSacThue = "ToKhaiGTGT" Then
        ' set gia tri default
         If m = 1 Then
                m = 12
                Y = Y - 1
            Else
                m = m - 1
            End If
        txtMonth.Text = m
        txtYear.Text = Y
        If Len(txtMonth.Text) = 1 Then
            txtMonth.Text = "0" & txtMonth.Text
        End If
        
    
        If chkTkhaiThang.value = 1 Then
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Then
                 strQuy = "TK_THANG"
                chkTKQuy.value = 0
                chkTKLanPS.value = 0
                
                OptChinhthuc.value = True
                
                Set lblMonth.Container = frmKy
                lblMonth.Top = 570
                lblMonth.Left = 1360
                
                Set txtMonth.Container = frmKy
                txtMonth.Top = 540
                txtMonth.Left = 1930
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2710
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 3130
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = False
                lblQuy.Visible = False
                lblMonth.Visible = True
                txtMonth.Visible = True
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
                
             ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Then
                 strQuy = "TK_THANG"
                chkTKhaiLanXB.value = 0
                chkTKLanPS.value = 0
                
                OptChinhthuc.value = True
                
                Set lblMonth.Container = frmKy
                lblMonth.Top = 570
                lblMonth.Left = 1360
                
                Set txtMonth.Container = frmKy
                txtMonth.Top = 540
                txtMonth.Left = 1930
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2710
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 3130
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                frmKy.Height = 2400
                Frame2.Top = 2700
                Set lblNganhKD.Container = frmKy
                lblNganhKD.Top = 1600
                lblNganhKD.Left = 120
                
                
                
                Set cboNganhKD.Container = frmKy
                cboNganhKD.Top = 1900
                cboNganhKD.Left = 120
                ' set gia tri nganh nghe kinh doanh cho combo
                'SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
                SetValueToListDK ("1")
                
                lblLanXuat.Visible = False
                txtLanXuat.Visible = False
                
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = False
                lblQuy.Visible = False
                lblMonth.Visible = True
                txtMonth.Visible = True
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
                Call Form_Resize
            Else
                strQuy = "TK_THANG"
                chkTKQuy.value = 0
                OptChinhthuc.value = True
                
                Set lblMonth.Container = frmKy
                lblMonth.Top = 570
                lblMonth.Left = 960
                
                Set txtMonth.Container = frmKy
                txtMonth.Top = 540
                txtMonth.Left = 1530
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2310
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 2730
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
                    frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1600
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1900
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
                End If
    
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = False
                lblQuy.Visible = False
                lblMonth.Visible = True
                txtMonth.Visible = True
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
            End If
        Else
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Then
                ' 3 checkbox nen khong set mac dinh
                'strQuy = ""
            Else
                strQuy = "TK_QUY"
                chkTKQuy.value = 1
                OptChinhthuc.value = True
                Set lblQuy.Container = frmKy
                lblQuy.Top = 570
                lblQuy.Left = 960
                
                ' Set gia tri mac dinh cho Quy
                q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)
                If q.q = 1 Then
                    q.q = 4
                    q.Y = q.Y - 1
                Else
                    q.q = q.q - 1
                End If
                cmbQuy.ListIndex = q.q - 1
                txtYear.Text = q.Y
                
                Set cmbQuy.Container = frmKy
                cmbQuy.Top = 540
                cmbQuy.Left = 1530
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2310
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 2730
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
                    frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1600
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1900
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
                End If
    
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = True
                lblQuy.Visible = True
                
                lblMonth.Visible = False
                txtMonth.Visible = False
                
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
            End If
        End If

    ElseIf strLoaiSacThue = "BC26" Then

        ' set gia tri default
        If m = 1 Then
            m = 12
            Y = Y - 1
        Else
            m = m - 1
        End If

        txtMonth.Text = m
        txtYear.Text = Y

        If Len(txtMonth.Text) = 1 Then
            txtMonth.Text = "0" & txtMonth.Text
        End If
    
        If chkTkhaiThang.value = 1 Then
            strQuy = "TK_THANG"
            chkTKQuy.value = 0
            
            Set lblMonth.Container = frmKy
            lblMonth.Top = 570
            lblMonth.Left = 960
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 540
            txtMonth.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730
            SetControlCaption Me, "frmPeriod"

            cmbQuy.Visible = False
            lblQuy.Visible = False
            lblMonth.Visible = True
            txtMonth.Visible = True
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            ' set ngay dau
            txtNgayDau.Text = "01/" & txtMonth.Text & "/" & txtYear.Text
            ' set ngay cuoi
            Dim temp  As Integer
            Dim temp1 As Date
            temp = CInt(txtMonth.Text) + 1
            If txtMonth.Text = "12" Then
                temp1 = DateSerial(CInt(txtYear.Text) + 1, 1, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            Else
                temp1 = DateSerial(CInt(txtYear.Text), temp, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            End If
            
            frmKy.Height = 1300
            Frame2.Top = 1600
        Else
            strQuy = "TK_QUY"
            chkTKQuy.value = 1
            Set lblQuy.Container = frmKy
            lblQuy.Top = 570
            lblQuy.Left = 960
            
            ' Set gia tri mac dinh cho Quy
            q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)

            If q.q = 1 Then
                q.q = 4
                q.Y = q.Y - 1
            Else
                q.q = q.q - 1
            End If

            cmbQuy.ListIndex = q.q - 1
            txtYear.Text = q.Y
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 540
            cmbQuy.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = True
            lblQuy.Visible = True
            
            lblMonth.Visible = False
            txtMonth.Visible = False
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            frmKy.Height = 1300
            Frame2.Top = 1600
        End If

    Else
        dTem = Date
        dtem1 = DateAdd("M", -1, Date)
        lblNgay.Visible = IIf(chkTkhaiThang.value = 0, True, False)
        txtDay.Visible = IIf(chkTkhaiThang.value = 0, True, False)
        If chkTkhaiThang.value = 0 Then
            strLoaiTKThang_PS = "TK_LANPS"
            m = month(dTem)
            Y = Year(dTem)
            d = Day(dTem)
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            chkTKLanPS.value = 1
        Else
            strLoaiTKThang_PS = "TK_THANG"
            m = month(dtem1)
            Y = Year(dtem1)
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            chkTKLanPS.value = 0
            'chkTKLanPS.Enabled = True
            'chkTkhaiThang.Enabled = False
            
            ' Khi chuyen sang to khai thang mac dinh set la to lan dau
            OptBosung.value = False
            OptChinhthuc.value = True
            
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "70" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "06" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "72" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "81" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "90" Then
                frmKy.Height = 1600
                Frame2.Top = 1700
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                OptChinhthuc.Visible = True
                OptBosung.Visible = True
                lblSolan.Visible = False
                txtSolan.Visible = False
                
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "90" Then
                    Frame2.Top = 2000
                ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Then
                    Frame2.Top = 1920
                    txtMonth.Visible = False
                    lblMonth.Visible = False
                    lblQuy.Visible = True
                    cmbQuy.Visible = True
                    ' Set gia tri mac dinh cho Quy
                    q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)
                    If q.q = 1 Then
                        q.q = 4
                        q.Y = q.Y - 1
                    Else
                        q.q = q.q - 1
                    End If
                    cmbQuy.ListIndex = q.q - 1
                    txtYear.Text = q.Y
                    
                    
                    ' Set loai TK
'                    frmKy.Height = 2400
'                    Frame2.Top = 2700
'                    Set lblNganhKD.Container = frmKy
'                    lblNganhKD.caption = TAX_Utilities_v1.Convert(GetAttribute(GetMessageCellById("0237"), "Msg"), UNICODE, TCVN)
'                    lblNganhKD.Top = 1600
'                    lblNganhKD.Left = 120
'
'
'                    Set cboNganhKD.Container = frmKy
'                    cboNganhKD.Top = 1900
'                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    ' SetValueToList "73"
                End If
                
                 
                
                Call Form_Resize
             Else
                frmKy.Height = 2400
                Frame2.Top = 2700
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                Set lblNganhKD.Container = frmKy
                lblNganhKD.Top = 1550
                lblNganhKD.Left = 120
                
                Set cboNganhKD.Container = frmKy
                cboNganhKD.Top = 1850
                cboNganhKD.Left = 120
                
                OptChinhthuc.Visible = True
                OptBosung.Visible = True
                lblSolan.Visible = False
                txtSolan.Visible = False
                
                Call Form_Resize
             
             End If
        End If
    End If
    chkSelectAll.value = "0"
    LoadGrid
'    If chkTKLanPS.value = 1 Then
'         If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Then
'            Frame2.Visible = False
'            lblSelectAll.Visible = False
'            chkSelectAll.Visible = False
'            fpSpread1.Visible = False
'            Call Form_Resize
'         End If
'    Else
'         If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Then
'            Frame2.Visible = True
'            lblSelectAll.Visible = True
'            chkSelectAll.Visible = True
'            fpSpread1.Visible = True
'            Call Form_Resize
'         End If
'    End If
End Sub

Private Sub chkTKKy_Click()
    Dim m, Y, d As Integer
    Dim dTem, dtem1 As Date
    Dim q As Quy
    
    m = month(Date)
    Y = Year(Date)
    
    If strLoaiSacThue = "BC01" Then
        If chkTKKy.value = 1 Then
            strQuy = "TK_KY"
            chkTKQuy.value = 0
                     
            SetControlCaption Me, "frmPeriod"
            
            lblQuy.caption = "Ky`"
            q = GetKyHienTai(iNgayTaiChinh, iThangTaiChinh)
            cmbQuy.Clear
            cmbQuy.AddItem (1)
            cmbQuy.AddItem (2)
            Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
            txtYear.Text = Y
            cmbQuy.ListIndex = q.q - 1
            Call initNgayDauNgayCuoiKy(CInt(Y), cmbQuy.ListIndex)
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            
            frmKy.Height = 1300
            Frame2.Top = 1600
        Else
            strQuy = "TK_QUY"
            chkTKQuy.value = 1
            Set lblQuy.Container = frmKy
            lblQuy.Top = 570
            lblQuy.Left = 960
            
            ' Set gia tri mac dinh cho Quy
            q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)

            If q.q = 1 Then
                q.q = 4
                q.Y = q.Y - 1
            Else
                q.q = q.q - 1
            End If

            cmbQuy.ListIndex = q.q - 1
            txtYear.Text = q.Y
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 540
            cmbQuy.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = True
            lblQuy.Visible = True
            
            lblMonth.Visible = False
            txtMonth.Visible = False
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            frmKy.Height = 1300
            Frame2.Top = 1600
        End If
    End If
End Sub

Private Sub chkTKLanPS_Click()
    Dim m, Y, d As Integer
    Dim dTem, dtem1, dtem2 As Date
    Dim varMenuId As String
    dtem2 = Date
    dTem = DateAdd("D", -1, Date)
    dtem1 = DateAdd("M", -1, Date)
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Then
        If chkTKLanPS.value = 1 Or chkTKhaiLanXB.value = 1 Then
            lblNgay.Visible = True
            txtDay.Visible = True
         Else
            lblNgay.Visible = False
            txtDay.Visible = False
        End If
    Else
        lblNgay.Visible = IIf(chkTKLanPS.value = 1, True, False)
        txtDay.Visible = IIf(chkTKLanPS.value = 1, True, False)
    End If
    If chkTKLanPS.value = 1 Then
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Then
            strLoaiTKThang_PS = "TK_LANPS"
            strQuy = "TK_LANPS"
            'strKieuKy = "D"
            OptChinhthuc.value = True
            lblSolan.Visible = False
            txtSolan.Visible = False
            fpsNgaykhaiBS.Visible = False
            
            chkTkhaiThang.value = 0
            chkTKQuy.value = 0
            frmKy.Height = 1600
            
            cmbQuy.Visible = False
            lblQuy.Visible = False
            txtMonth.Visible = True
            lblMonth.Visible = True
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 900
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1200
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1200
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1200
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False
                        
            m = month(dtem2)
            Y = Year(dtem2)
            d = Day(dtem2)
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            
            Call Form_Resize
            
        ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Then
            strLoaiTKThang_PS = "TK_LANPS"
            strQuy = "TK_LANPS"
            'strKieuKy = "D"
            OptChinhthuc.value = True
            lblSolan.Visible = False
            txtSolan.Visible = False
            fpsNgaykhaiBS.Visible = False
            
            m = month(dtem2)
            Y = Year(dtem2)
            d = Day(dtem2)
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            chkTkhaiThang.value = 0
            chkTKhaiLanXB.value = 0
            frmKy.Height = 2400
            Frame2.Top = 2700
            
            cmbQuy.Visible = False
            txtMonth.Visible = True
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 900
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1200
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1200
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1200
            txtSolan.Left = 3400
            
            Set lblNganhKD.Container = frmKy
            lblNganhKD.Top = 1600
            lblNganhKD.Left = 120
            
            
            
            Set cboNganhKD.Container = frmKy
            cboNganhKD.Top = 1900
            cboNganhKD.Left = 120
            
            lblLanXuat.Visible = False
            txtLanXuat.Visible = False
            
            lblSolan.Visible = False
            txtSolan.Visible = False
            Call Form_Resize
        Else
            strLoaiTKThang_PS = "TK_LANPS"
            'strKieuKy = "D"
            OptChinhthuc.value = True
            lblSolan.Visible = False
            txtSolan.Visible = False
            fpsNgaykhaiBS.Visible = False
            
            m = month(dTem)
            Y = Year(dTem)
            d = Day(dTem)
            txtDay.Text = d
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtDay.Text) = 1 Then
                txtDay.Text = "0" & txtDay.Text
            End If
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            chkTkhaiThang.value = 0
            'chkTkhaiThang.Enabled = True
            'chkTKLanPS.Enabled = False
            'chkTKLanPS.
            varMenuId = GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
            If varMenuId = "70" Or varMenuId = "06" Or varMenuId = "72" Or varMenuId = "73" Or varMenuId = "81" Or varMenuId = "90" Then
            'If varMenuId = "73" Or varMenuId = "81" Then
                 frmKy.Height = 1600
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
    '            frmKy.Height = 1065
    '            Frame2.Top = 1400
    '
    '            Set OptChinhthuc.Container = frmKy
    '            OptChinhthuc.Top = 8000
    '            OptChinhthuc.Left = 960
    '
    '            Set OptBosung.Container = frmKy
    '            OptBosung.Top = 8600
    '            OptBosung.Left = 960
    '
    '            Set lblSolan.Container = frmKy
    '            lblSolan.Top = 9200
    '            lblSolan.Left = 3000
    '            Set txtSolan.Container = frmKy
    '            txtSolan.Top = 9400
    '            txtSolan.Left = 3400
                
    '            OptChinhthuc.Visible = False
    '            OptBosung.Visible = False
    '            lblSolan.Visible = False
    '            txtSolan.Visible = False
                ' To khai TNDN an quy va hien thi thang
                If varMenuId = "73" Then
                    txtMonth.Visible = True
                    lblMonth.Visible = True
                    cmbQuy.Visible = False
                    lblQuy.Visible = False
                    
                    ' Set loai TK
    '                frmKy.Height = 1700
    '                Frame2.Top = 2000
    '                Set lblNganhKD.Container = frmKy
    '                lblNganhKD.caption = TAX_Utilities_v1.Convert(GetAttribute(GetMessageCellById("0237"), "Msg"), UNICODE, TCVN)
    '                lblNganhKD.Top = 950
    '                lblNganhKD.Left = 120
    '
    '
    '                Set cboNganhKD.Container = frmKy
    '                cboNganhKD.Top = 1250
    '                cboNganhKD.Left = 120
                End If
                
                Call Form_Resize
            ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "05" Then
    '            frmKy.Height = 1465
    '            Frame2.Top = 1800
    '
    '            Set OptChinhthuc.Container = frmKy
    '            OptChinhthuc.Top = 8000
    '            OptChinhthuc.Left = 960
    '
    '            Set OptBosung.Container = frmKy
    '            OptBosung.Top = 8600
    '            OptBosung.Left = 960
    '
    '            Set lblSolan.Container = frmKy
    '            lblSolan.Top = 9200
    '            lblSolan.Left = 3000
    '            Set txtSolan.Container = frmKy
    '            txtSolan.Top = 9400
    '            txtSolan.Left = 3400
    '
    '            Set lblNganhKD.Container = frmKy
    '            lblNganhKD.Top = 1550
    '            lblNganhKD.Left = 120
    '
    '            Set cboNganhKD.Container = frmKy
    '            cboNganhKD.Top = 1000
    '            cboNganhKD.Left = 120
    '
    '
    '            OptChinhthuc.Visible = False
    '            OptBosung.Visible = False
    '            lblSolan.Visible = False
    '            txtSolan.Visible = False
                frmKy.Height = 2400
                Frame2.Top = 2700
                
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                
        
                Set lblNganhKD.Container = frmKy
                lblNganhKD.Top = 1550
                lblNganhKD.Left = 120
                
                Set cboNganhKD.Container = frmKy
                cboNganhKD.Top = 1850
                cboNganhKD.Left = 120
                Call Form_Resize
            End If
        End If
    Else
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Then
            ' khòn set mac dinh cho to 04/GTGT
            'strQuy = ""
        Else
            strLoaiTKThang_PS = "TK_THANG"
            m = month(dtem1)
            Y = Year(dtem1)
            txtMonth.Text = m
            txtYear.Text = Y
            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If
            chkTkhaiThang.value = 1
        End If
    End If
    LoadGrid
'    If chkTKLanPS.value = 1 Then
'         If varMenuId = "73" Then
'            Frame2.Visible = False
'            lblSelectAll.Visible = False
'            chkSelectAll.Visible = False
'            fpSpread1.Visible = False
'            Call Form_Resize
'         End If
'    Else
'         If varMenuId = "73" Then
'            Frame2.Visible = True
'            lblSelectAll.Visible = True
'            chkSelectAll.Visible = True
'            fpSpread1.Visible = True
'            Call Form_Resize
'         End If
'    End If
End Sub

Private Sub chkTKQuy_Click()
    Dim m, Y, d As Integer
    Dim dTem, dtem1 As Date
    Dim q As Quy
    
    m = month(Date)
    Y = Year(Date)
    If strLoaiSacThue = "ToKhaiGTGT" Then
        If chkTKQuy.value = 0 Then
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Then
                ' khong set mac dinh cho to khai 04/GTGT
                'strQuy = ""
            Else
                strQuy = "TK_THANG"
                chkTkhaiThang.value = 1
                
                OptChinhthuc.value = True
                
                
                Set lblMonth.Container = frmKy
                lblMonth.Top = 570
                lblMonth.Left = 960
                
                Set txtMonth.Container = frmKy
                txtMonth.Top = 540
                txtMonth.Left = 1530
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2310
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 2730
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
                    frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1600
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1900
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
                End If
    
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = False
                lblQuy.Visible = False
                lblMonth.Visible = True
                txtMonth.Visible = True
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
            End If
        Else
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Then
                strQuy = "TK_QUY"
                chkTkhaiThang.value = 0
                chkTKLanPS.value = 0
                
                OptChinhthuc.value = True
                
                Set lblQuy.Container = frmKy
                lblQuy.Top = 570
                lblQuy.Left = 1360
                
                Set cmbQuy.Container = frmKy
                cmbQuy.Top = 540
                cmbQuy.Left = 1930
                
                
                ' Set gia tri mac dinh cho Quy
                q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)
    
                If q.q = 1 Then
                    q.q = 4
                    q.Y = q.Y - 1
                Else
                    q.q = q.q - 1
                End If
    
                cmbQuy.ListIndex = q.q - 1
                txtYear.Text = q.Y
                
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2710
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 3130
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
                    frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1600
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1900
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
                End If
    
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = True
                lblQuy.Visible = True
                
                lblMonth.Visible = False
                txtMonth.Visible = False
                
                 
                
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
            Else
                strQuy = "TK_QUY"
                chkTkhaiThang.value = 0
                
                OptChinhthuc.value = True
                
                Set lblQuy.Container = frmKy
                lblQuy.Top = 570
                lblQuy.Left = 960
                
                Set cmbQuy.Container = frmKy
                cmbQuy.Top = 540
                cmbQuy.Left = 1530
                
                Set lblYear.Container = frmKy
                lblYear.Top = 570
                lblYear.Left = 2310
                
                Set txtYear.Container = frmKy
                txtYear.Top = 540
                txtYear.Left = 2730
                
                Set OptChinhthuc.Container = frmKy
                OptChinhthuc.Top = 900
                OptChinhthuc.Left = 960
                
                Set OptBosung.Container = frmKy
                OptBosung.Top = 1200
                OptBosung.Left = 960
                
                Set lblSolan.Container = frmKy
                lblSolan.Top = 1200
                lblSolan.Left = 3000
                Set txtSolan.Container = frmKy
                txtSolan.Top = 1200
                txtSolan.Left = 3400
                
                lblSolan.Visible = False
                txtSolan.Visible = False
                fpsNgaykhaiBS.Visible = False
                
                ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
                    frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1600
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1900
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
                End If
    
                SetControlCaption Me, "frmPeriod"
       
                cmbQuy.Visible = True
                lblQuy.Visible = True
                
                lblMonth.Visible = False
                txtMonth.Visible = False
                
                 
                
                txtNgayDau.Visible = False
                txtNgayCuoi.Visible = False
            End If
        End If

    ElseIf strLoaiSacThue = "BC26" Then

        If chkTKQuy.value = 0 Then
            strQuy = "TK_THANG"
            chkTkhaiThang.value = 1
            
            Set lblMonth.Container = frmKy
            lblMonth.Top = 570
            lblMonth.Left = 960
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 540
            txtMonth.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730

            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            lblQuy.Visible = False
            lblMonth.Visible = True
            txtMonth.Visible = True
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            
            ' set ngay dau
            txtNgayDau.Text = "01/" & txtMonth.Text & "/" & txtYear.Text
            ' set ngay cuoi
            Dim temp  As Integer
            Dim temp1 As Date
            temp = CInt(txtMonth.Text) + 1
            If txtMonth.Text = "12" Then
                temp1 = DateSerial(CInt(txtYear.Text) + 1, 1, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            Else
                temp1 = DateSerial(CInt(txtYear.Text), temp, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            End If
            
            frmKy.Height = 1300
            Frame2.Top = 1600
        Else
            strQuy = "TK_QUY"
            chkTkhaiThang.value = 0
                    
            Set lblQuy.Container = frmKy
            lblQuy.Top = 570
            lblQuy.Left = 960
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 540
            cmbQuy.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = True
            lblQuy.Visible = True
            
            lblMonth.Visible = False
            txtMonth.Visible = False
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            frmKy.Height = 1300
            Frame2.Top = 1600
            
            LoadDefaultInfor
        End If
    ElseIf strLoaiSacThue = "BC01" Then

        If chkTKQuy.value = 0 Then
            strQuy = "TK_KY"
            chkTKKy.value = 1
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730

            SetControlCaption Me, "frmPeriod"
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            
            frmKy.Height = 1300
            Frame2.Top = 1600
        Else
            strQuy = "TK_QUY"
            chkTKKy.value = 0
                    
            Set lblQuy.Container = frmKy
            lblQuy.Top = 570
            lblQuy.Left = 960
            lblQuy.caption = "Quý"
                         
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 540
            cmbQuy.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 570
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 540
            txtYear.Left = 2730
            
            SetControlCaption Me, "frmPeriod"
            
            ' Set gia tri mac dinh cho Quy
            q = GetKyHienTai(iNgayTaiChinh, iThangTaiChinh)
            cmbQuy.Clear
            cmbQuy.AddItem (1)
            cmbQuy.AddItem (2)
            cmbQuy.AddItem (3)
            cmbQuy.AddItem (4)
            Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
            txtYear.Text = Y
            cmbQuy.ListIndex = q.q - 1
            Call initNgayDauNgayCuoiKy(CInt(txtYear.Text), cmbQuy.ListIndex)
   
   
            cmbQuy.Visible = True
            lblQuy.Visible = True
            
            txtNgayDau.Visible = True
            txtNgayCuoi.Visible = True
            lblNgayDau.Visible = True
            lblNgayCuoi.Visible = True
            frmKy.Height = 1300
            Frame2.Top = 1600
        End If
    Else
        If chkTKQuy.value = 0 Then
            strQuy = "TK_TU_THANG"
            chkTuThangDenThang.value = 1
            lblQuy.Visible = False
            cmbQuy.Visible = False
            lblYear.Visible = False
            txtYear.Visible = False
            
            'frmKy.Height = 1600
            frmKy.Height = 2200
            
            Set lblTuThang.Container = frmKy
            lblTuThang.Top = 570
            lblTuThang.Left = 180
            lblTuThang.Visible = True
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 540
            txtNgayDau.Left = 1000
            txtNgayDau.Visible = True
            
            Set lblDenThang.Container = frmKy
            lblDenThang.Top = 570
            lblDenThang.Left = 2200
            lblDenThang.Visible = True
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 540
            txtNgayCuoi.Left = 3100
            txtNgayCuoi.Visible = True
            
            Set chkQTTungNam.Container = frmKy
            chkQTTungNam.Top = 920
            chkQTTungNam.Left = 960
            chkQTTungNam.Visible = True
            
            Set chkQTNamDau.Container = frmKy
            chkQTNamDau.Top = 1200
            chkQTNamDau.Left = 960
            chkQTNamDau.Visible = True
            
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 1500
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1800
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1800
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1800
            txtSolan.Left = 3400
            
            OptChinhthuc.Visible = True
            OptBosung.Visible = True
            lblSolan.Visible = False
            txtSolan.Visible = False
    
            
        Else
            strQuy = "TK_QUY"
            chkTuThangDenThang.value = 0
            lblQuy.Visible = True
            cmbQuy.Visible = True
            lblYear.Visible = True
            txtYear.Visible = True
            
            frmKy.Height = 1600
            
            Set lblTuThang.Container = frmKy
            lblTuThang.Top = 570
            lblTuThang.Left = 500
            lblTuThang.Visible = False
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 5120
            txtNgayDau.Left = 1930
            txtNgayDau.Visible = False
            
            Set lblDenThang.Container = frmKy
            lblDenThang.Top = 570
            lblDenThang.Left = 2500
            lblDenThang.Visible = False
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 5220
            txtNgayCuoi.Left = 2700
            txtNgayCuoi.Visible = False
            
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 900
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1200
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1200
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1200
            txtSolan.Left = 3400
            
            Set chkQTTungNam.Container = frmKy
            chkQTTungNam.Top = 10920
            chkQTTungNam.Left = 960
            chkQTTungNam.Visible = False
            
            Set chkQTNamDau.Container = frmKy
            chkQTNamDau.Top = 10920
            chkQTNamDau.Left = 960
            chkQTNamDau.Visible = False
    
            
            lblSolan.Visible = False
            txtSolan.Visible = False
                    
        End If
    End If
    Call Form_Resize
    chkSelectAll.value = "0"
    LoadGrid

End Sub

Private Sub chkTuThangDenThang_Click()
    If chkTuThangDenThang.value = 1 Then
        strQuy = "TK_TU_THANG"
        lblQuy.Visible = False
        cmbQuy.Visible = False
        lblYear.Visible = False
        txtYear.Visible = False
        
        'frmKy.Height = 1600
        frmKy.Height = 2200
        
        Set lblTuThang.Container = frmKy
        lblTuThang.Top = 570
        lblTuThang.Left = 180
        lblTuThang.Visible = True
        
        Set txtNgayDau.Container = frmKy
        txtNgayDau.Top = 540
        txtNgayDau.Left = 1000
        txtNgayDau.Visible = True
        
        Set lblDenThang.Container = frmKy
        lblDenThang.Top = 570
        lblDenThang.Left = 2200
        lblDenThang.Visible = True
        
        Set txtNgayCuoi.Container = frmKy
        txtNgayCuoi.Top = 540
        txtNgayCuoi.Left = 3100
        txtNgayCuoi.Visible = True
        
        Set chkQTTungNam.Container = frmKy
        chkQTTungNam.Top = 920
        chkQTTungNam.Left = 960
        chkQTTungNam.Visible = True
        chkQTTungNam.value = 1
        
        Set chkQTNamDau.Container = frmKy
        chkQTNamDau.Top = 1200
        chkQTNamDau.Left = 960
        chkQTNamDau.Visible = True
        
        
        Set OptChinhthuc.Container = frmKy
        OptChinhthuc.Top = 1500
        OptChinhthuc.Left = 960
        
        Set OptBosung.Container = frmKy
        OptBosung.Top = 1800
        OptBosung.Left = 960
        
        Set lblSolan.Container = frmKy
        lblSolan.Top = 1800
        lblSolan.Left = 3000
        Set txtSolan.Container = frmKy
        txtSolan.Top = 1800
        txtSolan.Left = 3400
        
        OptChinhthuc.Visible = True
        OptBosung.Visible = True
        lblSolan.Visible = False
        txtSolan.Visible = False
        chkTKQuy.value = 0
    Else
        strQuy = "TK_QUY"
        lblQuy.Visible = True
        cmbQuy.Visible = True
        lblYear.Visible = True
        txtYear.Visible = True
        
        frmKy.Height = 1600
        
        Set lblTuThang.Container = frmKy
        lblTuThang.Top = 570
        lblTuThang.Left = 500
        lblTuThang.Visible = False
        
        Set txtNgayDau.Container = frmKy
        txtNgayDau.Top = 5120
        txtNgayDau.Left = 1930
        txtNgayDau.Visible = False
        
        Set lblDenThang.Container = frmKy
        lblDenThang.Top = 570
        lblDenThang.Left = 2500
        lblDenThang.Visible = False
        
        Set txtNgayCuoi.Container = frmKy
        txtNgayCuoi.Top = 5220
        txtNgayCuoi.Left = 2700
        txtNgayCuoi.Visible = False
        
        
        Set OptChinhthuc.Container = frmKy
        OptChinhthuc.Top = 900
        OptChinhthuc.Left = 960
        
        Set OptBosung.Container = frmKy
        OptBosung.Top = 1200
        OptBosung.Left = 960
        
        Set lblSolan.Container = frmKy
        lblSolan.Top = 1200
        lblSolan.Left = 3000
        Set txtSolan.Container = frmKy
        txtSolan.Top = 1200
        txtSolan.Left = 3400
        
        Set chkQTTungNam.Container = frmKy
        chkQTTungNam.Top = 10920
        chkQTTungNam.Left = 960
        chkQTTungNam.Visible = False
        chkQTTungNam.value = 0
        
        Set chkQTNamDau.Container = frmKy
        chkQTNamDau.Top = 10920
        chkQTNamDau.Left = 960
        chkQTNamDau.Visible = False
        chkQTNamDau.value = 0
        
        lblSolan.Visible = False
        txtSolan.Visible = False
        chkTKQuy.value = 1
    End If
    LoadGrid
End Sub

' Ham phuc vu cac mau an chi
Private Sub cmbQuy_LostFocus()
'dhdang sua loi thay doi quy cap nhat lai check box phu luc
'ngay 21/01/2011
    If (cmbQuy.Text & "/" & txtYear.Text <> TAX_Utilities_v1.ThreeMonths & "/" & yChange) Then
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "Year") = "1/2" And txtNgayCuoi.Enabled And txtNgayDau.Enabled Then
            Call initNgayDauNgayCuoiKy(CInt(txtYear.Text), cmbQuy.ListIndex)
        End If
        LoadGrid
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
    strKHBS = ""
    frmTreeviewMenu.Show
End Sub

Private Sub cmdClose_LostFocus()
    bIsClosed = False
End Sub

Private Sub cmdClose_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bIsClosed = True
End Sub

'****************************************************
'Description:cmdOK_Click procedure set value for variable common
'   Step 1: Check if user hasn't type value
'   Step 2: Set value strYear, strMonth, str3Month
'****************************************************

Public Sub cmdOK_Click()
    On Error GoTo ErrorHandle
    Dim frmTK As frmInterfaces
    Dim dNgayDau As Date
    Dim dNgayCuoi As Date
    Dim dNgayDauQuy As Date
    Dim dNgayCuoiQuy As Date
    Dim sNgay As String
    Dim sNgayDD As Date
    Dim objDateUtils As DateUtils
    
    Dim strTempValue As String
    
    Dim idToKhai As String
    
    If OptBosung.value = True Then
    strSolanBS = txtSolan.Text
    Else
    strSolanBS = ""
    End If
    
    ' xy ly cho bang ke NPT
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "95" Then
        strSolanKK = txtSolan.Text
    End If
    
    ' set so lan xuat ban dau tho
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98" Then
        If chkTKhaiLanXB.value = 1 Then
            strSoLanXuatBan = txtLanXuat.Text
        Else
            strSoLanXuatBan = ""
        End If
    Else
        strSoLanXuatBan = ""
    End If
    
    
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "64" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "07" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "91" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "27" Then
        txtDay_LostFocus
        txtMonth_LostFocus
        txtYear_LostFocus
    End If
    
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "37" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "38" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "40") And Val(txtYear.Text) >= 2010 Then
        DisplayMessage "0176", msOKOnly, miInformation
        Exit Sub
    End If
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "53" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "54") And ((Val(txtYear.Text) = 2010 And txtMonth.Text <> "01") Or Val(txtYear.Text) >= 2011) Then
        DisplayMessage "0176", msOKOnly, miInformation
        Exit Sub
    End If
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "39") And ((Val(txtYear.Text) = 2010 And txtMonth.Text <> "01") Or Val(txtYear.Text) >= 2011) Then
        DisplayMessage "0176", msOKOnly, miInformation
        Exit Sub
    End If
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "15" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "50") And Val(txtYear.Text) = 2010 And txtMonth.Text = "01" Then
        DisplayMessage "0177", msOKOnly, miInformation
        Exit Sub
    End If
    
    ' validate cho to 04TBAC
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "91") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "64") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "07") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "27") Then
        If objCvt Is Nothing Then
            Set objCvt = New DateUtils
        End If
        
        If (objCvt.ToDate(txtDay.Text + "/" + txtMonth.Text + "/" + txtYear.Text, "DD/MM/YYYY") > Date) Then
            DisplayMessage "0310", msOKOnly, miInformation
            Exit Sub
        End If
    End If
        
    ' validate cho to 04TBAC,01/TAIN-DK
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98") Then
'        If (chkKhiThien.value = "1") Then
'            If (objCvt.ToDate("01" + "/" + txtMonth.Text + "/" + txtYear.Text, "DD/MM/YYYY") > Date) Then
'                DisplayMessage "0188", msOKOnly, miInformation
'                Exit Sub
'            End If
'
'        Else
'
'            If (objCvt.ToDate(txtDay.Text + "/" + txtMonth.Text + "/" + txtYear.Text, "DD/MM/YYYY") > Date) Then
'                DisplayMessage "0188", msOKOnly, miInformation
'                Exit Sub
'            End If
'        End If
         If strQuy = "TK_LANPS" Or strQuy = "TK_LANXB" Then
            If (objCvt.ToDate(txtDay.Text + "/" + txtMonth.Text + "/" + txtYear.Text, "DD/MM/YYYY") > Date) Then
                DisplayMessage "0188", msOKOnly, miInformation
                Exit Sub
            End If
        End If
    End If
        
    ' validate cho to 04TBAC
'    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89") Then
'        If (chkDauTho.value = 0 And chkCondensate.value = 0 And chkKhiThien.value = 0) Then
'            DisplayMessage "0017", msOKOnly, miInformation
'            Exit Sub
'        End If
'    End If
    
    If strKieuKy = KIEU_KY_NGAY_NAM Then
        txtNgayDau_LostFocus
        If Not blnValidInfo(3) Then Exit Sub
        txtNgayCuoi_LostFocus
        If Not blnValidInfo(4) Then Exit Sub
    End If
    
    
    'end
    'dhdang
    ' comment???
    If strKieuKy <> "H_Y" Then
            txtMonth_LostFocus
            If Not blnValidInfo(1) Then Exit Sub
            txtYear_LostFocus
            If Not blnValidInfo(2) Then Exit Sub
    End If
    ' end
    
    'requirement
    If Len(txtMonth.Text) = 0 And txtMonth.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtMonth.SetFocus
        Exit Sub
    '****************
    ' added
    ElseIf Val(txtMonth.Text) = 0 And txtMonth.Visible Then
    '****************
        DisplayMessage "0018", msOKOnly, miInformation
        txtMonth.SetFocus
        Exit Sub
    End If
    If Len(cmbQuy.Text) = 0 And cmbQuy.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtMonth.SetFocus
        Exit Sub
    ElseIf Val(cmbQuy.Text) = 0 And cmbQuy.Visible Then
        DisplayMessage "0018", msOKOnly, miInformation
        txtMonth.SetFocus
        Exit Sub
    End If
    If Len(txtYear.Text) = 0 And txtYear.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtYear.SetFocus
        Exit Sub
    ElseIf Val(txtYear.Text) = 0 And txtYear.Visible Then
        DisplayMessage "0018", msOKOnly, miInformation
        txtYear.SetFocus
        Exit Sub
    End If
    ' Kiem tra txtDay
    If Len(txtDay.Text) = 0 And txtDay.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtDay.SetFocus
        Exit Sub
    ElseIf Val(txtDay.Text) = 0 And txtDay.Visible Then
        DisplayMessage "0018", msOKOnly, miInformation
        txtDay.SetFocus
        Exit Sub
    End If
    
    If Len(txtNgayDau.Text) = 0 And txtNgayDau.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtNgayDau.SetFocus
        Exit Sub
    ElseIf Val(txtNgayDau.Text) = 0 And txtNgayDau.Visible Then
        DisplayMessage "0018", msOKOnly, miInformation
        txtNgayDau.SetFocus
        Exit Sub
    End If
    If Len(txtNgayCuoi.Text) = 0 And txtNgayCuoi.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtNgayCuoi.SetFocus
        Exit Sub
    ElseIf Val(txtNgayCuoi.Text) = 0 And txtNgayCuoi.Visible Then
        DisplayMessage "0018", msOKOnly, miInformation
        txtNgayCuoi.SetFocus
        Exit Sub
    End If
    
    If Len(txtSolan.Text) = 0 And txtSolan.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtSolan.SetFocus
        Exit Sub
    ElseIf Val(txtSolan.Text) = 0 And txtSolan.Visible Then
        DisplayMessage "0018", msOKOnly, miInformation
        txtSolan.SetFocus
        Exit Sub
    End If
    
    
    If Len(txtLanXuat.Text) = 0 And txtLanXuat.Visible Then
        DisplayMessage "0017", msOKOnly, miInformation
        txtLanXuat.SetFocus
        Exit Sub
    End If
    '***************************
    'Check period with valid date
    If CInt(txtYear.Text) < CInt(Right$(GetAttribute(TAX_Utilities_v1.NodeMenu.childNodes(0), "StartDate"), 4)) Then
        DisplayMessage "0092", msOKOnly, miCriticalError, , mrOK
        txtYear.SetFocus
        Exit Sub
    End If
    '***************************
    
    '***************************
    If strKieuKy = KIEU_KY_THANG Then
        If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "96" _
        Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "94" Then

            If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "71" Then
                If strQuy = "TK_THANG" Then
                     If Not CheckPeriod(txtMonth.Text, txtYear.Text) Then
                        txtMonth.SetFocus
                        Exit Sub
                    End If
                ElseIf strQuy = "TK_QUY" Then
                    If Not CheckPeriod(cmbQuy.Text, txtYear.Text) Then
                        cmbQuy.SetFocus
                        Exit Sub
                    End If
                    
                    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "71" Then

                        If (Val(TAX_Utilities_v1.ThreeMonths) < 3 And Val(TAX_Utilities_v1.Year) = 2013) Or Val(TAX_Utilities_v1.Year) < 2013 Then
                            DisplayMessage "0272", msOKOnly, miCriticalError
                            cmbQuy.SetFocus
                            Exit Sub
                         End If
                    End If
                ElseIf strQuy = "TK_LANPS" Then
                    
                End If
            Else
                If strQuy = "TK_THANG" Then
                     If Not CheckPeriod(txtMonth.Text, txtYear.Text) Then
                        txtMonth.SetFocus
                        Exit Sub
                    End If
                ElseIf strQuy = "TK_QUY" Then
                    If Not CheckPeriod(cmbQuy.Text, txtYear.Text) Then
                        cmbQuy.SetFocus
                        Exit Sub
                    End If
                    
                    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "71" Then

                        If (Val(TAX_Utilities_v1.ThreeMonths) < 3 And Val(TAX_Utilities_v1.Year) = 2013) Or Val(TAX_Utilities_v1.Year) < 2013 Then
                            DisplayMessage "0272", msOKOnly, miCriticalError
                            cmbQuy.SetFocus
                            Exit Sub
                         End If
                    End If
                End If
            End If
        Else
            If Not CheckPeriod(txtMonth.Text, txtYear.Text) Then
                txtMonth.SetFocus
                Exit Sub
            End If
        End If
    ElseIf strKieuKy = KIEU_KY_QUY Then
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" And chkTKLanPS.value = "1" Then
        Else
            If Not CheckPeriod(cmbQuy.Text, txtYear.Text) Then
                cmbQuy.SetFocus
                Exit Sub
            End If
        End If
        ' BC 26
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "68" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "14" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "13" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "18" Then
            If strQuy = "TK_THANG" Then
                 If Not CheckPeriod(txtMonth.Text, txtYear.Text) Then
                    txtMonth.SetFocus
                    Exit Sub
                End If
            ElseIf strQuy = "TK_QUY" Then
                If Not CheckPeriod(cmbQuy.Text, txtYear.Text) Then
                    cmbQuy.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        ' To khai 02/TNDN
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Then
            If txtDay.Text <> "" Then
            sNgay = Right("00" & txtDay.Text, 2) & "/" & Right("00" & txtMonth.Text, 2) & "/" & Right("0000" & txtYear.Text, 4)
                If IsNull(objCvt.ToDate(sNgay, "DD/MM/YYYY")) Then
                    DisplayMessage "0071", msOKOnly, miCriticalError
                    txtDay.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        '01/KK-TTS
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "23" Then
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
        End If
    ElseIf strKieuKy = KIEU_KY_NAM Then
        If Not CheckPeriod("1", txtYear.Text) Then
            txtYear.SetFocus
            Exit Sub
        End If
        
        ' 02/TNDN-DK
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "87" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "97" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "77" _
        Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "88" Then
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
            ' check khong dc vuot qua 15 thang
            If DateDiff("M", format(txtNgayDau.Text, "mm/yyyy"), format(txtNgayCuoi.Text, "mm/yyyy")) + 1 > 15 Then
                    DisplayMessage "0335", msOKOnly, miInformation
                    txtNgayCuoi.SetFocus
                    Exit Sub
            End If
            ' check tu thang thuoc ky tinh thue
            If Right$(txtNgayDau.Text, 4) <> TAX_Utilities_v1.Year Then
                    DisplayMessage "0336", msOKOnly, miInformation
                    txtNgayCuoi.SetFocus
                    Exit Sub
            End If
        End If
        
        ' QT TNCN
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "76" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "59" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "43" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "41" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "17" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "26" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "45" Then
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
            ' check khong dc vuot qua 12 thang
            If DateDiff("M", format(txtNgayDau.Text, "mm/yyyy"), format(txtNgayCuoi.Text, "mm/yyyy")) + 1 > 12 Then
                    DisplayMessage "0339", msOKOnly, miInformation
                    txtNgayCuoi.SetFocus
                    Exit Sub
            End If
             ' check tu thang thuoc ky tinh thue
            If Right$(txtNgayDau.Text, 4) <> TAX_Utilities_v1.Year Then
                    DisplayMessage "0336", msOKOnly, miInformation
                    txtNgayCuoi.SetFocus
                    Exit Sub
            End If
        End If
    
   ' dntai them vao ngay 08/05/2011
    ElseIf strKieuKy = "H_Y" Then
        ' BC 26
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "68" Then
            If strQuy = "TK_THANG" Then
                 If Not CheckPeriod(txtMonth.Text, txtYear.Text) Then
                    txtMonth.SetFocus
                    Exit Sub
                End If
            ElseIf strQuy = "TK_QUY" Then
                If Not CheckPeriod(cmbQuy.Text, txtYear.Text) Then
                    cmbQuy.SetFocus
                    Exit Sub
                End If
            End If
        Else
            If Not CheckPeriod(cmbQuy.Text, txtYear.Text) Then
                txtYear.SetFocus
                Exit Sub
            End If
        End If
    ElseIf strKieuKy = KIEU_KY_NGAY_NAM Then
        If Not CheckPeriod("1", txtYear.Text) Then
            txtYear.SetFocus
            Exit Sub
        End If
        
        Set objDateUtils = New DateUtils
        dNgayDauQuy = GetNgayDauQuy(4, CInt(txtYear.Text) - 1, iNgayTaiChinh, iThangTaiChinh)
        dNgayCuoiQuy = GetNgayCuoiQuy(1, CInt(txtYear.Text) + 1, iNgayTaiChinh, iThangTaiChinh)
        dNgayDau = objDateUtils.ToDate(txtNgayDau, "DD/MM/YYYY")
        dNgayCuoi = objDateUtils.ToDate(txtNgayCuoi, "DD/MM/YYYY")
        ' Neu to khai khong phai la to khai quyet toan TNCN thi khong kiem tra ngay bat dau nam tai chinh
        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "17") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "41") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "42") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "43") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "26") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "59") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "44") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue <> "45") Then

            If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "80") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "82") Then
'                ' To khai 02/NTNN vaf 04/NTNN se khong check dk nay
'                If DateDiff("M", dNgayDau, dNgayCuoi) + 1 > 15 Then
'                    DisplayMessage "0068", msOKOnly, miInformation
'                    txtNgayCuoi.SetFocus
'                    Exit Sub
'                End If
            Else
                'Cap nhat to 02/PHLP
                If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "03" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "88" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "87" _
                Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "97" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "77" Then
                Else
                    If dNgayDau < dNgayDauQuy Then
                        DisplayMessage "0065", msOKOnly, miInformation
                        txtNgayDau.SetFocus
                        Exit Sub
                    End If
                    If dNgayCuoi > dNgayCuoiQuy Then
                        DisplayMessage "0066", msOKOnly, miInformation
                        txtNgayCuoi.SetFocus
                        Exit Sub
                    End If
                End If
                If DateDiff("M", dNgayDau, dNgayCuoi) + 1 > 15 Then
                    DisplayMessage "0068", msOKOnly, miInformation
                    txtNgayCuoi.SetFocus
                    Exit Sub
                End If
            End If
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0069", msOKOnly, miInformation
                txtNgayCuoi.SetFocus
                Exit Sub
            End If
            ' To khai 02/NTNN vaf 04/NTNN them dk den thang phai nho hon hoac bang thang hien tai
            If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "80") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "82") Then
                If DateDiff("D", DateSerial(Year(dNgayCuoi), month(dNgayCuoi), 1), DateSerial(Year(Date), month(Date), 1)) < 0 Then
                    DisplayMessage "0247", msOKOnly, miInformation
                    txtNgayCuoi.SetFocus
                    Exit Sub
                End If
            End If
        End If
        txtNgayCuoi.Text = objDateUtils.ToString(dNgayCuoi, "DD/MM/YYYY")
      ElseIf strKieuKy = KIEU_KY_NGAY_THANG Then
        If Not CheckPeriod(txtMonth.Text, txtYear.Text) Then
            txtMonth.SetFocus
            Exit Sub
        End If
'htphuong add check valid ngay mau 05/GTGT
        If txtDay.Text <> "" Then
        sNgay = Right("00" & txtDay.Text, 2) & "/" & Right("00" & txtMonth.Text, 2) & "/" & Right("0000" & txtYear.Text, 4)
            If IsNull(objCvt.ToDate(sNgay, "DD/MM/YYYY")) Then
                DisplayMessage "0071", msOKOnly, miCriticalError
                txtDay.SetFocus
                Exit Sub
            End If
        End If
    End If
    
   ' Kiem tra to khai 07/TNCN
   If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "36") Then
        If strQuy = "TK_THANG" Then
            If (Val(TAX_Utilities_v1.month) >= 7 And Val(TAX_Utilities_v1.Year) = 2013) Or Val(TAX_Utilities_v1.Year) > 2013 Then
                DisplayMessage "0266", msOKOnly, miCriticalError
                txtMonth.SetFocus
                Exit Sub
            End If
        ElseIf strQuy = "TK_QUY" Then
            If (Val(TAX_Utilities_v1.ThreeMonths) < 3 And Val(TAX_Utilities_v1.Year) = 2013) Or Val(TAX_Utilities_v1.Year) < 2013 Then
                DisplayMessage "0267", msOKOnly, miCriticalError
                cmbQuy.SetFocus
                Exit Sub
            End If
        End If
   End If
    
    ' Kiem tra to khai 08 tu ngay phai nho hon den ngay
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "74" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "75" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "23") And strQuy = "TK_TU_THANG" Then
            If Trim(txtNgayDau.Text) <> "" Then
                 strTempValue = Trim(txtNgayDau.Text)
                 dNgayDau = objCvt.ToDate("01" & "/" & strTempValue, "DD/MM/YYYY")
            End If
            
            If Trim(txtNgayCuoi.Text) <> "" Then
                 strTempValue = Trim(txtNgayCuoi.Text)
                 dNgayCuoi = objCvt.ToDate("01" & "/" & strTempValue, "DD/MM/YYYY")
            End If
            
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0069", msOKOnly, miInformation
                txtNgayCuoi.SetFocus
                Exit Sub
            End If
    End If
    
    ' 01/KK-TTS
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "23") And strQuy = "TK_TU_THANG" Then
            If Trim(txtNgayDau.Text) <> "" Then
                 strTempValue = Trim(txtNgayDau.Text)
                 dNgayDau = objCvt.ToDate("01" & "/" & strTempValue, "DD/MM/YYYY")
            End If
            
            If Trim(txtNgayCuoi.Text) <> "" Then
                 strTempValue = Trim(txtNgayCuoi.Text)
                 dNgayCuoi = objCvt.ToDate("01" & "/" & strTempValue, "DD/MM/YYYY")
            End If
            
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0069", msOKOnly, miInformation
                txtNgayCuoi.SetFocus
                Exit Sub
            End If
    End If
    
    ' 02/TNDN-DK
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "87" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "97" _
    Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "77" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "88" Then
            If Trim(txtNgayDau.Text) <> "" Then
                 strTempValue = Trim(txtNgayDau.Text)
                 dNgayDau = objCvt.ToDate("01" & "/" & strTempValue, "DD/MM/YYYY")
            End If
            
            If Trim(txtNgayCuoi.Text) <> "" Then
                 strTempValue = Trim(txtNgayCuoi.Text)
                 dNgayCuoi = objCvt.ToDate("01" & "/" & strTempValue, "DD/MM/YYYY")
            End If
            
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0069", msOKOnly, miInformation
                txtNgayCuoi.SetFocus
                Exit Sub
            End If
    End If
    
    'chan ngay kk lan phat sinh < ngay hien tai
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "70" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "06" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "72" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "81" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "90") And chkTKLanPS.value = "1" Then

        If Trim(sNgay) <> "" Then
            sNgayDD = DateSerial(CInt(Mid$(sNgay, 7, 4)), CInt(Mid$(sNgay, 4, 2)), CInt(Mid$(sNgay, 1, 2)))
            If DateDiff("D", Date, sNgayDD) > 0 Then
                DisplayMessage "0223", msOKOnly, miInformation
                Exit Sub
            End If
        End If
    End If
    '***************************
    
    Dim idxPL As Long
    Dim countPL As Integer
    ' BCTC kiem tra chi chon PL LCTTGT hoac LCTTTT
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "69" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "19" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "20" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "22" Then
        With fpSpread1
            .Col = 1
            For idxPL = 2 To .MaxRows
                .Row = idxPL
                If .value = 1 And (idxPL = 3 Or idxPL = 4) Then
                    countPL = countPL + 1
                End If
            Next idxPL
        End With
        
        If countPL = 2 Then
            If DisplayMessage("0314", msYesNo, miQuestion, , mrNo) = mrYes Then
            Else
                Exit Sub
            End If
        End If
    End If
    
    'set data
    TAX_Utilities_v1.Year = txtYear.Text
    If strKieuKy = KIEU_KY_THANG Then
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "01" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "02" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "04" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "71" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "36" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "25" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "96" _
        Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "94" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Then

            If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "71" Then
                If strQuy = "TK_THANG" Then
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.ThreeMonths = vbNullString
                    TAX_Utilities_v1.FirstDay = vbNullString
                    TAX_Utilities_v1.LastDay = vbNullString
                ElseIf strQuy = "TK_QUY" Then
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.ThreeMonths = cmbQuy.Text
                    TAX_Utilities_v1.FirstDay = vbNullString
                    TAX_Utilities_v1.LastDay = vbNullString
                ElseIf strQuy = "TK_LANPS" Then
                    TAX_Utilities_v1.ThreeMonths = "1"
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.Day = txtDay.Text
                    TAX_Utilities_v1.month = txtMonth.Text
                End If
            ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Then
                If strQuy = "TK_THANG" Then
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.ThreeMonths = vbNullString
                    TAX_Utilities_v1.FirstDay = vbNullString
                    TAX_Utilities_v1.LastDay = vbNullString
                ElseIf strQuy = "TK_LANXB" Then
                    TAX_Utilities_v1.ThreeMonths = "1"
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.Day = txtDay.Text
                    TAX_Utilities_v1.month = txtMonth.Text
                ElseIf strQuy = "TK_LANPS" Then
                    TAX_Utilities_v1.ThreeMonths = "1"
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.Day = txtDay.Text
                    TAX_Utilities_v1.month = txtMonth.Text
                End If
            Else
                If strQuy = "TK_THANG" Then
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.ThreeMonths = vbNullString
                    TAX_Utilities_v1.FirstDay = vbNullString
                    TAX_Utilities_v1.LastDay = vbNullString
                ElseIf strQuy = "TK_QUY" Then
                    TAX_Utilities_v1.month = txtMonth.Text
                    TAX_Utilities_v1.ThreeMonths = cmbQuy.Text
                    TAX_Utilities_v1.FirstDay = vbNullString
                    TAX_Utilities_v1.LastDay = vbNullString
                End If
            End If
        Else
            TAX_Utilities_v1.month = txtMonth.Text
            TAX_Utilities_v1.ThreeMonths = vbNullString
            TAX_Utilities_v1.FirstDay = vbNullString
            TAX_Utilities_v1.LastDay = vbNullString
        End If
    ElseIf strKieuKy = KIEU_KY_QUY Then
        TAX_Utilities_v1.month = vbNullString
        TAX_Utilities_v1.ThreeMonths = cmbQuy.Text
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "74" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "75" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "23" Then
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
        Else
            TAX_Utilities_v1.FirstDay = vbNullString
            TAX_Utilities_v1.LastDay = vbNullString
        End If
        ' To khai 02/TNDN
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Then
            If strLoaiTKThang_PS = "TK_LANPS" Then
                TAX_Utilities_v1.Day = txtDay.Text
                TAX_Utilities_v1.month = txtMonth.Text
            Else
                TAX_Utilities_v1.Day = vbNullString
                TAX_Utilities_v1.month = vbNullString
            End If
        End If
' phuc vu an chi
' dhdang comment to khai nao???
    ElseIf strKieuKy = "H_Y" Then
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "68" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "14" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "13" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "65" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "18" Then
            If strQuy = "TK_THANG" Then
                TAX_Utilities_v1.month = txtMonth.Text
            Else
                TAX_Utilities_v1.month = vbNullString
            End If
            TAX_Utilities_v1.ThreeMonths = cmbQuy.Text
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
            TAX_Utilities_v1.Year = txtYear.Text
        Else
            TAX_Utilities_v1.month = vbNullString
            TAX_Utilities_v1.ThreeMonths = cmbQuy.Text
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
            TAX_Utilities_v1.Year = txtYear.Text
        End If
 ' end
    ElseIf strKieuKy = KIEU_KY_NGAY_NAM Then
        TAX_Utilities_v1.month = vbNullString
        TAX_Utilities_v1.ThreeMonths = vbNullString
        TAX_Utilities_v1.FirstDay = txtNgayDau.Text
        TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
    ElseIf strKieuKy = KIEU_KY_NAM Then
        TAX_Utilities_v1.month = vbNullString
        TAX_Utilities_v1.ThreeMonths = vbNullString
        TAX_Utilities_v1.FirstDay = vbNullString
        TAX_Utilities_v1.LastDay = vbNullString
        If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "87" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "97" _
        Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "77" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "88" Then
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
        ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "76" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "59" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "43" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "41" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "17" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "26" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "45" Then
            ' QT TNCN
            TAX_Utilities_v1.FirstDay = txtNgayDau.Text
            TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
        Else
            TAX_Utilities_v1.FirstDay = vbNullString
            TAX_Utilities_v1.LastDay = vbNullString
        End If
  ' htphuong add them to khai 05/GTGT
    ElseIf strKieuKy = KIEU_KY_NGAY_THANG Then
        TAX_Utilities_v1.month = txtMonth.Text
        TAX_Utilities_v1.Day = txtDay.Text
        TAX_Utilities_v1.ThreeMonths = vbNullString
        TAX_Utilities_v1.FirstDay = vbNullString
        TAX_Utilities_v1.LastDay = vbNullString
    ElseIf strKieuKy = KIEU_KY_NGAY_PS Then
        TAX_Utilities_v1.Day = txtDay.Text
        TAX_Utilities_v1.month = txtMonth.Text
        TAX_Utilities_v1.Year = txtYear.Text
        TAX_Utilities_v1.ThreeMonths = vbNullString
        TAX_Utilities_v1.FirstDay = vbNullString
        TAX_Utilities_v1.LastDay = vbNullString
    End If
    
    ' Luu gia tri nganh nghe kinh doanh
    idToKhai = TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    If idToKhai = "01" Or idToKhai = "11" Or idToKhai = "12" Or idToKhai = "05" Or idToKhai = "03" Then
    'If idToKhai = "01" Or idToKhai = "11" Or idToKhai = "12" Or idToKhai = "05" Or idToKhai = "03" Or idToKhai = "73" Then
        strLoaiNNKD = cboNganhKD.ItemData(cboNganhKD.ListIndex)
    End If
    
    If idToKhai = "98" Or idToKhai = "92" Or idToKhai = "93" Or idToKhai = "89" Then
       strLoaiNNKD = cboNganhKD.ItemData(cboNganhKD.ListIndex)
        If strLoaiNNKD = 1 Then
            strLoaiTkDk = "DT"
        ElseIf strLoaiNNKD = 2 Then
            strLoaiTkDk = "KTN"
        ElseIf strLoaiNNKD = 3 Then
            strLoaiTkDk = "CD"
        End If
    End If
    
    If TAX_Utilities_v1.NodeMenu Is Nothing Then Exit Sub
    
    If TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "15" Then
            If ExistTokhai("02B_TNCN10_", False, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0126", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "16" Then
            If ExistTokhai("02A_TNCN10_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
    ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "53" Then
            If ExistTokhai("02B_TNCN_", True, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
    ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "37" Then
            If ExistTokhai("02A_TNCN_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "50" Then
            If ExistTokhai("03B_TNCN10_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0126", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "51" Then
            If ExistTokhai("03A_TNCN10_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "54" Then
            If ExistTokhai("03B_TNCN_", True, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "38" Then
            If ExistTokhai("03A_TNCN_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "39" Then
            If ExistTokhai("04B_TNCN_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0126", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "40" Then
            If ExistTokhai("04A_TNCN_", True, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
    ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "46" Then
            If ExistTokhai("01B_TNCN_BH_", False, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0126", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
     ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "47" Then
            If ExistTokhai("01A_TNCN_BH_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
    ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "48" Then

        If ExistTokhai("01B_TNCN_XS_", False, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
            DisplayMessage "0126", msOKOnly, miWarning
            Unload Me
            frmTreeviewMenu.Show
            Exit Sub
        End If

    ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "49" Then

        If ExistTokhai("01A_TNCN_XS_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
            DisplayMessage "0125", msOKOnly, miWarning
            Unload Me
            frmTreeviewMenu.Show
            Exit Sub
        End If

        ' xu ly cho to khai thang/ quy
    ElseIf TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "01" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "02" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "04" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "71" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "25" Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "96" _
    Or TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "94" Then

        If strQuy = "TK_QUY" Then
            If ExistTokhai(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_", True, TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0125", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
                 End If
            ElseIf strQuy = "TK_THANG" Then
                If ExistTokhai(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_Q", False, TAX_Utilities_v1.month & TAX_Utilities_v1.Year) = True Then
                DisplayMessage "0126", msOKOnly, miWarning
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            End If
            End If
    End If
    
    ' Check validate BC AC
    ' BC26
    ' Kiem tra tu ngay
    If idToKhai = "68" Or idToKhai = "14" Or idToKhai = "13" Or idToKhai = "65" Or idToKhai = "18" Then
        If strQuy = "TK_THANG" Then
            dNgayDau = DateSerial(CInt(Mid$(TAX_Utilities_v1.FirstDay, 7, 4)), CInt(Mid$(TAX_Utilities_v1.FirstDay, 4, 2)), CInt(Mid$(TAX_Utilities_v1.FirstDay, 1, 2)))
            dNgayCuoi = DateSerial(CInt(Mid$(TAX_Utilities_v1.LastDay, 7, 4)), CInt(Mid$(TAX_Utilities_v1.LastDay, 4, 2)), CInt(Mid$(TAX_Utilities_v1.LastDay, 1, 2)))

            dNgayDauQuy = DateSerial(CInt(TAX_Utilities_v1.Year), CInt(TAX_Utilities_v1.month), 1)
            Dim temp As Integer
            Dim temp1 As Date
            temp = CInt(TAX_Utilities_v1.month) + 1
            If TAX_Utilities_v1.month = "12" Then
                temp1 = DateSerial(CInt(TAX_Utilities_v1.Year) + 1, 1, 1)
                dNgayCuoiQuy = DateAdd("D", -1, temp1)
            Else
                temp1 = DateSerial(CInt(TAX_Utilities_v1.Year), temp, 1)
                dNgayCuoiQuy = DateAdd("D", -1, temp1)
            End If

            ' Ky bao cao tu ngay khong duoc lon hon ky bao cao den ngay
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0254", msOKOnly, miWarning
                Exit Sub
            End If
            ' Ky bao cao den ngay khong duoc lon hon ngay cuoi quy
            If dNgayCuoi > dNgayCuoiQuy Then
                DisplayMessage "0319", msOKOnly, miWarning
                Exit Sub
            End If
            ' Ky bao cao tu ngay khong duoc nho hon ngay dau quy
            If dNgayDau < dNgayDauQuy Then
                DisplayMessage "0318", msOKOnly, miWarning
                txtNgayDau.SetFocus
                Exit Sub
            End If
        ElseIf strQuy = "TK_KY" Then
            dNgayDau = DateSerial(CInt(Mid$(TAX_Utilities_v1.FirstDay, 7, 4)), CInt(Mid$(TAX_Utilities_v1.FirstDay, 4, 2)), CInt(Mid$(TAX_Utilities_v1.FirstDay, 1, 2)))
            dNgayCuoi = DateSerial(CInt(Mid$(TAX_Utilities_v1.LastDay, 7, 4)), CInt(Mid$(TAX_Utilities_v1.LastDay, 4, 2)), CInt(Mid$(TAX_Utilities_v1.LastDay, 1, 2)))
            
            
'            dNgayDauQuy = GetNgayDauQuy(CInt(TAX_Utilities_v1.ThreeMonths), TAX_Utilities_v1.Year, 1, 1)
'            dNgayCuoiQuy = GetNgayCuoiK(CInt(TAX_Utilities_v1.ThreeMonths), TAX_Utilities_v1.Year, 1, 1)
            ' Ky bao cao tu ngay khong duoc lon hon ky bao cao den ngay
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0254", msOKOnly, miWarning
                Exit Sub
            End If
'            ' Ky bao cao den ngay khong duoc lon hon ngay cuoi quy
'            If dNgayCuoi > dNgayCuoiQuy Then
'                DisplayMessage "0255", msOKOnly, miWarning
'                Exit Sub
'            End If
'            ' Ky bao cao tu ngay khong duoc nho hon ngay dau quy
'            If dNgayDau < dNgayDauQuy Then
'                DisplayMessage "0256", msOKOnly, miWarning
'                txtNgayDau.SetFocus
'                Exit Sub
'            End If
'            ' Kiem tra ngay dau quy khong dc nho hon ngay 01/01/2011
'            If dNgayCuoi > DateSerial(2014, 6, 30) Then
'                DisplayMessage "0317", msOKOnly, miWarning
'                txtNgayDau.SetFocus
'                Exit Sub
'            End If
        Else
            dNgayDau = DateSerial(CInt(Mid$(TAX_Utilities_v1.FirstDay, 7, 4)), CInt(Mid$(TAX_Utilities_v1.FirstDay, 4, 2)), CInt(Mid$(TAX_Utilities_v1.FirstDay, 1, 2)))
            dNgayCuoi = DateSerial(CInt(Mid$(TAX_Utilities_v1.LastDay, 7, 4)), CInt(Mid$(TAX_Utilities_v1.LastDay, 4, 2)), CInt(Mid$(TAX_Utilities_v1.LastDay, 1, 2)))
            dNgayDauQuy = GetNgayDauQuy(CInt(TAX_Utilities_v1.ThreeMonths), TAX_Utilities_v1.Year, 1, 1)
            dNgayCuoiQuy = GetNgayCuoiQuy(CInt(TAX_Utilities_v1.ThreeMonths), TAX_Utilities_v1.Year, 1, 1)
            ' Ky bao cao tu ngay khong duoc lon hon ky bao cao den ngay
            If dNgayCuoi < dNgayDau Then
                DisplayMessage "0254", msOKOnly, miWarning
                Exit Sub
            End If
            ' Ky bao cao den ngay khong duoc lon hon ngay cuoi quy
            If dNgayCuoi > dNgayCuoiQuy Then
                DisplayMessage "0255", msOKOnly, miWarning
                Exit Sub
            End If
            ' Ky bao cao tu ngay khong duoc nho hon ngay dau quy
            If dNgayDau < dNgayDauQuy Then
                DisplayMessage "0256", msOKOnly, miWarning
                txtNgayDau.SetFocus
                Exit Sub
            End If
            ' Kiem tra ngay dau quy khong dc nho hon ngay 01/01/2011
            If dNgayDau < DateSerial(2011, 1, 1) Then
                DisplayMessage "0257", msOKOnly, miWarning
                txtNgayDau.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    ' check bao cao nhan in
    ' BC01/AC
    If idToKhai = "65" Then
        If strQuy = "TK_QUY" Then
            If Val(txtYear) < 2014 Or (Val(txtYear) = 2014 And Val(cmbQuy.Text) < 3) Then
                DisplayMessage "0316", msOKOnly, miWarning
                cmbQuy.SetFocus
                Exit Sub
            End If
        Else
            If Val(txtYear) > 2014 Or (Val(txtYear) = 2014 And Val(cmbQuy.Text) >= 2) Then
                DisplayMessage "0315", msOKOnly, miWarning
                cmbQuy.SetFocus
                Exit Sub
            End If

        End If
    End If
    ' end
    
    If idToKhai = "71" Then
        If chkTkhaiThang.value = 0 And chkTKLanPS.value = 0 And chkTKQuy.value = 0 Then
            DisplayMessage "0295", msOKOnly, miWarning
            chkTkhaiThang.SetFocus
            Exit Sub
        End If
        
        ' kiem tra ngay ps khong duoc lon hon ngay hien tai
        If strQuy = "TK_LANPS" Then
            If DateDiff("D", Date, DateSerial(Val(txtYear.Text), Val(txtMonth.Text), Val(txtDay.Text))) > 0 Then
                DisplayMessage "0223", msOKOnly, miWarning
                txtDay.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    ' chan to khai 01A/TNDN, 01B/TNDN ky ke khai quy 4/2014 tro di
    If idToKhai = "11" Or idToKhai = "12" Then
        If (Val(TAX_Utilities_v1.ThreeMonths) >= 4 And Val(TAX_Utilities_v1.Year) = 2014) Or Val(TAX_Utilities_v1.Year) > 2014 Then
            DisplayMessage "0341", msOKOnly, miWarning
            cmbQuy.SetFocus
            Exit Sub
        End If
    End If
    ' chan to khai bs 02/TNDN quy
    If idToKhai = "73" Then
        If strLoaiTKThang_PS = "TK_THANG" Then
            If ((Val(TAX_Utilities_v1.ThreeMonths) >= 4 And Val(TAX_Utilities_v1.Year) = 2014) Or Val(TAX_Utilities_v1.Year) > 2014) Or strKHBS = "TKBS" Then
                DisplayMessage "0341", msOKOnly, miWarning
                cmbQuy.SetFocus
                Exit Sub
            End If
        End If
    End If
    ' kiem tra trung khoang doi voi to khai QT co ky bo sung tu thang den thang
    If strKHBS = "TKCT" And strKieuKy = KIEU_KY_NAM Then
        If idToKhai = "93" Or idToKhai = "89" Then
            If checkKyKKTrung(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk, Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text), txtYear.Text) = True Then
                DisplayMessage "0340", msOKOnly, miWarning
                Exit Sub
            End If
        ' to khai 09/TNCN khong check
        ElseIf idToKhai = "41" Or idToKhai = "76" Then
        Else
            If checkKyKKTrung(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile"), Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text), txtYear.Text) = True Then
                DisplayMessage "0340", msOKOnly, miWarning
                Exit Sub
            End If
        End If
    ElseIf strKHBS = "TKBS" And strKieuKy = KIEU_KY_NAM Then
        If idToKhai = "93" Or idToKhai = "89" Then
            If checkKyKKTrung("bs" & Trim$(txtSolan.Text) & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk, Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text), txtYear.Text) = True Then
                DisplayMessage "0340", msOKOnly, miWarning
                Exit Sub
            End If
        ' to khai 09/TNCN khong check
        ElseIf idToKhai = "41" Or idToKhai = "76" Then
        Else
            If checkKyKKTrung("bs" & Trim$(txtSolan.Text) & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile"), Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text), txtYear.Text) = True Then
                DisplayMessage "0340", msOKOnly, miWarning
                Exit Sub
            End If
        End If
    End If
    
    ' check to khai 03/TNDN
    If strKHBS = "TKCT" Then
        If idToKhai = "03" Then
            If checkKyKKTrungNgay(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile"), Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text), Trim$(txtYear.Text)) = True Then
                DisplayMessage "0344", msOKOnly, miWarning
                Exit Sub
            End If
        End If
    ElseIf strKHBS = "TKBS" Then
        If idToKhai = "03" Then
            If checkKyKKTrungNgay("bs" & Trim$(txtSolan.Text) & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile"), Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text), Trim$(txtYear.Text)) = True Then
                DisplayMessage "0344", msOKOnly, miWarning
                Exit Sub
            End If
        End If
    End If
    
    ' kiem tra tu ngay den ngay cho 2 to khai NTNN
    
'    If strKHBS = "TKCT" Then
'        If idToKhai = "80" Or idToKhai = "82" Then
'            If checkKyKKTrungNgay(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile"), Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text)) = True Then
'                DisplayMessage "0344", msOKOnly, miWarning
'                Exit Sub
'            End If
'        End If
'    ElseIf strKHBS = "TKBS" Then
'        If idToKhai = "80" Or idToKhai = "82" Then
'            If checkKyKKTrungNgay("bs" & Trim$(txtSolan.Text) & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile"), Trim(txtNgayDau.Text), Trim(txtNgayCuoi.Text)) = True Then
'                DisplayMessage "0344", msOKOnly, miWarning
'                Exit Sub
'            End If
'        End If
'    End If
    
    ' To khai bo sung
    'If strKHBS = "TKBS" And (idToKhai = "02" Or idToKhai = "01" Or idToKhai = "04" Or idToKhai = "03" Or idToKhai = "11" Or idToKhai = "12" Or idToKhai = "06" Or idToKhai = "05" Or idToKhai = "86" Or idToKhai = "87" Or idToKhai = "71" Or idToKhai = "72" Or idToKhai = "77" Or idToKhai = "73" Or idToKhai = "80" Or idToKhai = "81" Or idToKhai = "70" Or idToKhai = "82" Or idToKhai = "83" Or idToKhai = "85" Or idToKhai = "90" Or idToKhai = "92" Or idToKhai = "93" Or idToKhai = "95" Or idToKhai = "88" Or idToKhai = "94" Or idToKhai = "96" Or idToKhai = "97" Or idToKhai = "98" Or idToKhai = "92" Or idToKhai = "99" Or idToKhai = "26") Then
    If strKHBS = "TKBS" And (InStr(1, strIdKHBS_TT156, "~" & Trim$(idToKhai) & "~", vbTextCompare) > 0 Or idToKhai = "01") Then
        'dhdang them lay ngay KHBS
        'kiem tra ton tai TK chinh thuc
        Dim strDay As Variant
        Dim strDataFileBS As Variant
        Dim fso1 As New FileSystemObject
        If Trim(TAX_Utilities_v1.month) <> "" Then
            ' to khai nha thau nuoc ngoai
            If idToKhai = "70" Or idToKhai = "06" Or idToKhai = "81" Or idToKhai = "90" Or idToKhai = "05" Then
                If strLoaiTKThang_PS = "TK_THANG" Then
                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                End If
            ElseIf idToKhai = "72" Then
                If strLoaiTKThang_PS = "TK_THANG" Then
                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                End If
            ElseIf idToKhai = "73" Then
                If strLoaiTKThang_PS = "TK_LANPS" Then
                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                Else
                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
                End If

            ElseIf idToKhai = "01" Or idToKhai = "02" Or idToKhai = "04" Or idToKhai = "71" Or idToKhai = "96" Or idToKhai = "94" Then

                If idToKhai = "71" Then
                    If strQuy = "TK_THANG" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
                    ElseIf strQuy = "TK_LANPS" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    End If

                Else
                    If strQuy = "TK_THANG" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
                    End If
                End If

                'ElseIf idToKhai = "98" Or idToKhai = "92" Or idToKhai = "93" Or idToKhai = "99" Then
            ElseIf idToKhai = "98" Or idToKhai = "92" Then

'                If strLoaiTKThang_PS = "TK_THANG" Then
'                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
'                ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                    strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
'                End If
                 If strQuy = "TK_THANG" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                 ElseIf strQuy = "TK_LANPS" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                 ElseIf strQuy = "TK_LANXB" Then
                        strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                 End If
            Else
                strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
            End If
        ElseIf Trim(TAX_Utilities_v1.ThreeMonths) <> "" Then
            strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
        ElseIf Trim(TAX_Utilities_v1.Year) <> "" Then
'            If idToKhai = "77" Or idToKhai = "88" Or idToKhai = "87" Then
'                strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & ".xml"
'            Else
            If idToKhai = "80" Or idToKhai = "82" Then
                strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml"
            ElseIf idToKhai = "93" Or idToKhai = "89" Then
                strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Year & "_" & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml"
            Else
                strDataFileBS = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & "_" & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml"
            End If
        End If
        If Not fso1.FileExists(strDataFileBS) Then
            DisplayMessage "0107", msOKOnly, miInformation
            Exit Sub
        End If

        If Val(txtSolan.Text) > 1 Then
            Dim strTK_BS As String
            Dim strBsSplit() As String
            strBsSplit = Split(strDataFileBS, "\")
            strTK_BS = Replace$(strDataFileBS, strBsSplit(UBound(strBsSplit)), "bs" & CStr(Val(txtSolan.Text) - 1) & "_" & strBsSplit(UBound(strBsSplit)))

            If Not fso1.FileExists(strTK_BS) Then
                DisplayMessage "0296", msOKOnly, miInformation
                Exit Sub
            End If
        End If

        With fpsNgaykhaiBS
        .Col = .ColLetterToNumber("C")
        .Row = 2
        strDateKHBS = .Text
        If strDateKHBS <> "" And strDateKHBS <> "../../...." Then
            If Format_ddmmyyyy(CStr(strDateKHBS)) <> "" Then
                .SetText .ColLetterToNumber("C"), 2, Format_ddmmyyyy(CStr(strDateKHBS))
                .TypeHAlign = TypeHAlignLeft
            Else
                .SetFocus
                .SetActiveCell .ColLetterToNumber("C"), 2
                Exit Sub
            End If
        Else
         .SetText .ColLetterToNumber("C"), 2, format(Date, "dd/mm/yyyy")
         Exit Sub
        End If
        
        
        ' kiem tra voi ngay hien tai
        Dim arrDate() As String
        Dim hn As Date
        Dim ngayBs As Date
        Dim ngayHt As Date
        
        arrDate = Split(strDateKHBS, "/")
        ngayBs = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
        ngayHt = DateSerial(Year(Date), month(Date), Day(Date))
        
        If DateDiff("D", ngayHt, ngayBs) > 0 Then
            DisplayMessage "0224", msOKOnly, miInformation
            Exit Sub
        End If
        
        'kiem tra voi ky kk
        hanNopTk = GetHanNopTk
        arrDate = Split(hanNopTk, "/")
        hn = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
        arrDate = Split(strDateKHBS, "/")
        ngayBs = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
        
        Dim hnps As Date
        If strLoaiTKThang_PS = "TK_LANPS" Then
            If idToKhai = "98" Or idToKhai = "92" Then
                hnps = DateAdd("D", 35, DateSerial(CInt(TAX_Utilities_v1.Year), CInt(TAX_Utilities_v1.month), CInt(TAX_Utilities_v1.Day)))
            Else
                hnps = DateAdd("D", 10, DateSerial(CInt(TAX_Utilities_v1.Year), CInt(TAX_Utilities_v1.month), CInt(TAX_Utilities_v1.Day)))
            End If
            If DateDiff("D", hnps, ngayBs) <= 0 Then
                DisplayMessage "0271", msOKOnly, miInformation
                Exit Sub
            End If
        ElseIf strQuy = "TK_LANPS" And idToKhai = "71" Then
                hnps = DateAdd("D", 10, DateSerial(CInt(TAX_Utilities_v1.Year), CInt(TAX_Utilities_v1.month), CInt(TAX_Utilities_v1.Day)))
            If DateDiff("D", hnps, ngayBs) <= 0 Then
                DisplayMessage "0271", msOKOnly, miInformation
                Exit Sub
            End If
        Else
            If DateDiff("D", hn, ngayBs) < 0 Then
                DisplayMessage "0271", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ngayLapTkBs = strDateKHBS
        
        End With
        
        If strDateKHBS <> vbNullString Then
            TAX_Utilities_v1.DateKHBS = Replace(strDateKHBS, "/", "")
        End If
        SetActiveValueKHBS
    Else
        '***************************
        ' added
        SetActiveValue
    End If
    
    ' Set phuluc 26,27MT-TNCN
    ' chi ky ke khai nam 2012 bo sung them PL
    If idToKhai = "41" Then
        If TAX_Utilities_v1.Year = "2012" Then
            SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(4), "Active", "1"
        End If
    End If
    
    If idToKhai = "17" Then
        If TAX_Utilities_v1.Year = "2012" Then
            SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(4), "Active", "1"
        End If
    End If
    
    '***************************
    'show form
    
'    Debug.Print "Bat dau load" & Time
    'frmInterfaces.LblaodTK.Visible = True
    Set frmTK = New frmInterfaces
    Unload Me
    frmTK.Show
    frmSystem.Hide
    
'    Debug.Print "Ket thuc load" & Time
    'frmInterfaces.LblaodTK.Visible = False
    
    ' Doi voi cac to khai TNCN noi chung thi An nut xoa va nut Insert di
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        frmTK.cmdLoadToKhai.Visible = True
        frmTK.cmdInsert.Visible = False
        'frmTK.cmdDelete.Visible = False
        frmTK.cmdDelete.Left = frmTK.Frame1.Width - 12100
        frmTK.cmdKiemTra.Visible = True
    End If
    ' Neu la cac mau in tong hop tu to quyet toan 05TNCN->09TNCN va cac chung tu cua TNCN thi an cac nut di, chi de In va Thoat
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "45" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "52" Then
                
        frmTK.cmdLoadToKhai.Visible = False
        frmTK.cmdClear.Visible = False
        frmTK.cmdDelete.Visible = True
        frmTK.cmdDelete.Visible = False
        frmTK.cmdExport.Visible = False
        frmTK.cmdSave.Visible = False
        frmTK.cmdInsert.Visible = False
        frmTK.cmdKiemTra.Visible = False
        frmTK.cmdPrint.Left = frmTK.Frame1.Width - 2420
        frmTK.cmdExit.Left = frmTK.Frame1.Width - 1240
        
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
    
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnClick = True
End Sub

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
'****************************************************
'Description:Form_Load procedure load form Period
'   Step 1: Read information from node menu to show, hide controls
'   Step 2: Setup layout
'   Step 3: load default information
'****************************************************

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
'    Dim fso As New FileSystemObject
    
'    If fso.FileExists("..\InterfaceTemplates\Template.xls") Then
'        If fpSpread1.IsExcelFile("..\InterfaceTemplates\Template.xls") Then
'            fpSpread1.EventEnabled(EventAllEvents) = False
'            fpSpread1.ImportExcelBook GetAbsolutePath("..\InterfaceTemplates\Template.xls"), vbNullString
'            fpSpread1.EventEnabled(EventAllEvents) = True
'        End If
'    End If
    
'    Set fso = Nothing
    ' reset cac tham so
    strLoaiTKThang_PS = ""
    strLoaiTkDk = ""
    ' end
    
    fpSpread1.Reset 'Can be removed
    fpSpread1.FontName = "Tahoma"
    
    'initInforLayout
    hasActiveForm = True
    
    'Lay kieu ky ke khai
    strKieuKy = GetKieuKy
    
    'Lay ngay bat dau nam tai chinh
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "FinanceYear") = "1" Then
        strNgayTaiChinh = GetNgayBatDauNamTaiChinh
        'strNgayTaiChinh = "01/01" 'for testing
'        If Not KiemTraNgayTaiChinh(strNgayTaiChinh) Then
'            Unload Me
'            frmTreeviewMenu.Show
'            Exit Sub
'        Else
            iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
            iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
'        End If
    Else
        strNgayTaiChinh = "01/01"
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
    End If
    
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "15" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "16" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "37" _
            Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "38" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "39" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "40" _
                 Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "42" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "44" _
                    Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "46" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "47" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "48" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "49" _
                        Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "50" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "51" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "53" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "54" _
                        Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "74" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "75" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "24" Then
        SetupLayoutTNCN (strKieuKy)
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "76" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "59" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "43" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "41" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "17" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "26" Then
        SetupLayoutTNCN_QT
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "45" Then
        SetupLayoutTNCN_HT_IN_QT
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "99" Then
        SetupLayoutTNDN_DK (strKieuKy)
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "02" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "04" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "36" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "25" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "96" _
    Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "94" Then
        SetLayoutToKhaiThangQuy
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "71" Then
        SetLayoutToKhaiThangQuyLanPS
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "11" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "12" _
    Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "86" _
     Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "83" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "85" Then
        SetupLayoutGTGT strKieuKy, GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "70" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "72" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "06" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "81" Then
        SetupLayoutNTNN
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "05" Then
        SetupLayoutTTDB
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Then
        SetupLayout02TNDN
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "90" Then
        SetupLayout01TBVMT
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "03" Then
        SetupLayout03TNDN
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "88" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "97" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "77" Then
        SetupLayout02BVMT
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "80" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "82" Then
        SetupLayout02NTNN
'    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "74" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "75" Then
'        SetupLayout08TNCN
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "91" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "64" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "07" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "27" Then
        SetupLayout04TBAC
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Then
        SetupLayout01_TAIN_DK
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "89" Then
        SetupLayout02_TNDN_DK
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "23" Then
        SetupLayout01TTS
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Then
        SetupLayoutBC26
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "65" Then
        SetupLayoutBC01
    ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "95" Then
        SetupLayout16TH
    Else
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "18" Then
            strQuy = "TK_QUY"
        End If
        SetupLayout (strKieuKy)
    End If
    LoadDefaultInfor

    bIsClosed = False
    
    '********************
    ' added
    LoadGrid
    
    ' xu ly cho to khai DK
    Dim m, Y, d As Integer
    Dim dTem, dtem1, dtem2 As Date
    Dim varMenuId As String
    dtem2 = Date
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Then
         m = month(dtem2)
        Y = Year(dtem2)
        d = Day(dtem2)
        txtDay.Text = d
        txtMonth.Text = m
        txtYear.Text = Y
        If Len(txtDay.Text) = 1 Then
            txtDay.Text = "0" & txtDay.Text
        End If
        If Len(txtMonth.Text) = 1 Then
            txtMonth.Text = "0" & txtMonth.Text
        End If
    End If
    
    ' Cac to quyet toan TNCN kiem tra xem de dat lai nut Dong y, Dong cho dung, dat sau Frame 2
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        Frame2.Visible = False
        lblSelectAll.Visible = False
        chkSelectAll.Visible = False
        fpSpread1.Visible = False
        Call Form_Resize
    End If
    
    Me.Top = Me.Top - 500
    '********************
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
    
End Sub

'****************************************************
'Description:SetupLayout procedure setup layout
'****************************************************

Private Sub SetupLayout(strKieuKy As String)
    On Error GoTo ErrorHandle
    
    Me.Height = 3285
    Me.Width = 4905
    
    Select Case strKieuKy
        Case KIEU_KY_THANG
            Set lblMonth.Container = frmKy
            lblMonth.Top = 480
            lblMonth.Left = 960
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 450
            txtMonth.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 450
            txtYear.Left = 2730
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
        
        Case KIEU_KY_QUY
            Set lblQuy.Container = frmKy
            lblQuy.Top = 480
            lblQuy.Left = 1050
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 450
            cmbQuy.Left = 1440
            
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = 2220
            
            Set txtYear.Container = frmKy
            txtYear.Top = 450
            txtYear.Left = 2640
            
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
                SetControlCaption Me, "frmPeriodBCTC"
            Else
                SetControlCaption Me, "frmPeriodQuy"
            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            cmdOK.Top = 1500
            cmdClose.Top = cmdOK.Top
            Me.Height = 2040
            Me.Width = 4905
        'dhdang sua  them kieu ky nua nam phuc vu an chi
        Case "H_Y"
            Set lblQuy.Container = frmKy
            lblQuy.Top = 300
            lblQuy.Left = 120
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "65" Then
                lblQuy.caption = "Ky`"
            End If
            SetControlCaption Me, "frmPeriodHY"
            'lblCaption.caption = GetAttribute(GetMessageCellById("0183"), "Msg")
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 240
            cmbQuy.Left = 1000
            
'            lblYear.Top = 300
'            lblYear.Left = 1000
            lblYear.Visible = False
            
            
            Set txtYear.Container = frmKy
            txtYear.Top = 240
            txtYear.Left = 1600
            
            Set lblNgayDau.Container = frmKy
            lblNgayDau.Top = 630
            lblNgayDau.Left = 120
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 600
            txtNgayDau.Left = 1000 '1200
            'txtNgayDau.Locked = True
            
            Set lblNgayCuoi.Container = frmKy
            lblNgayCuoi.Top = 630
            lblNgayCuoi.Left = 2600 '2400
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 600
            txtNgayCuoi.Left = 3480
            
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            'cmbQuy.Visible = False
             ' end
        Case KIEU_KY_NAM
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = (frmKy.Width - txtYear.Width) / 2 - 100
            
            Set txtYear.Container = frmKy
            txtYear.Top = 420
            txtYear.Left = lblYear.Left + lblYear.Width + 50 '1200
            
            
            
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
                SetControlCaption Me, "frmPeriodBCTC"
            Else
                SetControlCaption Me, "frmPeriodQuy"
            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            txtMonth.Visible = False
            cmbQuy.Visible = False
        Case KIEU_KY_NGAY_NAM
            Set lblYear.Container = frmKy
            lblYear.Top = 300
            lblYear.Left = 120
            
            Set txtYear.Container = frmKy
            txtYear.Top = 240
            txtYear.Left = 1000 '1200
            
            Set lblNgayDau.Container = frmKy
            lblNgayDau.Top = 630
            lblNgayDau.Left = 120
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 600
            txtNgayDau.Left = 1000 '1200
            
            Set lblNgayCuoi.Container = frmKy
            lblNgayCuoi.Top = 630
            lblNgayCuoi.Left = 2600 '2400
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 600
            txtNgayCuoi.Left = 3480
            
            SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            cmbQuy.Visible = False
'htphuong add for TK 05/GTGT
        Case KIEU_KY_NGAY_THANG
            Set lblNgay.Container = frmKy
            lblNgay.Top = 480
            lblNgay.Left = 240
            
            Set txtDay.Container = frmKy
            txtDay.Top = 450
            txtDay.Left = 840
            
            Set lblMonth.Container = frmKy
            lblMonth.Top = 480
            lblMonth.Left = 1560
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 450
            txtMonth.Left = 2160
            
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = 2880
            
            Set txtYear.Container = frmKy
            txtYear.Top = 450
            txtYear.Left = 3360
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            
    End Select
'    If intMonth = 0 Then
'        Me.Width = 3900 '3840
'        Me.Height = 1600
'        Frame1.Height = 720
        
'       lblYear.Top = 300
'        lblYearFormat.Top = 300
        
'        txtYear.Top = 255
        
'        cmdOk.Top = 1100
'        cmdClose.Top = 1100
'    Else
'        Me.Width = 3900 '3840
'        Me.Height = 1920
'        Frame1.Height = 1065
    
'        lblMonth.Top = 300
'        lblYear.Top = 615
'        lblYearFormat.Top = 615
        
'        txtMonth.Top = 255
'        txtYear.Top = 585
        
'        cmdOk.Top = 1410
'        cmdClose.Top = 1410
'    End If
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout", Err.Number, Err.Description
    
End Sub

Private Sub SetupLayoutTNCN(strKieuKy As String)
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1365
    Select Case strKieuKy
        Case KIEU_KY_THANG
            Set lblMonth.Container = frmKy
            lblMonth.Top = 200
            lblMonth.Left = 960
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 150
            txtMonth.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = 2730
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 600
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 950
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 950
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
        
        Case KIEU_KY_QUY
            Set lblQuy.Container = frmKy
            lblQuy.Top = 200
            lblQuy.Left = 1050
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 150
            cmbQuy.Left = 1440
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = 2220
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = 2640
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 650
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 900
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 850
            txtSolan.Left = 3400
            lblSolan.Visible = False
            txtSolan.Visible = False
                        
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
                SetControlCaption Me, "frmPeriodBCTC"
            Else
                SetControlCaption Me, "frmPeriodQuy"
            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            cmdOK.Top = 1500
            cmdClose.Top = cmdOK.Top
            Me.Height = 2040
            Me.Width = 4905
            
        Case KIEU_KY_NAM
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = (frmKy.Width - txtYear.Width) / 2 - 100
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = lblYear.Left + lblYear.Width + 50
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 650
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 900
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 850
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False

            
'            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
'                SetControlCaption Me, "frmPeriodBCTC"
'            Else
                SetControlCaption Me, "frmPeriodQuy"
'            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            txtMonth.Visible = False
            cmbQuy.Visible = False
            
        Case KIEU_KY_NGAY_NAM
            Me.Height = 3285
            Me.Width = 5305
            frmKy.Height = 1665
            Frame2.Top = 1815
    
            Set lblYear.Container = frmKy
            lblYear.Top = 300
            lblYear.Left = 120
            
            Set txtYear.Container = frmKy
            txtYear.Top = 240
            txtYear.Left = 1000 '1200
            
            Set lblNgayDau.Container = frmKy
            lblNgayDau.Top = 630
            lblNgayDau.Left = 120
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 600
            txtNgayDau.Left = 1000 '1200
            
            Set lblNgayCuoi.Container = frmKy
            lblNgayCuoi.Top = 630
            lblNgayCuoi.Left = 2600 '2400
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 600
            txtNgayCuoi.Left = 3480
            
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 1000
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1250
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1250
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1250
            txtSolan.Left = 3480
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            
            SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            cmbQuy.Visible = False

        Case KIEU_KY_NGAY_THANG
            Set lblNgay.Container = frmKy
            lblNgay.Top = 480
            lblNgay.Left = 240
            
            Set txtDay.Container = frmKy
            txtDay.Top = 450
            txtDay.Left = 840
            
            Set lblMonth.Container = frmKy
            lblMonth.Top = 480
            lblMonth.Left = 1560
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 450
            txtMonth.Left = 2160
            
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = 2880
            
            Set txtYear.Container = frmKy
            txtYear.Top = 450
            txtYear.Left = 3360
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            
    End Select
    strKHBS = "TKCT"
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutTNCN", Err.Number, Err.Description
    
End Sub

Private Sub SetupLayoutTNCN_QT()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1740
    Frame2.Top = 2050

    Set lblYear.Container = frmKy
    lblYear.Top = 300
    lblYear.Left = 120
    
    Set txtYear.Container = frmKy
    txtYear.Top = 240
    txtYear.Left = 1000 '1200
    
    Set lblTuThang.Container = frmKy
    lblTuThang.Top = 630
    lblTuThang.Left = 120
    
    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 600
    txtNgayDau.Left = 1000 '1200
    
    Set lblDenThang.Container = frmKy
    lblDenThang.Top = 630
    lblDenThang.Left = 2600 '2400
    
    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 600
    txtNgayCuoi.Left = 3480
    
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 1000
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1250
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1250
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1250
    txtSolan.Left = 3480
            
    
    lblSolan.Visible = False
    txtSolan.Visible = False

    SetControlCaption Me, "frmPeriodQuy"
        
    txtMonth.Visible = False
    cmbQuy.Visible = False

    strKHBS = "TKCT"
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutTNCN_QT", Err.Number, Err.Description

End Sub

Private Sub SetupLayoutTNCN_HT_IN_QT()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1140
    Frame2.Top = 2050

    Set lblYear.Container = frmKy
    lblYear.Top = 300
    lblYear.Left = 120
    
    Set txtYear.Container = frmKy
    txtYear.Top = 240
    txtYear.Left = 1000 '1200
    
    Set lblTuThang.Container = frmKy
    lblTuThang.Top = 630
    lblTuThang.Left = 120
    
    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 600
    txtNgayDau.Left = 1000 '1200
    
    Set lblDenThang.Container = frmKy
    lblDenThang.Top = 630
    lblDenThang.Left = 2600 '2400
    
    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 600
    txtNgayCuoi.Left = 3480
        
    lblSolan.Visible = False
    txtSolan.Visible = False

    SetControlCaption Me, "frmPeriodQuy"
        
    txtMonth.Visible = False
    cmbQuy.Visible = False

    strKHBS = "TKCT"
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutTNCN_HT_IN_QT", Err.Number, Err.Description

End Sub

Private Sub SetupLayoutTNDN_DK(strKieuKy As String)
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1365
    Select Case strKieuKy
        Case KIEU_KY_THANG
            Set lblMonth.Container = frmKy
            lblMonth.Top = 200
            lblMonth.Left = 960
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 150
            txtMonth.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = 2730
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 600
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 950
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 950
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
        
        Case KIEU_KY_QUY
            Set lblQuy.Container = frmKy
            lblQuy.Top = 200
            lblQuy.Left = 1050
            
            Frame2.Top = 1815
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 150
            cmbQuy.Left = 1440
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = 2220
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = 2640
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 650
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 900
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 850
            txtSolan.Left = 3400
            lblSolan.Visible = False
            txtSolan.Visible = False
                        
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
                SetControlCaption Me, "frmPeriodBCTC"
            Else
                SetControlCaption Me, "frmPeriodQuy"
            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            cmdOK.Top = 1500
            cmdClose.Top = cmdOK.Top
            Me.Height = 2240
            Me.Width = 4905
            
        Case KIEU_KY_NAM
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = (frmKy.Width - txtYear.Width) / 2 - 100
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = lblYear.Left + lblYear.Width + 50
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 650
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 900
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 850
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False

            
'            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
'                SetControlCaption Me, "frmPeriodBCTC"
'            Else
                SetControlCaption Me, "frmPeriodQuy"
'            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            txtMonth.Visible = False
            cmbQuy.Visible = False
            
        Case KIEU_KY_NGAY_NAM
            Me.Height = 3285
            Me.Width = 5305
            frmKy.Height = 1665
            Frame2.Top = 1815
    
            Set lblYear.Container = frmKy
            lblYear.Top = 300
            lblYear.Left = 120
            
            Set txtYear.Container = frmKy
            txtYear.Top = 240
            txtYear.Left = 1000 '1200
            
            Set lblNgayDau.Container = frmKy
            lblNgayDau.Top = 630
            lblNgayDau.Left = 120
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 600
            txtNgayDau.Left = 1000 '1200
            
            Set lblNgayCuoi.Container = frmKy
            lblNgayCuoi.Top = 630
            lblNgayCuoi.Left = 2600 '2400
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 600
            txtNgayCuoi.Left = 3480
            
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 1000
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1250
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1250
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1250
            txtSolan.Left = 3480
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            
            SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            cmbQuy.Visible = False

        Case KIEU_KY_NGAY_THANG
            Set lblNgay.Container = frmKy
            lblNgay.Top = 480
            lblNgay.Left = 240
            
            Set txtDay.Container = frmKy
            txtDay.Top = 450
            txtDay.Left = 840
            
            Set lblMonth.Container = frmKy
            lblMonth.Top = 480
            lblMonth.Left = 1560
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 450
            txtMonth.Left = 2160
            
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = 2880
            
            Set txtYear.Container = frmKy
            txtYear.Top = 450
            txtYear.Left = 3360
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            
    End Select
    strKHBS = "TKCT"
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutTNDN_DK", Err.Number, Err.Description
    
End Sub


' Set up layout cho to 03TNDN
Private Sub SetupLayout03TNDN()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1740
    Frame2.Top = 2050

    Set lblYear.Container = frmKy
    lblYear.Top = 300
    lblYear.Left = 120
    
    Set txtYear.Container = frmKy
    txtYear.Top = 240
    txtYear.Left = 1000 '1200
    
    Set lblNgayDau.Container = frmKy
    lblNgayDau.Top = 630
    lblNgayDau.Left = 120
    
    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 600
    txtNgayDau.Left = 1000 '1200
    
    Set lblNgayCuoi.Container = frmKy
    lblNgayCuoi.Top = 630
    lblNgayCuoi.Left = 2600 '2400
    
    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 600
    txtNgayCuoi.Left = 3480
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 1050
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1400
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1400
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1400
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    ' Nganh nghe kinh doanh
     frmKy.Height = 2400
    Frame2.Top = 2700
    Set lblNganhKD.Container = frmKy
    lblNganhKD.Top = 1700
    lblNganhKD.Left = 120
    
    Set cboNganhKD.Container = frmKy
    cboNganhKD.Top = 1950
    cboNganhKD.Left = 120
    ' set gia tri nganh nghe kinh doanh cho combo
    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
    

    
    SetControlCaption Me, "frmPeriodQuy"
    
    txtMonth.Visible = False
    cmbQuy.Visible = False
        
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2 - 400
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout03TNDN", Err.Number, Err.Description
    
End Sub

' set up layout to khai 02/NTNN
Private Sub SetupLayout02NTNN()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1400
    Frame2.Top = 1700

'    Set lblYear.Container = frmKy
'    lblYear.Top = 300
'    lblYear.Left = 120
'
'    Set txtYear.Container = frmKy
'    txtYear.Top = 240
'    txtYear.Left = 1000 '1200
    
    Set lblNgayDau.Container = frmKy
    lblNgayDau.Top = 300
    lblNgayDau.Left = 120
    
    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 240
    txtNgayDau.Left = 1000 '1200
    
    Set lblNgayCuoi.Container = frmKy
    lblNgayCuoi.Top = 300
    lblNgayCuoi.Left = 2600 '2400
    
    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 240
    txtNgayCuoi.Left = 3480
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 630
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1000
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1000
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1000
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False

    
    SetControlCaption Me, "frmPeriodQuy"
    
    txtMonth.Visible = False
    cmbQuy.Visible = False
        
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2 - 400
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout02NTNN", Err.Number, Err.Description
    
End Sub

' set up layout to khai 04/TBAC
Private Sub SetupLayout04TBAC()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1400
    
    lblNgay.Visible = True
    
        lblMonth.Visible = True
    txtMonth.Visible = True
    lblYear.Visible = True
    txtYear.Visible = True
    cmbQuy.Visible = False
    
    Set lblNgay.Container = frmKy
    lblNgay.Top = 570
    lblNgay.Left = 120
    
    txtDay.Visible = True
    Set txtDay.Container = frmKy
        txtDay.Top = 540
        txtDay.Left = 700
    
    
    Set lblMonth.Container = frmKy
        lblMonth.Top = 570
        lblMonth.Left = 1360
        
        Set txtMonth.Container = frmKy
        txtMonth.Top = 540
        txtMonth.Left = 1930
        
        Set lblYear.Container = frmKy
        lblYear.Top = 570
        lblYear.Left = 2710
        
        Set txtYear.Container = frmKy
        txtYear.Top = 540
        txtYear.Left = 3130
        
        Dim dTem As Date
        dTem = Date
        txtDay.Text = Day(dTem)
        txtMonth.Text = month(dTem)
        txtYear.Text = Year(dTem)
        
        If Len(txtDay.Text) = 1 Then
            txtDay.Text = "0" & txtDay.Text
        End If
        If Len(txtMonth.Text) = 1 Then
            txtMonth.Text = "0" & txtMonth.Text
        End If
    
        Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2 - 400
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout04TBAC", Err.Number, Err.Description
End Sub

' set up layout to khai 01_TAIN_DK
Private Sub SetupLayout01_TAIN_DK()
    On Error GoTo ErrorHandle
    Dim m, Y, d As Integer
    Dim dTem, dtem1, dtem2 As Date
    Dim varMenuId As String
    dtem2 = Date
    dTem = DateAdd("D", -1, Date)
    dtem1 = DateAdd("M", -1, Date)
    
    
    strLoaiSacThue = "ToKhaiGTGT"
    strLoaiTkDk = "DT"
    Me.Height = 3285
    Me.Width = 4905
    
    'frmKy.Height = 2400
    frmKy.Height = 3000
    'Frame2.Top = 2700
    Frame2.Top = 3300
    
    'frmKy.Height = 1300
    Set chkTKhaiLanXB.Container = frmKy
    chkTKhaiLanXB.Top = 200
    'chkTKhaiLanXB.Left = 3100
    chkTKhaiLanXB.Left = 250

    
    
'    Set lblLanXuat.Container = frmKy
'    lblLanXuat.Top = 1050
'    lblLanXuat.Left = 120
'    lblLanXuat.Visible = True
    
'    Set txtLanXuat.Container = frmKy
'    txtLanXuat.Top = 1050
'    txtLanXuat.Left = 1200
'    txtLanXuat.Visible = True
    
    Set chkTkhaiThang.Container = frmKy
    chkTkhaiThang.Top = 200
    'chkTkhaiThang.Left = 250
    chkTkhaiThang.Left = 3100
    
    Set chkTKLanPS.Container = frmKy
    chkTKLanPS.Top = 200
    chkTKLanPS.Left = 3100
    chkTKLanPS.Width = 1615
    chkTKLanPS.caption = GetAttribute(GetMessageCellById("0299"), "Msg")
    chkTKLanPS.Visible = False
    
    
'    Set lblNganhKD.Container = frmKy
'    lblNganhKD.Top = 1600
'    lblNganhKD.Left = 120
    
    
'    Set cboNganhKD.Container = frmKy
'    cboNganhKD.Top = 1900
'    cboNganhKD.Left = 120

    
    Set lblNgay.Container = frmKy
    lblNgay.Top = 570
    lblNgay.Left = 120
    lblNgay.Visible = False

    Set txtDay.Container = frmKy
    txtDay.Top = 540
    txtDay.Left = 700
    txtDay.Visible = False
    
    Set lblMonth.Container = frmKy
    lblMonth.Top = 570
    lblMonth.Left = 1360
        
    Set txtMonth.Container = frmKy
    txtMonth.Top = 540
    txtMonth.Left = 1930
        
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2710
        
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 3130
        
        
    lblNgay.Visible = True
    txtDay.Visible = True

   ' frmKy.Height = 1600
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 900
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1200
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1200
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1200
    txtSolan.Left = 3400
    
'    lblSolan.Visible = False
'    txtSolan.Visible = False
    
    
    SetValueToListDK ("0")
    strLoaiTKThang_PS = "TK_LANPS"
    'strKieuKy = "D"
    OptChinhthuc.value = True
    lblSolan.Visible = False
    txtSolan.Visible = False
    fpsNgaykhaiBS.Visible = False
    
    
    chkTkhaiThang.value = 0
    chkTKLanPS.value = 0
    chkTKhaiLanXB.value = 1
    frmKy.Height = 3000
    
    cmbQuy.Visible = False
    txtMonth.Visible = True
    
    Set lblLanXuat.Container = frmKy
    lblLanXuat.Top = 1050
    lblLanXuat.Left = 120
    lblLanXuat.Visible = True
    
    Set txtLanXuat.Container = frmKy
    txtLanXuat.Top = 1050
    txtLanXuat.Left = 1200
    txtLanXuat.Visible = True
    
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 1500
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1800
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1800
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1800
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    
    m = month(dtem2)
    Y = Year(dtem2)
    d = Day(dtem2)
    txtDay.Text = d
    txtMonth.Text = m
    txtYear.Text = Y
    If Len(txtDay.Text) = 1 Then
        txtDay.Text = "0" & txtDay.Text
    End If
    If Len(txtMonth.Text) = 1 Then
        txtMonth.Text = "0" & txtMonth.Text
    End If
    
    Frame2.Top = 3300
    
    Set lblNganhKD.Container = frmKy
    lblNganhKD.Top = 2100
    lblNganhKD.Left = 120
    
    Set cboNganhKD.Container = frmKy
    cboNganhKD.Top = 2500
    cboNganhKD.Left = 120
    
    cmbQuy.Visible = False
    lblQuy.Visible = False
    lblNgayDau.Visible = False
    txtNgayDau.Visible = False
    lblNgayCuoi.Visible = False
    txtNgayCuoi.Visible = False
    
    SetControlCaption Me, "frmPeriod"
    strKHBS = "TKCT"
    strQuy = "TK_LANXB"
    
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout01_TAIN_DK", Err.Number, Err.Description
End Sub


' set up layout to khai 02_TNDN_DK
Private Sub SetupLayout02_TNDN_DK()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1740
    Frame2.Top = 2050

    Set lblYear.Container = frmKy
    lblYear.Top = 300
    lblYear.Left = 120
    
    Set txtYear.Container = frmKy
    txtYear.Top = 240
    txtYear.Left = 1000 '1200
    
    Set lblTuThang.Container = frmKy
    lblTuThang.Top = 630
    lblTuThang.Left = 120
    
    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 600
    txtNgayDau.Left = 1000 '1200
    
    Set lblDenThang.Container = frmKy
    lblDenThang.Top = 630
    lblDenThang.Left = 2600 '2400
    
    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 600
    txtNgayCuoi.Left = 3480
    
    ' Set lai max lengh cho to khai tu thang den thang
    txtNgayDau.MaxLength = 7
    txtNgayCuoi.MaxLength = 7
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 1050
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1400
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1400
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1400
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    ' Nganh nghe kinh doanh
     frmKy.Height = 2400
    Frame2.Top = 2700
    Set lblNganhKD.Container = frmKy
    lblNganhKD.Top = 1700
    lblNganhKD.Left = 120
    
    Set cboNganhKD.Container = frmKy
    cboNganhKD.Top = 1950
    cboNganhKD.Left = 120
    ' set gia tri nganh nghe kinh doanh cho combo
    SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
    

    
    SetControlCaption Me, "frmPeriodQuy"
    
    txtMonth.Visible = False
    cmbQuy.Visible = False
        
    SetControlCaption Me, "frmPeriod"
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2 - 400
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout02_TNDN_DK", Err.Number, Err.Description
End Sub

' set up layout to khai 02_TAIN_DK
Private Sub SetupLayout02_TAIN_DK()
    On Error GoTo ErrorHandle
    
    Me.Height = 6000 '4385
    Me.Width = 4905
    frmKy.Height = 2000
    Frame2.Top = 2200
    Frame2.Height = 1500
    Frame2.Visible = True
    Frame2.Enabled = True
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
    
    'lblNgay.Visible = True
    'lblMonth.Visible = True
    'txtMonth.Visible = True
    lblYear.Visible = True
    txtYear.Visible = True
    cmbQuy.Visible = False
    
    'Set lblNgay.Container = frmKy
    'lblNgay.Top = 670
    'lblNgay.Left = 120
    
    'txtDay.Visible = True
'    Set txtDay.Container = frmKy
'    txtDay.Top = 640
'    txtDay.Left = 700
'
'    Set lblMonth.Container = frmKy
'    lblMonth.Top = 670
'    lblMonth.Left = 1360
'
'    Set txtMonth.Container = frmKy
'    txtMonth.Top = 640
'    txtMonth.Left = 1930
        
    Set lblYear.Container = frmKy
    lblYear.Top = 670
    lblYear.Left = 1360
        
    Set txtYear.Container = frmKy
    txtYear.Top = 640
    txtYear.Left = 1760
        
    Dim dTem As Date
    dTem = Date
    txtDay.Text = Day(dTem)
    txtMonth.Text = month(dTem)
    txtYear.Text = Year(dTem)
        
    If Len(txtDay.Text) = 1 Then
        txtDay.Text = "0" & txtDay.Text
    End If

    If Len(txtMonth.Text) = 1 Then
        txtMonth.Text = "0" & txtMonth.Text
    End If
    
    'option
    chkDauTho.Visible = True
    chkCondensate.Visible = True
    chkKhiThien.Visible = True
    Set chkDauTho.Container = frmKy
    Set chkCondensate.Container = frmKy
    Set chkKhiThien.Container = frmKy

    chkDauTho.Top = 200
    chkDauTho.Left = 120
    
    chkCondensate.Top = 200
    chkCondensate.Left = 1500
    chkKhiThien.Top = 200
    chkKhiThien.Left = 3100
    
    OptBosung.Visible = True
    OptChinhthuc.Visible = True
    Set OptBosung.Container = frmKy
    Set OptChinhthuc.Container = frmKy
    
    OptChinhthuc.Top = 1200
    OptChinhthuc.Left = 1600
    OptBosung.Top = 1500
    OptBosung.Left = 1600
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1500
            lblSolan.Left = 3100
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1500
            txtSolan.Left = 3500
            lblSolan.Visible = False
            txtSolan.Visible = False
    
    'set value default
    chkDauTho.value = 1
    chkCondensate.value = 0
    chkKhiThien.value = 0
    
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2 - 400
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout02_TAIN_DK", Err.Number, Err.Description
End Sub

'****************************************************
'Description:LoadDefaultInfor procedure layout default information
'****************************************************

Private Sub LoadDefaultInfor()
    On Error GoTo ErrorHandle
    Dim m As Integer
    Dim q As Quy
    Dim Y As Integer
    Dim d As Integer
    
    'frmTreeviewMenu.Hide
    m = month(Date)
    Y = Year(Date)
    d = Day(Date)
    
    Select Case strKieuKy

        Case KIEU_KY_THANG

            If m = 1 Then
                m = 12
                Y = Y - 1
            Else
                m = m - 1
            End If

            txtMonth.Text = m
            txtYear.Text = Y

            If Len(txtMonth.Text) = 1 Then
                txtMonth.Text = "0" & txtMonth.Text
            End If

        Case KIEU_KY_QUY
            q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)

            If q.q = 1 Then
                q.q = 4
                q.Y = q.Y - 1
            Else
                q.q = q.q - 1
            End If

            cmbQuy.ListIndex = q.q - 1
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "73" Then
            Else
                txtYear.Text = q.Y
            End If

            'dhdang sua them kieu ky nua nam phuc vu an chi
        Case "H_Y"

            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "65" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "18" Then
                q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)

                If q.q = 1 Then
                    q.q = 4
                    q.Y = q.Y - 1
                Else
                    q.q = q.q - 1
                End If

                cmbQuy.ListIndex = q.q - 1
                txtYear.Text = q.Y
                Call initNgayDauNgayCuoiKy(CInt(txtYear.Text), cmbQuy.ListIndex)
            Else
                q = GetKyHienTai(iNgayTaiChinh, iThangTaiChinh)
                cmbQuy.Clear
                cmbQuy.AddItem (1)
                cmbQuy.AddItem (2)
                Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
                txtYear.Text = Y
                cmbQuy.ListIndex = q.q - 1
                Call initNgayDauNgayCuoiKy(CInt(Y), cmbQuy.ListIndex)
            End If

            ' end
        Case KIEU_KY_NAM

            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "89" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "97" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "77" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "88" Then
                Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
                Y = Y - 1
                txtYear.Text = Y
                
                txtNgayDau.Text = "01/" & Y
                txtNgayCuoi.Text = "12/" & Y
            ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "24" Then
                Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
                txtYear.Text = Y
            Else
                Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
                Y = Y - 1
                txtYear.Text = Y
                If GetAttribute(TAX_Utilities_v1.NodeMenu, "ParentID") = "101_10" Then
                    txtNgayDau.Text = "01/" & Y
                    txtNgayCuoi.Text = "12/" & Y
                End If
            End If

        Case KIEU_KY_NGAY_NAM
            Y = GetNamHienTai(iNgayTaiChinh, iThangTaiChinh)
            Y = Y - 1
            txtYear.Text = Y
            yChange = Y
            Call initNgayDauNgayCuoi(CInt(Y))

        Case KIEU_KY_NGAY_THANG

            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "91" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "64" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "07" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "27" Then
                txtDay.Text = d
                txtMonth.Text = m
                txtYear.Text = Y

                If Len(txtDay.Text) = 1 Then
                    txtDay.Text = "0" & txtDay.Text
                End If

                If Len(txtMonth.Text) = 1 Then
                    txtMonth.Text = "0" & txtMonth.Text
                End If

            Else

                If m = 1 Then
                    m = 12
                    Y = Y - 1
                Else
                    m = m - 1
                End If

                txtMonth.Text = m
                txtYear.Text = Y

                If Len(txtMonth.Text) = 1 Then
                    txtMonth.Text = "0" & txtMonth.Text
                End If
            End If

        Case KIEU_KY_NGAY_PS
            txtYear.Text = Y
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "05" Then
                    txtMonth.Text = m - 1
                    If Len(txtMonth.Text) = 1 Then
                        txtMonth.Text = "0" & txtMonth.Text
                    End If

            Else
                    txtDay.Text = d
                    txtMonth.Text = m
                    If Len(txtDay.Text) = 1 Then
                        txtDay.Text = "0" & txtDay.Text
                    End If

                    If Len(txtMonth.Text) = 1 Then
                        txtMonth.Text = "0" & txtMonth.Text
                    End If
            End If
    End Select
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout", Err.Number, Err.Description
    
End Sub

Private Sub initNgayDauNgayCuoi(Y As Integer)
    Dim dDauKyNam As Date
    Dim dCuoiKyNam As Date
    Dim objDateUtils As DateUtils
        
    dDauKyNam = DateSerial(CInt(Y), iThangTaiChinh, iNgayTaiChinh)
    dCuoiKyNam = DateAdd("M", 12, dDauKyNam) - 1
    Set objDateUtils = New DateUtils
    txtNgayDau.Text = objDateUtils.ToString(dDauKyNam, "DD/MM/YYYY")
    txtNgayCuoi.Text = objDateUtils.ToString(dCuoiKyNam, "DD/MM/YYYY")
    Set objDateUtils = Nothing
End Sub
'dhdang sua ngay dau cuoi ky phuc vu an chi
Private Sub initNgayDauNgayCuoiKy(Y As Integer, ky As Integer)
    Dim dDauKyNam As Date
    Dim dCuoiKyNam As Date
    Dim objDateUtils As DateUtils
    
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "65" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "18" Then
            If strQuy = "TK_KY" Then
                If ky = 0 Then
                    dDauKyNam = DateSerial(CInt(Y), 1, 1)
                    dCuoiKyNam = DateSerial(CInt(Y), 7, 1)
                    dCuoiKyNam = DateAdd("D", -1, dCuoiKyNam)
                Else
                    dDauKyNam = DateSerial(CInt(Y), 7, 1)
                    dCuoiKyNam = DateSerial(CInt(Y), 12, 31)
                End If
            Else
                If ky = 0 Then
                    dDauKyNam = DateSerial(CInt(Y), 1, 1)
                    dCuoiKyNam = DateAdd("M", 3, dDauKyNam)
                    dCuoiKyNam = DateDiff("D", 1, dCuoiKyNam)
                ElseIf ky = 1 Then
                    dDauKyNam = DateSerial(CInt(Y), 4, 1)
                    dCuoiKyNam = DateAdd("M", 3, dDauKyNam)
                    dCuoiKyNam = DateDiff("D", 1, dCuoiKyNam)
                ElseIf ky = 2 Then
                    dDauKyNam = DateSerial(CInt(Y), 7, 1)
                    dCuoiKyNam = DateAdd("M", 3, dDauKyNam)
                    dCuoiKyNam = DateDiff("D", 1, dCuoiKyNam)
                ElseIf ky = 3 Then
                    dDauKyNam = DateSerial(CInt(Y), 10, 1)
                    dCuoiKyNam = DateAdd("M", 3, dDauKyNam)
                    dCuoiKyNam = DateDiff("D", 1, dCuoiKyNam)
                End If
            End If
    Else
            If ky = 0 Then
                dDauKyNam = DateSerial(CInt(Y), 1, 1)
                dCuoiKyNam = DateSerial(CInt(Y), 7, 1)
                dCuoiKyNam = DateAdd("D", -1, dCuoiKyNam)
            Else
                dDauKyNam = DateSerial(CInt(Y), 7, 1)
                dCuoiKyNam = DateSerial(CInt(Y), 12, 31)
            End If
    End If
    Set objDateUtils = New DateUtils
    txtNgayDau.Text = objDateUtils.ToString(dDauKyNam, "DD/MM/YYYY")
    txtNgayCuoi.Text = objDateUtils.ToString(dCuoiKyNam, "DD/MM/YYYY")
    Set objDateUtils = Nothing
End Sub
' end
Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    If Frame2.Visible = True Then
        cmdOK.Top = Frame2.Top + Frame2.Height + 70
    Else
        cmdOK.Top = frmKy.Top + frmKy.Height + 70
    End If
    cmdClose.Top = cmdOK.Top
    Me.Height = cmdOK.Top + cmdOK.Height + 140
End Sub

'****************************************************
'Description:Form_Unload procedure realse variable
'****************************************************

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    TAX_Utilities_v1.NodeValidity = Nothing
    i = getFormIndex(TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    arrActiveForm(i).showed = False
    hasActiveForm = False
    
    Set frmPeriod = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Unload", Err.Number, Err.Description
    
End Sub

Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim lCtrl As Long
Dim intSelectAll As Integer

Dim isCheckPLLCTTGT As Boolean

'dhdang sua
Dim rowcheck As Integer

If Not blnFPChange Then
    Exit Sub
End If
rowcheck = 0
intSelectAll = 2
fpSpread1.Col = 1
For lCtrl = 2 To fpSpread1.MaxRows
    fpSpread1.Row = lCtrl
    If fpSpread1.value = 0 Then
        intSelectAll = 0
        Exit For
    ElseIf fpSpread1.value = 1 Then
        intSelectAll = 1
        'rowcheck = lCtrl
    End If
Next lCtrl
For lCtrl = 2 To fpSpread1.MaxRows
    fpSpread1.Row = lCtrl
    If fpSpread1.value = 1 Then
        rowcheck = lCtrl
    ElseIf fpSpread1.value = 2 Then
        rowcheck = -1
    End If
Next lCtrl
If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "18" Then
    chkSelectAll.Enabled = False
    fpSpread1.Col = 1
        For lCtrl = 2 To fpSpread1.MaxRows
            fpSpread1.Row = lCtrl
            If fpSpread1.Row = rowcheck Or rowcheck = 0 Then
               fpSpread1.Lock = False
            Else
                fpSpread1.Lock = True
            End If
        Next lCtrl
Else
    If intSelectAll = 2 Then 'Third state of check
        chkSelectAll.Enabled = False
        chkSelectAll.value = 1
    ElseIf intSelectAll = 1 Then
        chkSelectAll.Enabled = True
        chkSelectAll.value = 1
    Else
        chkSelectAll.Enabled = True
        chkSelectAll.value = 0
    End If
End If

End Sub

Private Sub fpSpread1_Change(ByVal Col As Long, ByVal Row As Long)
'Dim lCtrl As Long
'Dim bSelectAll As Boolean
'
'blnFPChange = True
'bSelectAll = True
'fpSpread1.Col = 1
'For lCtrl = 2 To fpSpread1.MaxRows
'    fpSpread1.Row = lCtrl
'    If fpSpread1.Value = 0 Then
'        bSelectAll = False
'        Exit For
'    End If
'Next lCtrl
'chkSelectAll.Value = bSelectAll
End Sub

Private Sub fpSpread1_GotFocus()
    blnFPChange = True
End Sub

Private Sub fpSpread1_LostFocus()
    blnFPChange = False
End Sub


Private Sub fpSpread1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnClick = True
End Sub



Private Sub lblSelectAll_Click()
    blnFPChange = False
    'chkSelectAll.SetFocus
    If chkSelectAll.Enabled Then
        If chkSelectAll.value = 1 Then
            chkSelectAll.value = 0
        Else
            chkSelectAll.value = 1
        End If
        chkSelectAll_Click
        
    End If
End Sub

Private Sub lblSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnClick = True
    If chkSelectAll.Enabled Then
        chkSelectAll.SetFocus
    Else
        cmdOK.SetFocus
    End If
End Sub



Private Sub OptBosung_Click()
    Dim strDataFileName As String
    Dim xmlDomLastData As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim varMenuId As String
    Dim TIEN_TO_DATA_FILE As String
    Dim CELL_SO_LAN_BS As String
    Dim fso As New FileSystemObject
    Dim i As Integer
    
    Dim vdtehientai As String
    Dim strarrdate() As String

    
    'Check period with valid date
    If CInt(txtYear.Text) < CInt(Right$(GetAttribute(TAX_Utilities_v1.NodeMenu.childNodes(0), "StartDate"), 4)) Then
        DisplayMessage "0092", msOKOnly, miCriticalError, , mrOK
        txtYear.SetFocus
        Exit Sub
    End If
    
    varMenuId = GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID")
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11") Or ((TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15")) Or varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "70" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "73" _
    Or varMenuId = "03" Or varMenuId = "77" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "86" Or varMenuId = "90" Or varMenuId = "87" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "90" Or varMenuId = "88" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "96" Or varMenuId = "97" Or varMenuId = "98" Or varMenuId = "99" Or varMenuId = "24" Then
        For i = 1 To 50
            ' Doi voi to khai thang neu la truong hop bo sung thi quet tat ca cac file xem lan bo sung lon nhat la bao nhieu
            ' Thu tu file bo sung tu 1 den 50
            If (varMenuId = "46" Or varMenuId = "48" Or varMenuId = "15" Or varMenuId = "50" Or varMenuId = "39" Or varMenuId = "36" Or varMenuId = "25" Or varMenuId = "53" Or varMenuId = "54" Or varMenuId = "70" Or varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "71" _
            Or varMenuId = "72" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "86" Or varMenuId = "90" Or varMenuId = "90" Or varMenuId = "92" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "98" Or varMenuId = "96") Then
                If varMenuId = "70" Or varMenuId = "06" Or varMenuId = "90" Then
                    If strLoaiTKThang_PS = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                        strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    End If
                ElseIf varMenuId = "72" Then
                    If strLoaiTKThang_PS = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                        strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                    End If
                'ElseIf varMenuId = "98" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "99" Then
                ElseIf varMenuId = "98" Or varMenuId = "92" Then
'                    If strLoaiTKThang_PS = "TK_THANG" Then
'                        strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
'                    ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                        strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
'                    End If
                     If strQuy = "TK_THANG" Then
                          strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                     ElseIf strQuy = "TK_LANPS" Then
                          strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                     ElseIf strQuy = "TK_LANXB" Then
                          strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                     End If
                Else
                    strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                End If
            ElseIf (varMenuId = "47" Or varMenuId = "49" Or varMenuId = "16" Or varMenuId = "37" Or varMenuId = "51" Or varMenuId = "38" Or varMenuId = "40") Then
                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
            ElseIf varMenuId = "11" Or varMenuId = "12" Or varMenuId = "73" Then
                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
            ElseIf varMenuId = "03" Or varMenuId = "87" Or varMenuId = "88" Or varMenuId = "97" Then
                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & "_" & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml"
'            ElseIf varMenuId = "77" Then
'                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & ".xml"
            ElseIf varMenuId = "93" Or varMenuId = "89" Then
                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Year & ".xml"
'            ElseIf varMenuId = "89" Then
'                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & ".xml"
            ElseIf varMenuId = "80" Or varMenuId = "82" Then
                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml"
            ElseIf varMenuId = "24" Then
                strDataFileName = TAX_Utilities_v1.DataFolder & "bs" & i & "_" & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & ".xml"
            End If
            If fso.FileExists(strDataFileName) Then
                txtSolan.Text = i
            End If
        Next
        
        If OptBosung.value = True Then
            lblSolan.Visible = True
            txtSolan.Visible = True
            strKHBS = "TKBS"
            strSolanBS = txtSolan.Text
            ' kiem tra to khai bs GTGT
            If varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "73" Or varMenuId = "03" _
            Or varMenuId = "77" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "90" Or varMenuId = "88" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "96" Or varMenuId = "97" Or varMenuId = "98" Or varMenuId = "99" Then
                Frame2.Visible = False
                lblSelectAll.Visible = False
                chkSelectAll.Visible = False
                fpSpread1.Visible = False
                fpsNgaykhaiBS.Visible = True
                If varMenuId = "73" Then
                    frmKy.Height = 1800
                    Frame2.Top = 1600
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1250
                    fpsNgaykhaiBS.Left = 960
                'Cap nhat to 02/PHLP
                ElseIf varMenuId = "03" Or varMenuId = "88" Or varMenuId = "87" Or varMenuId = "97" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "77" Then
                    frmKy.Height = 2250
                    Frame2.Top = 2100
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1750
                    fpsNgaykhaiBS.Left = 960
                ElseIf varMenuId = "81" Or varMenuId = "70" Or varMenuId = "06" Or varMenuId = "90" Or varMenuId = "72" Then
                    frmKy.Height = 2050
                    Frame2.Top = 2100
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1550
                    fpsNgaykhaiBS.Left = 960
                ElseIf varMenuId = "80" Or varMenuId = "82" Then
                    frmKy.Height = 1850
                    Frame2.Top = 2000
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1350
                    fpsNgaykhaiBS.Left = 960
'                ElseIf varMenuId = "93" Or varMenuId = "89" Then
'                    frmKy.Height = 2650
'                    Frame2.Top = 2600
'                    Set fpsNgaykhaiBS.Container = frmKy
'                    lblSolan.Top = 1850
'                    lblSolan.Left = 3400
'                    txtSolan.Top = 1850
'                    txtSolan.Left = 3800
'                    txtSolan.Width = 420
'                    lblSolan.Visible = True
'                    fpsNgaykhaiBS.Top = 2150
'                    fpsNgaykhaiBS.Left = 1600
                ElseIf varMenuId = "99" Then
                    frmKy.Height = 1850
                    Frame2.Top = 2100
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1250
                    fpsNgaykhaiBS.Left = 960
        
                Else
                    'kiem tra cac to khai GTGT bo sung them to khai quy
                    If strLoaiSacThue = "ToKhaiGTGT" Then
                        If varMenuId = "98" Or varMenuId = "92" Then
                            If chkTKhaiLanXB.value = 1 Then
                                frmKy.Height = 3600
                                Frame2.Top = 3600
                                Set fpsNgaykhaiBS.Container = frmKy
                                fpsNgaykhaiBS.Top = 2100
                                fpsNgaykhaiBS.Left = 960
                                
                                Set lblNganhKD.Container = frmKy
                                lblNganhKD.Top = 3000
                                lblNganhKD.Left = 120
                                
                                
                                
                                Set cboNganhKD.Container = frmKy
                                cboNganhKD.Top = 3300
                                cboNganhKD.Left = 120
                            Else
                                frmKy.Height = 2000
                                Frame2.Top = 2000
                                Set fpsNgaykhaiBS.Container = frmKy
                                fpsNgaykhaiBS.Top = 1500
                                fpsNgaykhaiBS.Left = 960
                            End If
                        Else
                            frmKy.Height = 2000
                            Frame2.Top = 2000
                            Set fpsNgaykhaiBS.Container = frmKy
                            fpsNgaykhaiBS.Top = 1500
                            fpsNgaykhaiBS.Left = 960
                        End If
                    Else
                        frmKy.Height = 1800
                        Frame2.Top = 2100
                        Set fpsNgaykhaiBS.Container = frmKy
                        fpsNgaykhaiBS.Top = 1250
                        fpsNgaykhaiBS.Left = 960
                    End If
                End If
                ' set gia tri mac dinh cho ngay KHBS
                vdtehientai = format(Date, "dd/mm/yyyy")
                formatPrefix vdtehientai, strarrdate

                With fpsNgaykhaiBS
                    .BackColor = -2147483633
                    .ColHeadersShow = False
                    .RowHeadersShow = False
                    .EditModePermanent = True
                    .EditModeReplace = True
                    .Col = .ColLetterToNumber("C")
                    .Row = 2
                    .BackColor = vbWhite
                    .CellType = CellTypePic
                    .TypePicMask = "99//99//9999"
                    .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
                    
                 End With
                 
                 
                 ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If varMenuId = "01" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "98" Or varMenuId = "92" Then
                    If strLoaiSacThue = "ToKhaiGTGT" Then
                         If varMenuId = "98" Or varMenuId = "92" Then
                            If chkTKhaiLanXB.value = 1 Then
                                frmKy.Height = 3300
                                Frame2.Top = 3300
                                Set fpsNgaykhaiBS.Container = frmKy
                                fpsNgaykhaiBS.Top = 2100
                                fpsNgaykhaiBS.Left = 960
                                
                                Set lblNganhKD.Container = frmKy
                                lblNganhKD.Top = 2500
                                lblNganhKD.Left = 120
                                
                                
                                
                                Set cboNganhKD.Container = frmKy
                                cboNganhKD.Top = 2800
                                cboNganhKD.Left = 120
                            Else
                                frmKy.Height = 2800
                                Frame2.Top = 2900
                                Set lblNganhKD.Container = frmKy
                                lblNganhKD.Top = 1900
                                lblNganhKD.Left = 120
                                
                                
                                Set cboNganhKD.Container = frmKy
                                cboNganhKD.Top = 2200
                                cboNganhKD.Left = 120
                            End If
                        Else
                        
                            frmKy.Height = 2800
                            Frame2.Top = 2900
                            Set lblNganhKD.Container = frmKy
                            lblNganhKD.Top = 1900
                            lblNganhKD.Left = 120
                            
                            
                            Set cboNganhKD.Container = frmKy
                            cboNganhKD.Top = 2200
                            cboNganhKD.Left = 120
                         End If
                        ' set gia tri nganh nghe kinh doanh cho combo
                        'SetValueToList varMenuId
                    Else
                        frmKy.Height = 2600
                        Frame2.Top = 2900
                        Set lblNganhKD.Container = frmKy
                        lblNganhKD.Top = 1700
                        lblNganhKD.Left = 120
                        
                        
                        Set cboNganhKD.Container = frmKy
                        cboNganhKD.Top = 2000
                        cboNganhKD.Left = 120
                        ' set gia tri nganh nghe kinh doanh cho combo
                        'SetValueToList varMenuId
                    End If
                End If
                
                ' Set TK 02/TNDN
                 ' Set loai TK
                'If varMenuId = "73" Then
'                    frmKy.Height = 2800
'                    Frame2.Top = 3000
'                    Set lblNganhKD.Container = frmKy
'                    lblNganhKD.caption = TAX_Utilities_v1.Convert(GetAttribute(GetMessageCellById("0237"), "Msg"), UNICODE, TCVN)
'                    lblNganhKD.Top = 2000
'                    lblNganhKD.Left = 120
'
'
'                    Set cboNganhKD.Container = frmKy
'                    cboNganhKD.Top = 2300
'                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    'SetValueToList "73"
                'End If
                
                ' Set gia tri cho to khai 03/TNDN
                If varMenuId = "03" Or varMenuId = "93" Or varMenuId = "89" Then
                 ' Nganh nghe kinh doanh
                     frmKy.Height = 2800
                    Frame2.Top = 2900
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 2100
                    lblNganhKD.Left = 120
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 2400
                    cboNganhKD.Left = 120
                End If
                
                ' set to khai 02/BVMT
                If varMenuId = "87" Or varMenuId = "97" Or varMenuId = "77" Or varMenuId = "88" Then
                     frmKy.Height = 2250
                    Frame2.Top = 2900
                End If
                
                ' to khai 01/TTDB co them danh muc nganh nghe kinh doanh
                If varMenuId = "05" Then
                    frmKy.Height = 2600
                    Frame2.Top = 2900
                    
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1520
                    fpsNgaykhaiBS.Left = 960
                    
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 2000
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 2000
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    'SetValueToList varMenuId
                End If
                
                                
                Call Form_Resize
            End If
        Else
            lblSolan.Visible = False
            txtSolan.Visible = False
            strKHBS = "TKCT"
            strSolanBS = ""
            ' kiem tra to khai bs GTGT
            If varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "73" _
            Or varMenuId = "77" Or varMenuId = "86" Or varMenuId = "90" Or varMenuId = "87" Or varMenuId = "03" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "90" Or varMenuId = "88" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "96" Or varMenuId = "97" Or varMenuId = "98" Or varMenuId = "99" Then
                ' to khai nao co phu luc thi moi hien thi len
                If Not fpSpread1.Visible And TAX_Utilities_v1.NodeValidity.childNodes.length > 2 Then
                    Frame2.Visible = True
                    lblSelectAll.Visible = True
                    chkSelectAll.Visible = True
                    fpSpread1.Visible = True
                End If
                Set fpsNgaykhaiBS.Container = frmKy
                fpsNgaykhaiBS.Top = 8400
                fpsNgaykhaiBS.Left = 360
                fpsNgaykhaiBS.Visible = False
                frmKy.Height = 1400
                Frame2.Top = 1700
                
                If varMenuId = "01" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "98" Or varMenuId = "92" Then
                    frmKy.Height = 2100
                    Frame2.Top = 2400
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1300
                    lblNganhKD.Left = 120
                    
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1600
                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    'SetValueToList strIdToKhai
                End If
                 
                ' Set TK 02/TNDN
                 ' Set loai TK
                'If varMenuId = "73" Then
'                    frmKy.Height = 2400
'                    Frame2.Top = 2700
'                    Set lblNganhKD.Container = frmKy
'                    lblNganhKD.caption = TAX_Utilities_v1.Convert(GetAttribute(GetMessageCellById("0237"), "Msg"), UNICODE, TCVN)
'                    lblNganhKD.Top = 1600
'                    lblNganhKD.Left = 120
'
'
'                    Set cboNganhKD.Container = frmKy
'                    cboNganhKD.Top = 1900
'                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    'SetValueToList "73"
                'End If
                 
                ' Set gia tri cho to khai 03/TNDN
                If varMenuId = "03" Or varMenuId = "93" Or varMenuId = "89" Then
                 ' Nganh nghe kinh doanh
                     frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1700
                    lblNganhKD.Left = 120
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1950
                    cboNganhKD.Left = 120
                End If
                 
                ' Set to khai 01B/TNDN - DK
                 If varMenuId = "99" Then
                    Frame2.Top = 1815
                 End If
                 Call Form_Resize
            End If
        End If
    ' Doi voi truong hop to khai quyet toan thi van de nhu hien nay
    ElseIf (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        strSolanBS = ""
        
        If Trim(varMenuId) = "17" Then
            TIEN_TO_DATA_FILE = "05_TNCN"
            CELL_SO_LAN_BS = "I_67"
        ElseIf Trim(varMenuId) = "41" Then
            TIEN_TO_DATA_FILE = "09_TNCN"
            CELL_SO_LAN_BS = "I_70"
        ElseIf Trim(varMenuId) = "42" Then
            TIEN_TO_DATA_FILE = "02_TNCN_BH"
            CELL_SO_LAN_BS = "I_48"
        ElseIf Trim(varMenuId) = "43" Then
            TIEN_TO_DATA_FILE = "02_TNCN_SX"
            CELL_SO_LAN_BS = "I_51"
        ElseIf Trim(varMenuId) = "26" Then
            TIEN_TO_DATA_FILE = "02_TNCN_BHDC"
            CELL_SO_LAN_BS = "I_64"
        ElseIf Trim(varMenuId) = "59" Then
            TIEN_TO_DATA_FILE = "06_TNCN10"
            CELL_SO_LAN_BS = "I_61"
        ElseIf Trim(varMenuId) = "76" Then
            TIEN_TO_DATA_FILE = "08B_TNCN"
            CELL_SO_LAN_BS = "I_38"
        ElseIf Trim(varMenuId) = "95" Then
            TIEN_TO_DATA_FILE = "16_TH_DKNPT"
            CELL_SO_LAN_BS = "J_6"
        End If
        
        With xmlDomLastData
            .resolveExternals = True
            .validateOnParse = True
            .async = False
            
                strDataFileName = TAX_Utilities_v1.DataFolder & TIEN_TO_DATA_FILE & "_" & txtYear.Text & ".xml"
                If .Load(strDataFileName) = True Then
                    Set xmlNode = .nodeFromID(CELL_SO_LAN_BS)  'I_41: So lan bo sung
                    txtSolan.Text = GetAttribute(xmlNode, "Value") ' So lan bo sung gan nhat
                    If Trim(txtSolan.Text) = "" Or Trim(txtSolan.Text) = "" Then txtSolan.Text = "1"
                    Set xmlNode = Nothing
                ElseIf .parseError.reason <> vbNullString Then
                    If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                        'file data of last month does not exist
                    Else
                        MsgBox .parseError.reason
                    End If
                    txtSolan.Text = 1  ' So lan bo sung gan nhat
                End If
            
        End With
        
        If OptBosung.value = True Then
            lblSolan.Visible = True
            txtSolan.Visible = True
            strKHBS = "TKCT"
            strSolanBS = txtSolan.Text
        Else
            lblSolan.Visible = False
            txtSolan.Visible = False
            strKHBS = "TKCT"
            strSolanBS = ""
        End If
        
    End If
    
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
End Sub
Private Sub OptChinhthuc_Click()
    Dim varMenuId As String
    varMenuId = GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID")
    ' Doi voi truong hop to khai thang/quy thi van phai giu lai ghi bo sung nhu thong nhat tu phien ban 2.1.0
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11") Or (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15") Or varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "73" _
    Or varMenuId = "03" Or varMenuId = "77" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "90" Or varMenuId = "88" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "96" Or varMenuId = "97" Or varMenuId = "98" Or varMenuId = "99" Or varMenuId = "24" Or varMenuId = "93" Or varMenuId = "89" Then
        If OptBosung.value = True Then
            lblSolan.Visible = False
            txtSolan.Visible = False
            strKHBS = "TKBS"
            strSolanBS = txtSolan.Text
            ' kiem tra to khai bs GTGT
            If varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "73" _
            Or varMenuId = "77" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "03" Or varMenuId = "90" Or varMenuId = "88" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "96" Or varMenuId = "97" Or varMenuId = "98" Or varMenuId = "99" Then
                Frame2.Visible = False
                lblSelectAll.Visible = False
                chkSelectAll.Visible = False
                fpSpread1.Visible = False
                fpsNgaykhaiBS.Visible = True
                
                ' bo sung them to khai quy
                If strLoaiSacThue = "ToKhaiGTGT" Then
                    frmKy.Height = 2000
                    Frame2.Top = 2300
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1450
                    fpsNgaykhaiBS.Left = 960
                Else
                    frmKy.Height = 1800
                    Frame2.Top = 2100
                    Set fpsNgaykhaiBS.Container = frmKy
                    fpsNgaykhaiBS.Top = 1250
                    fpsNgaykhaiBS.Left = 960
                End If
                ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
                If varMenuId = "01" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "98" Or varMenuId = "92" Then
                    ' bo sung them phan to khai quy
                    If strLoaiSacThue = "ToKhaiGTGT" Then
                        frmKy.Height = 2800
                        Frame2.Top = 3100
                        Set lblNganhKD.Container = frmKy
                        lblNganhKD.Top = 1900
                        lblNganhKD.Left = 120
                        
                        
                        Set cboNganhKD.Container = frmKy
                        cboNganhKD.Top = 2200
                        cboNganhKD.Left = 120
                        ' set gia tri nganh nghe kinh doanh cho combo
                        'SetValueToList varMenuId

                    Else
                        frmKy.Height = 2600
                        Frame2.Top = 2900
                        Set lblNganhKD.Container = frmKy
                        lblNganhKD.Top = 1700
                        lblNganhKD.Left = 120
                        
                        
                        Set cboNganhKD.Container = frmKy
                        cboNganhKD.Top = 2000
                        cboNganhKD.Left = 120
                        ' set gia tri nganh nghe kinh doanh cho combo
                        'SetValueToList varMenuId
                    End If
                End If
                
           ' Set TK 02/TNDN
                 ' Set loai TK
                'If varMenuId = "73" Then
'                    frmKy.Height = 2800
'                    Frame2.Top = 3000
'                    Set lblNganhKD.Container = frmKy
'                    lblNganhKD.caption = TAX_Utilities_v1.Convert(GetAttribute(GetMessageCellById("0237"), "Msg"), UNICODE, TCVN)
'                    lblNganhKD.Top = 2000
'                    lblNganhKD.Left = 120
'
'
'                    Set cboNganhKD.Container = frmKy
'                    cboNganhKD.Top = 2300
'                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    'SetValueToList "73"
                'End If
     
            ' Set gia tri cho to khai 03/TNDN
                If varMenuId = "03" Or varMenuId = "93" Or varMenuId = "89" Then
                 ' Nganh nghe kinh doanh
                     frmKy.Height = 2800
                    Frame2.Top = 2900
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 2100
                    lblNganhKD.Left = 120
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 2400
                    cboNganhKD.Left = 120
                End If
 
                ' Set gia tri cho to khai 01/TBVMT
                If varMenuId = "90" Then
                    frmKy.Height = 2200
                    Frame2.Top = 2600
                End If
 

                Call Form_Resize
            End If
        Else
            lblSolan.Visible = False
            txtSolan.Visible = False
            strKHBS = "TKCT"
            strSolanBS = ""
            ' kiem tra to khai bs GTGT
            If varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "06" Or varMenuId = "05" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "73" Or varMenuId = "03" _
            Or varMenuId = "77" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" Or varMenuId = "85" Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "90" Or varMenuId = "88" Or varMenuId = "92" Or varMenuId = "93" Or varMenuId = "89" Or varMenuId = "94" Or varMenuId = "96" Or varMenuId = "97" Or varMenuId = "98" Or varMenuId = "99" Then
                If Not fpSpread1.Visible And TAX_Utilities_v1.NodeValidity.childNodes.length > 2 Then
                    Frame2.Visible = True
                    lblSelectAll.Visible = True
                    chkSelectAll.Visible = True
                    fpSpread1.Visible = True
                End If

                Set fpsNgaykhaiBS.Container = frmKy
                fpsNgaykhaiBS.Top = 8400
                fpsNgaykhaiBS.Left = 360
                fpsNgaykhaiBS.Visible = False
                If varMenuId = "73" Then
'                    frmKy.Height = 1550
'                    Frame2.Top = 1700
                    frmKy.Height = 1400
                    Frame2.Top = 1920
                'Cap nhat to 02/PHLP
                ElseIf varMenuId = "03" Or varMenuId = "88" Or varMenuId = "93" Or varMenuId = "89" Then
                    frmKy.Height = 1740
                    Frame2.Top = 2050
                ElseIf varMenuId = "81" Or varMenuId = "70" Or varMenuId = "06" Or varMenuId = "72" Then
                    frmKy.Height = 1600
                    Frame2.Top = 2050
                ElseIf varMenuId = "80" Or varMenuId = "82" Then
                    frmKy.Height = 1400
                    Frame2.Top = 1700
'                ElseIf varMenuId = "93" Or varMenuId = "89" Then
'                    frmKy.Height = 2250
'                    Frame2.Top = 2600
                ElseIf varMenuId = "99" Then
                    frmKy.Height = 1365
                    Frame2.Top = 1815
                               ' Set gia tri cho to khai 01/TBVMT
                ElseIf varMenuId = "90" Then
                 ' Nganh nghe kinh doanh
                     frmKy.Height = 1600
                    Frame2.Top = 2000
                Else
                    ' bo sung them phan to khai quy
                    If strLoaiSacThue = "ToKhaiGTGT" Then
                        frmKy.Height = 1600
                        Frame2.Top = 1900
                    Else
                        frmKy.Height = 1400
                        Frame2.Top = 1700
                    End If
                End If
                ' set gia tri nganh nghe kinh doanh cho to 01GTGT
                If varMenuId = "01" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "98" Or varMenuId = "92" Then
                    ' bo sung them phan to khai quy
                    If strLoaiSacThue = "ToKhaiGTGT" Then
                        If varMenuId = "98" Or varMenuId = "92" Then
                            If chkTKhaiLanXB.value = 1 Then
                                frmKy.Height = 3000
                                Frame2.Top = 3300
                                Set lblNganhKD.Container = frmKy
                                lblNganhKD.Top = 2100
                                lblNganhKD.Left = 120
                                
                                
                                Set cboNganhKD.Container = frmKy
                                cboNganhKD.Top = 2500
                                cboNganhKD.Left = 120
                            Else
                                frmKy.Height = 2300
                                Frame2.Top = 2600
                                Set lblNganhKD.Container = frmKy
                                lblNganhKD.Top = 1500
                                lblNganhKD.Left = 120
                                
                                
                                Set cboNganhKD.Container = frmKy
                                cboNganhKD.Top = 1800
                                cboNganhKD.Left = 120
                            End If
                        Else
                            frmKy.Height = 2300
                            Frame2.Top = 2600
                            Set lblNganhKD.Container = frmKy
                            lblNganhKD.Top = 1500
                            lblNganhKD.Left = 120
                            
                            
                            Set cboNganhKD.Container = frmKy
                            cboNganhKD.Top = 1800
                            cboNganhKD.Left = 120
                        End If
                        ' set gia tri nganh nghe kinh doanh cho combo
                        'SetValueToList strIdToKhai
                    Else
                        frmKy.Height = 2100
                        Frame2.Top = 2400
                        Set lblNganhKD.Container = frmKy
                        lblNganhKD.Top = 1300
                        lblNganhKD.Left = 120
                        
                        
                        Set cboNganhKD.Container = frmKy
                        cboNganhKD.Top = 1600
                        cboNganhKD.Left = 120
                        ' set gia tri nganh nghe kinh doanh cho combo
                        'SetValueToList strIdToKhai
                    End If
                End If
                
                 ' Set TK 02/TNDN
                 ' Set loai TK
                'If varMenuId = "73" Then
'                    frmKy.Height = 2400
'                    Frame2.Top = 2700
'                    Set lblNganhKD.Container = frmKy
'                    lblNganhKD.caption = TAX_Utilities_v1.Convert(GetAttribute(GetMessageCellById("0237"), "Msg"), UNICODE, TCVN)
'                    lblNganhKD.Top = 1600
'                    lblNganhKD.Left = 120
'
'
'                    Set cboNganhKD.Container = frmKy
'                    cboNganhKD.Top = 1900
'                    cboNganhKD.Left = 120
                    ' set gia tri nganh nghe kinh doanh cho combo
                    'SetValueToList "73"
                'End If
                
                ' Set gia tri cho to khai 03/TNDN
                If varMenuId = "03" Or varMenuId = "93" Or varMenuId = "89" Then
                 ' Nganh nghe kinh doanh
                     frmKy.Height = 2400
                    Frame2.Top = 2700
                    Set lblNganhKD.Container = frmKy
                    lblNganhKD.Top = 1700
                    lblNganhKD.Left = 120
                    
                    Set cboNganhKD.Container = frmKy
                    cboNganhKD.Top = 1950
                    cboNganhKD.Left = 120
                End If
                
                ' set to khai 02/BVMT
                If varMenuId = "87" Or varMenuId = "97" Or varMenuId = "77" Or varMenuId = "88" Then
                     frmKy.Height = 1740
                    Frame2.Top = 2050
                End If
                
                ' set to khai 01/TTDB
                If varMenuId = "05" Then
                        frmKy.Height = 2400
                        Frame2.Top = 2700
                
                        Set lblNganhKD.Container = frmKy
                        lblNganhKD.Top = 1550
                        lblNganhKD.Left = 120
                        
                        Set cboNganhKD.Container = frmKy
                        cboNganhKD.Top = 1850
                        cboNganhKD.Left = 120
                   End If
                Call Form_Resize
            End If
        End If
    ' Doi voi truong hop to khai quyet toan thi van de nhu hien nay
    ElseIf (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        strSolanBS = ""
        If OptBosung.value = True Then
            lblSolan.Visible = False
            txtSolan.Visible = False
            strKHBS = "TKCT"
            strSolanBS = txtSolan.Text
        Else
            lblSolan.Visible = False
            txtSolan.Visible = False
            strKHBS = "TKCT"
            strSolanBS = ""
        End If
    End If
    
End Sub


Private Sub OptTKLanPS_Click()
'    Dim m, Y, d As Integer
'    Dim dTem, dtem1 As Date
'    dTem = DateAdd("D", -1, Date)
'    dtem1 = DateAdd("M", -1, Date)
'    lblNgay.Visible = OptTKLanPS.value
'    txtDay.Visible = OptTKLanPS.value
'    If OptTKLanPS.value = True Then
'        strLoaiTKThang_PS = "TK_LANPS"
'        m = month(dTem)
'        Y = Year(dTem)
'        d = Day(dTem)
'        txtDay.Text = d
'        txtMonth.Text = m
'        txtYear.Text = Y
'        If Len(txtDay.Text) = 1 Then
'            txtDay.Text = "0" & txtDay.Text
'        End If
'        If Len(txtMonth.Text) = 1 Then
'            txtMonth.Text = "0" & txtMonth.Text
'        End If
'    Else
'        strLoaiTKThang_PS = "TK_THANG"
'        m = month(dtem1)
'        Y = Year(dtem1)
'        txtMonth.Text = m
'        txtYear.Text = Y
'        If Len(txtMonth.Text) = 1 Then
'            txtMonth.Text = "0" & txtMonth.Text
'        End If
'    End If
End Sub

Private Sub OptTKThang_Click()
'    Dim m, Y, d As Integer
'    Dim dTem, dtem1 As Date
'    dTem = DateAdd("D", -1, Date)
'    dtem1 = DateAdd("M", -1, Date)
'    lblNgay.Visible = OptTKLanPS.value
'    txtDay.Visible = OptTKLanPS.value
'    If OptTKLanPS.value = True Then
'        strLoaiTKThang_PS = "TK_LANPS"
'        m = month(dTem)
'        Y = Year(dTem)
'        d = Day(dTem)
'        txtDay.Text = d
'        txtMonth.Text = m
'        txtYear.Text = Y
'        If Len(txtDay.Text) = 1 Then
'            txtDay.Text = "0" & txtDay.Text
'        End If
'        If Len(txtMonth.Text) = 1 Then
'            txtMonth.Text = "0" & txtMonth.Text
'        End If
'    Else
'        strLoaiTKThang_PS = "TK_THANG"
'        m = month(dtem1)
'        Y = Year(dtem1)
'        txtMonth.Text = m
'        txtYear.Text = Y
'        If Len(txtMonth.Text) = 1 Then
'            txtMonth.Text = "0" & txtMonth.Text
'        End If
'    End If
    
End Sub

Private Sub txtDay_LostFocus()
    If Len(txtDay.Text) = 1 Then
         txtDay.Text = "0" & txtDay.Text
    End If
End Sub

Private Sub txtLanXuat_Change()
    'If txtLanXuat.Text = "0" Then txtLanXuat.Text = "1"
    'strSoLanXuatBan = txtLanXuat.Text
End Sub

Private Sub txtLanXuat_KeyPress(KeyAscii As Integer)
'    On Error GoTo ErrorHandle
'    Dim sNumber As String
'    sNumber = "0123456789"
'
'    If KeyAscii = vbKeyBack Then Exit Sub
'    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
'        KeyAscii = 0
'    End If
'    Exit Sub
'ErrorHandle:
'    SaveErrorLog Me.Name, "txtLanXuat_KeyPress", Err.Number, Err.Description
End Sub

Private Sub txtLanXuat_LostFocus()
    'txtLanXuat.Text = txtLanXuat.Text)
    If txtLanXuat.Text <> strSoLanXuatBan Then
        strSoLanXuatBan = txtLanXuat.Text
        LoadGrid
    End If
End Sub

Private Sub txtMonth_Change()
    On Error GoTo ErrorHandle

    If Len(txtMonth.Text) <> 0 And Not IsNumeric(txtMonth.Text) Then
        txtMonth.Text = oldMonth
    Else
        oldMonth = txtMonth.Text
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "txtMonth_Change", Err.Number, Err.Description
    
End Sub

'****************************************************
'Description:txtMonth_KeyPress procedure allow only enter number
'****************************************************

Private Sub txtMonth_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    Dim sNumber As String

    sNumber = "0123456789"
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        txtYear.SetFocus
        Exit Sub
    End If
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "txtMonth_KeyPress", Err.Number, Err.Description

End Sub
Private Sub txtDay_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    Dim sNumber As String

    sNumber = "0123456789"
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        txtMonth.SetFocus
        Exit Sub
    End If
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "txtDay_KeyPress", Err.Number, Err.Description

End Sub


Private Sub txtMonth_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Len(Chr(KeyCode)) <> 0 Then
'        If Not IsNumeric(Chr(KeyCode)) Then
'            KeyCode = 0
'        End If
'    End If
End Sub

Private Sub txtMonth_LostFocus()
    On Error GoTo ErrorHandle
    'Dim sFormat As String
    Static blnLostFocusCalling As Boolean 'This variable is used to check
                                          'whether LostFocus procedure is calling

    If blnLostFocusCalling Then Exit Sub
    
'    If bIsClosed Then
'        Exit Sub
'    End If
'
'    If intMonth = 1 Then
'        sFormat = "M"
'    Else
'        sFormat = "Q"
'    End If
'    If Len(txtMonth.Text) > 0 Then
'        Call ValidFormatDate(txtMonth, sFormat)
'    End If
'*************************
'     added
    blnValidInfo(1) = False
    
'    If intMonth = 1 Then
'        sFormat = "M"
'    Else
'        sFormat = "Q"
'    End If
    
    If bIsClosed Then
        Exit Sub
    End If
    
    blnLostFocusCalling = True
    If Len(txtMonth.Text) > 0 Then
        Set objCvt = New DateUtils
        If IsNull(objCvt.ToDate(txtMonth, "MM")) Then
            blnLostFocusCalling = False
            DisplayMessage "0062", msOKOnly, miCriticalError
            txtMonth.SetFocus
            Exit Sub
        Else
            txtMonth = objCvt.ToString(objCvt.ToDate(txtMonth, "MM"), "MM")
        End If
    End If
    
'*************************
    If txtMonth.Text <> TAX_Utilities_v1.month Then
        LoadGrid
    End If
    
    ' set lai ngay dau ky va cuoi ky
    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "18" Then
        If strQuy = "TK_THANG" Then
            ' set ngay dau
            txtNgayDau.Text = "01/" & txtMonth.Text & "/" & txtYear.Text
            ' set ngay cuoi
            Dim temp  As Integer
            Dim temp1 As Date
            temp = CInt(txtMonth.Text) + 1
            If txtMonth.Text = "12" Then
                temp1 = DateSerial(CInt(txtYear.Text) + 1, 1, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            Else
                temp1 = DateSerial(CInt(txtYear.Text), temp, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            End If
        End If
    End If
    
    blnLostFocusCalling = False
    blnValidInfo(1) = True
    Exit Sub
     
ErrorHandle:
    blnLostFocusCalling = False
    blnValidInfo(2) = True
    SaveErrorLog Me.Name, "txtMonth_LostFocus", Err.Number, Err.Description
    
End Sub

Private Sub txtNgayCuoi_LostFocus()
    If bIsClosed Then Exit Sub
    blnValidInfo(4) = False
    
    If Len(txtNgayCuoi.Text) > 0 Then
        Set objCvt = New DateUtils
        '01/KK-TTS
        ' To khai 08/TNCN va to khai 08A/TNCN set o nay tu thang den thang
        If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "75" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "23" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "89" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "97" _
        Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "76" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "59" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "43" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "41" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "17" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "88" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "26" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "45" Then
            If IsNull(objCvt.ToDate(txtNgayCuoi, "MM/YYYY")) Then
                DisplayMessage "0248", msOKOnly, miCriticalError
                txtNgayCuoi.SetFocus
                Exit Sub
            Else
                txtNgayCuoi = objCvt.ToString(objCvt.ToDate(txtNgayCuoi, "MM/YYYY"), "MM/YYYY")
                If txtNgayCuoi.Text <> TAX_Utilities_v1.LastDay Then
                    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "89" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "97" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "88" Then
                        LoadGrid
                    End If
                End If
            End If
        Else
            If IsNull(objCvt.ToDate(txtNgayCuoi, "DD/MM/YYYY")) Then
                DisplayMessage "0071", msOKOnly, miCriticalError
                txtNgayCuoi.SetFocus
                Exit Sub
            Else
                txtNgayCuoi = objCvt.ToString(objCvt.ToDate(txtNgayCuoi, "DD/MM/YYYY"), "DD/MM/YYYY")
                ' NTNN 02,04
                If txtNgayCuoi.Text <> TAX_Utilities_v1.LastDay Then
                    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "82" Then
                        LoadGrid
                    End If
                End If
            End If
        End If
    End If
    blnValidInfo(4) = True
End Sub

Private Sub txtNgayDau_LostFocus()
    If bIsClosed Then Exit Sub
    blnValidInfo(3) = False
    If Len(txtNgayDau.Text) > 0 Then
        Set objCvt = New DateUtils
        ' To khai 08/TNCN va to khai 08A/TNCN set o nay tu thang den thang
        If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "75" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "23" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "89" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "97" _
        Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "76" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "59" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "43" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "41" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "17" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "88" _
        Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "26" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "45" Then
            If IsNull(objCvt.ToDate(txtNgayDau, "MM/YYYY")) Then
                DisplayMessage "0248", msOKOnly, miCriticalError
                txtNgayDau.SetFocus
                Exit Sub
            Else
                txtNgayDau = objCvt.ToString(objCvt.ToDate(txtNgayDau, "MM/YYYY"), "MM/YYYY")
                If txtNgayDau.Text <> TAX_Utilities_v1.LastDay Then
                    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "89" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "97" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "88" Then
                        LoadGrid
                    End If
                End If
            End If
        Else
            If IsNull(objCvt.ToDate(txtNgayDau, "DD/MM/YYYY")) Then
                DisplayMessage "0071", msOKOnly, miCriticalError
                txtNgayDau.SetFocus
                Exit Sub
            Else
                txtNgayDau = objCvt.ToString(objCvt.ToDate(txtNgayDau, "DD/MM/YYYY"), "DD/MM/YYYY")
                ' NTNN 02,04
                If txtNgayDau.Text <> TAX_Utilities_v1.FirstDay Then
                    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "82" Then
                        LoadGrid
                    End If
                End If
            End If
        End If
    End If
    blnValidInfo(3) = True
End Sub

Private Sub txtSolan_Change()
    If txtSolan.Text = "0" Then txtSolan.Text = "1"
    strSolanBS = txtSolan.Text
End Sub

Private Sub txtSolan_LostFocus()
    txtSolan.Text = Val(txtSolan.Text)
End Sub

Private Sub txtYear_Change()
    On Error GoTo ErrorHandle
    
    If Len(txtYear.Text) <> 0 And Not IsNumeric(txtYear.Text) Then
        txtYear.Text = oldYear
    Else
        oldYear = txtYear.Text
    End If
    
    Exit Sub
 
ErrorHandle:
    SaveErrorLog Me.Name, "txtMonth_Change", Err.Number, Err.Description
    
End Sub

Private Sub txtYear_GotFocus()
    blnClick = False
End Sub

'****************************************************
'Description:txtYear_KeyPress procedure allow only enter number
'****************************************************

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    Dim sNumber As String
    sNumber = "0123456789"
    
    If KeyAscii = vbKeyBack Then Exit Sub
    '********************************
    If KeyAscii = vbKeyReturn Then
        If chkSelectAll.Enabled Then
            chkSelectAll.SetFocus
        Else
            cmdOK.SetFocus
        End If
        Exit Sub
    End If
    '********************************
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If

    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "txtYear_KeyPress", Err.Number, Err.Description

End Sub

Private Sub txtSolan_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    Dim sNumber As String
    sNumber = "0123456789"
    
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "txtSolan_KeyPress", Err.Number, Err.Description

End Sub



Private Sub txtYear_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Len(Chr(KeyCode)) <> 0 Then
'        If Not IsNumeric(Chr(KeyCode)) Then
'            KeyCode = 0
'        End If
'    End If
End Sub

Private Sub txtYear_LostFocus()
    On Error GoTo ErrorHandle
    Static blnLostFocusCalling As Boolean 'This variable is used to check
                                          'whether LostFocus procedure is calling
    
    
    If blnLostFocusCalling Then Exit Sub
    
    blnValidInfo(2) = False
    
    If bIsClosed Then
        Exit Sub
    End If
    If txtYear.Text = "9999" Then txtYear.Text = "9998"
    blnLostFocusCalling = True
    If Len(txtYear.Text) > 0 Then
        Set objCvt = New DateUtils
        If IsNull(objCvt.ToDate(txtYear, "YYYY")) Then
            blnLostFocusCalling = False
            DisplayMessage "0062", msOKOnly, miCriticalError
            txtYear.SetFocus
            Exit Sub
        Else
            txtYear.Text = objCvt.ToString(objCvt.ToDate(txtYear, "YYYY"), "YYYY")
        End If
        
        If CInt(txtYear.Text) < 2000 Then
            blnLostFocusCalling = False
            DisplayMessage "0067", msOKOnly, miCriticalError
            txtYear.SetFocus
            Exit Sub
        End If
    End If
    If txtYear.Text <> TAX_Utilities_v1.Year And GetAttribute(TAX_Utilities_v1.NodeMenu, "Day") = "1" Then
        txtNgayCuoi.Enabled = True
        txtNgayDau.Enabled = True
    End If
    
    If txtYear.Text <> yChange And GetAttribute(TAX_Utilities_v1.NodeMenu, "Day") = "1" And txtNgayCuoi.Enabled And txtNgayDau.Enabled Then
        Call initNgayDauNgayCuoi(CInt(txtYear.Text))
    End If
    'dhdang them ky nua nam phuc vu an chi
    If txtYear.Text <> yChange And GetAttribute(TAX_Utilities_v1.NodeMenu, "Year") = "1/2" And txtNgayCuoi.Enabled And txtNgayDau.Enabled Then
        Call initNgayDauNgayCuoiKy(CInt(txtYear.Text), cmbQuy.ListIndex)
    End If
    ' end
    ' set lai tu thang den thang
    If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") _
    Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "88" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "97" _
    Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "89" Then
         If txtYear.Text <> yChange And txtNgayCuoi.Enabled And txtNgayDau.Enabled Then
            txtNgayDau.Text = "01/" & txtYear.Text
        End If
        
        If txtYear.Text <> yChange And txtNgayCuoi.Enabled And txtNgayDau.Enabled Then
            txtNgayCuoi.Text = "12/" & txtYear.Text
        End If
    End If
    ' end
    
    If txtYear.Text <> TAX_Utilities_v1.Year Then
        
        LoadGrid
        
        ' Cac to quyet toan TNCN kiem tra xem de dat lai nut Dong y, Dong cho dung, dat sau Frame 2
        If (TAX_Utilities_v1.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
            Frame2.Visible = False
            lblSelectAll.Visible = False
            chkSelectAll.Visible = False
            fpSpread1.Visible = False
            Call Form_Resize
        End If
    End If
    
    If txtNgayDau.Enabled And txtNgayDau.Visible And Not blnClick Then
        txtNgayDau.SetFocus
    End If


    ' set lai ngay dau ky va cuoi ky
    If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "18" Then
        If strQuy = "TK_THANG" Then
            ' set ngay dau
            txtNgayDau.Text = "01/" & txtMonth.Text & "/" & txtYear.Text
            ' set ngay cuoi
            Dim temp  As Integer
            Dim temp1 As Date
            temp = CInt(txtMonth.Text) + 1
            If txtMonth.Text = "12" Then
                temp1 = DateSerial(CInt(txtYear.Text) + 1, 1, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            Else
                temp1 = DateSerial(CInt(txtYear.Text), temp, 1)
                temp1 = DateAdd("D", -1, temp1)
                txtNgayCuoi.Text = Day(temp1) & "/" & format(month(temp1), "0#") & "/" & Year(temp1)
            End If
        End If
    End If

    blnLostFocusCalling = False
    blnValidInfo(2) = True
    yChange = txtYear.Text
    
    Me.Refresh
    Exit Sub
ErrorHandle:
    blnLostFocusCalling = False
    blnValidInfo(2) = True
    SaveErrorLog Me.Name, "txtYear_LostFocus", Err.Number, Err.Description

End Sub

Private Sub fpSpread1_KeyPress(KeyAscii As Integer)
Dim lRow As Long

lRow = fpSpread1.ActiveRow
If KeyAscii = vbKeyReturn Then
    If lRow = fpSpread1.MaxRows Then
        cmdOK.SetFocus
    End If
    fpSpread1.SetActiveCell 1, fpSpread1.ActiveRow + 1
    'This is the last unlocked cell
    If lRow = fpSpread1.ActiveRow Then
        cmdOK.SetFocus
    End If
End If
End Sub

Private Sub SetActiveValue()
Dim lCtrl As Long

For lCtrl = 1 To fpSpread1.MaxRows
    fpSpread1.Row = lCtrl
    SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(lCtrl - 1), "Active", fpSpread1.value
Next lCtrl
End Sub

'*****************************************************
'Description: LoadGrid procedure initialize and setup value to grid
'Author:
'Date:10/11/2005
'Input:
'Output:
'Return:
'*****************************************************
Private Sub LoadGrid()
Dim xmlNode As MSXML.IXMLDOMNode
Dim fso As New FileSystemObject, fle As file
Dim lCtrl As Long, lRow As Long, lLoc As Long
Dim strDataFileName As String
Dim blnExistData As Boolean, blnExceptData As Boolean

Dim strIDTkhai As String

On Error GoTo ErrHandle
    
    'set data
    TAX_Utilities_v1.Year = txtYear.Text
    TAX_Utilities_v1.month = txtMonth.Text
    TAX_Utilities_v1.Day = txtDay.Text
    TAX_Utilities_v1.ThreeMonths = cmbQuy.Text
    TAX_Utilities_v1.FirstDay = txtNgayDau.Text
    TAX_Utilities_v1.LastDay = txtNgayCuoi.Text
    TAX_Utilities_v1.NodeValidity = GetValidityNode
    
    SetDefaultActiveProperties

    lRow = 1
    blnExistData = True
    
    With fpSpread1
        .NoBeep = True
        .ReDraw = False
        .EventEnabled(EventButtonClicked) = False
        .MaxCols = 1
        .MaxRows = 1
        .TabStripPolicy = TabStripPolicyNever
        .RowHeadersShow = False
        .ColHeadersShow = False
        .ScrollBars = ScrollBarsVertical

        .Col = 1
        .ColWidth(1) = 37
        
        ' xu ly cho to khai DK
        strIDTkhai = GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID")
        ' end
        
        For Each xmlNode In TAX_Utilities_v1.NodeValidity.childNodes
          If GetAttribute(xmlNode, "Caption") <> "KHBS" Then
                ' Get name of data file
                If blnExistData Then
                    If strKieuKy = KIEU_KY_THANG Then
                        If Trim$(strIDTkhai) = "92" Or Trim$(strIDTkhai) = "98" Then
'                            If strQuy = "TK_THANG" Then
'                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtMonth.Text & txtYear.Text & ".xml"
'                            ElseIf strQuy = "TK_LANXB" Then
'                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_Q0" & cmbQuy.Text & txtYear.Text & ".xml"
'                            Else
'                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtMonth.Text & txtYear.Text & ".xml"
'                            End If
                            
                             If strQuy = "TK_THANG" Then
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                            ElseIf strQuy = "TK_LANPS" Then
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                            ElseIf strQuy = "TK_LANXB" Then
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v1.Day & TAX_Utilities_v1.month & TAX_Utilities_v1.Year & ".xml"
                            End If
                        Else
                            If strQuy = "TK_THANG" Then
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtMonth.Text & txtYear.Text & ".xml"
                            ElseIf strQuy = "TK_QUY" Then
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_Q0" & cmbQuy.Text & txtYear.Text & ".xml"
                            Else
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtMonth.Text & txtYear.Text & ".xml"
                            End If
                        End If
                    ElseIf strKieuKy = KIEU_KY_QUY Then
                        If strQuy = "TK_TU_THANG" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & TAX_Utilities_v1.FirstDay & "_" & TAX_Utilities_v1.LastDay & ".xml"
                        Else
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_0" & cmbQuy.Text & txtYear.Text & ".xml"
                        End If
                    ' phuc vu an chi
                    ElseIf strKieuKy = "H_Y" Then
                        If strQuy = "TK_THANG" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_T" & txtMonth.Text & txtYear.Text & ".xml"
                        Else
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_0" & cmbQuy.Text & txtYear.Text & ".xml"
                        End If
                    ' end
                    ElseIf strKieuKy = KIEU_KY_NGAY_NAM Then
                        If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "82" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" _
                            & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml"
                        Else
                            If xmlNode Is TAX_Utilities_v1.NodeValidity.firstChild Then
                                For Each fle In fso.GetFolder(TAX_Utilities_v1.DataFolder).Files
                                    lLoc = InStr(1, fle.Name, GetAttribute(xmlNode, "DataFile") & "_" & txtYear.Text & "_")
                                    If lLoc <> 0 Then
                                        lLoc = lLoc + Len(GetAttribute(xmlNode, "DataFile") & "_" & txtYear.Text & "_")
                                        txtNgayDau.Text = Mid$(fle.Name, lLoc, 2) & "/" & Mid$(fle.Name, lLoc + 2, 2) & "/" & Mid$(fle.Name, lLoc + 4, 4)
                                        txtNgayCuoi.Text = Mid$(fle.Name, lLoc + 9, 2) & "/" & Mid$(fle.Name, lLoc + 11, 2) & "/" & Mid$(fle.Name, lLoc + 13, 4)
'                                        txtNgayCuoi.Enabled = False
'                                        txtNgayDau.Enabled = False
                                        blnExceptData = True
                                        Exit For
                                    End If
                                Next
                            Else
                                strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtYear.Text & _
                                        "_" & Replace$(txtNgayDau.Text, "/", "") & "_" & Replace$(txtNgayCuoi.Text, "/", "") & ".xml"
                            End If
                        End If
                    ElseIf strKieuKy = KIEU_KY_NAM Then
                        If GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "89" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & strLoaiTkDk & "_" & txtYear.Text & "_" & Replace$(txtNgayDau.Text, "/", "") & "_" & Replace$(txtNgayCuoi.Text, "/", "") & ".xml"
                        ElseIf GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "97" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "77" _
                        Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "88" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtYear.Text & "_" & Replace$(txtNgayDau.Text, "/", "") & "_" & Replace$(txtNgayCuoi.Text, "/", "") & ".xml"
                        ElseIf GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "76" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "59" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "43" _
                        Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "41" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "17" Or GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "26" Then
                            ' QT TNCN
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtYear.Text & "_" & Replace$(txtNgayDau.Text, "/", "") & "_" & Replace$(txtNgayCuoi.Text, "/", "") & ".xml"
                        ElseIf GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID") = "95" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_L" & strSolanKK & "_" & txtYear.Text & ".xml"
                        Else
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtYear.Text & ".xml"
                        End If
                    ElseIf strKieuKy = KIEU_KY_NGAY_THANG Then
                        If strLoaiTKThang_PS = "TK_LANPS" Then
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtDay.Text & txtMonth.Text & txtYear.Text & ".xml"
                        Else
                            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(xmlNode, "DataFile") & "_" & txtMonth.Text & txtYear.Text & ".xml"
                        End If
                    End If
                End If
                
                'By default number of row is one
                'If it has more than one row
                If lRow > 1 Then  'Insert new row
                    .MaxRows = .MaxRows + 1
                    .InsertRows lRow, 1
                End If
                
                .Row = lRow
                .CellType = CellTypeCheckBox
                .TypeCheckType = TypeCheckTypeNormal
                .TypeCheckTextAlign = TypeCheckTextAlignRight
                .TypeCheckText = GetAttribute(xmlNode, "Caption")
                
                ' Check the exist of data file -> Set value to Checkbox
                
                If xmlNode Is TAX_Utilities_v1.NodeValidity.firstChild Then
                    If fso.FileExists(strDataFileName) Or blnExceptData Then
                    Else
                        'To khai ko ton tai
                        blnExistData = False
                    End If
                End If
                
                If blnExistData Then
                    If fso.FileExists(strDataFileName) Or blnExceptData Then
                        .TypeCheckType = TypeCheckTypeThreeState
                        .value = 2
                        .Lock = True
                        blnExceptData = False
                    End If
                End If
                'Resize row height: Auto fit with content
                .RowHeight(lRow) = .MaxTextRowHeight(lRow)
            
            lRow = lRow + 1
          End If
        Next
        
        'New tax -> Set default value by Menu.xml
        If Not blnExistData Then
            For lCtrl = 2 To lRow - 1
                .Row = lCtrl
                .value = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(lCtrl - 1), "Active")
            Next lCtrl
        End If
        'Do not allow edit and hide first row
        .Row = 1
        .TypeCheckType = TypeCheckTypeThreeState
        .Lock = True
        .value = 2
        .RowHidden = True
        
        If Not fpSpread1.Visible And TAX_Utilities_v1.NodeValidity.childNodes.length > 2 Then
            fpSpread1.Visible = True
           ' Frame2.Visible = True
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "21" Then
                Frame2.Visible = False
            Else
               Frame2.Visible = True
            End If
            
            ' to khai BS se khong co danh sach phu luc
            If strKHBS = "TKBS" Then
                Frame2.Visible = False
            End If
            'Me.Height = Me.Height + Frame2.Height - 50
            'cmdOK.Top = cmdOK.Top + Frame2.Height - 50
            'cmdClose.Top = cmdClose.Top + Frame2.Height - 50
        End If
        
        'Set cursor style and edit mode to fpSpread
        .CursorStyle = CursorStyleArrow
        .EditModePermanent = True
        .GrayAreaBackColor = vbButtonFace
        .ReDraw = True
        .EventEnabled(EventButtonClicked) = True
    End With
    
    blnFPChange = True
    fpSpread1_ButtonClicked 1, 1, 1
    blnFPChange = False
Exit Sub

ErrHandle:
    SaveErrorLog Me.Name, "LoadGrid", Err.Number, Err.Description
End Sub

Private Sub SetDefaultActiveProperties()
    Dim xmlDom      As New MSXML.DOMDocument
    Dim xmlNodeMenu As MSXML.IXMLDOMNode, xmlNodeValidity As MSXML.IXMLDOMNode
    Dim strTemp     As String, lCtrl As Long

    xmlDom.Load App.path & "\Menu.xml"

    strTemp = GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")

    For Each xmlNodeMenu In xmlDom.getElementsByTagName("Menu")

        If GetAttribute(xmlNodeMenu, "ID") = strTemp Then Exit For
    Next

    strTemp = GetAttribute(TAX_Utilities_v1.NodeValidity, "StartDate")

    For Each xmlNodeValidity In xmlNodeMenu.childNodes

        If GetAttribute(xmlNodeValidity, "StartDate") = strTemp Then Exit For
    Next

    For lCtrl = 0 To TAX_Utilities_v1.NodeValidity.childNodes.length - 1
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(lCtrl), "Active", GetAttribute(xmlNodeValidity.childNodes(lCtrl), "Active")
    Next lCtrl

    Set xmlNodeMenu = Nothing
    Set xmlNodeValidity = Nothing
    Set xmlDom = Nothing
    

End Sub


Private Function ExistTokhai(LoaiTk As String, thang As Boolean, period As String) As Boolean
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As file
    Dim strFileName1 As String, strFileName2 As String, strFileName3 As String, strFileName4 As String, strFileName5 As String, strFileName6 As String
    Dim strloaitk As String
    ExistTokhai = False
    
    
    If thang Then
        If Left(period, 1) = "1" Then
           strFileName1 = LoaiTk & "01" & Right(period, 4) & ".xml"
           strFileName2 = LoaiTk & "02" & Right(period, 4) & ".xml"
           strFileName3 = LoaiTk & "03" & Right(period, 4) & ".xml"
        ElseIf Left(period, 1) = "2" Then
           strFileName1 = LoaiTk & "04" & Right(period, 4) & ".xml"
           strFileName2 = LoaiTk & "05" & Right(period, 4) & ".xml"
           strFileName3 = LoaiTk & "06" & Right(period, 4) & ".xml"
        ElseIf Left(period, 1) = "3" Then
           strFileName1 = LoaiTk & "07" & Right(period, 4) & ".xml"
           strFileName2 = LoaiTk & "08" & Right(period, 4) & ".xml"
           strFileName3 = LoaiTk & "09" & Right(period, 4) & ".xml"
        Else
           strFileName1 = LoaiTk & "10" & Right(period, 4) & ".xml"
           strFileName2 = LoaiTk & "11" & Right(period, 4) & ".xml"
           strFileName3 = LoaiTk & "12" & Right(period, 4) & ".xml"
         End If
    Else
        If Left(period, 2) = "01" Or Left(period, 2) = "02" Or Left(period, 2) = "03" Then
            period = "01" & Right(period, 4)
        ElseIf Left(period, 2) = "04" Or Left(period, 2) = "05" Or Left(period, 2) = "06" Then
            period = "02" & Right(period, 4)
        ElseIf Left(period, 2) = "07" Or Left(period, 2) = "08" Or Left(period, 2) = "09" Then
            period = "03" & Right(period, 4)
        Else
            period = "04" & Right(period, 4)
        End If
        
         strFileName1 = LoaiTk & period & ".xml"
        
    End If
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v1.DataFolder)).Files
        If thang Then
            If InStr(1, fle.Name, strFileName1) > 0 Or InStr(1, fle.Name, strFileName2) > 0 Or InStr(1, fle.Name, strFileName3) > 0 Then
                ExistTokhai = True
                Exit Function
            End If
        Else
            If InStr(1, fle.Name, strFileName1) > 0 Then
                ExistTokhai = True
                Exit Function
            End If
        End If
    Next
End Function

Private Sub SetupLayoutNTNN()
    On Error GoTo ErrorHandle
    
    Me.Height = 3285
    Me.Width = 4905
    
    'frmKy.Height = 1300
        Set chkTkhaiThang.Container = frmKy
        chkTkhaiThang.Top = 200
        chkTkhaiThang.Left = 120
        chkTkhaiThang.value = 1
        'chkTkhaiThang.Enabled = False
        chkTKLanPS.value = 0
        
        
        Set chkTKLanPS.Container = frmKy
        chkTKLanPS.Top = 200
        chkTKLanPS.Left = 1800
        
        Set lblNgay.Container = frmKy
        lblNgay.Top = 570
        lblNgay.Left = 120

        Set txtDay.Container = frmKy
        txtDay.Top = 540
        txtDay.Left = 700
        
        
        
        Set lblMonth.Container = frmKy
        lblMonth.Top = 570
        lblMonth.Left = 1360
        
        Set txtMonth.Container = frmKy
        txtMonth.Top = 540
        txtMonth.Left = 1930
        
        Set lblYear.Container = frmKy
        lblYear.Top = 570
        lblYear.Left = 2710
        
        Set txtYear.Container = frmKy
        txtYear.Top = 540
        txtYear.Left = 3130
        
        cmbQuy.Visible = False
        txtNgayDau.Visible = False
        txtNgayCuoi.Visible = False
        
        If chkTkhaiThang.value = 1 Then
            frmKy.Height = 1600
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 900
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1200
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1200
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1200
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False
            
    
        Else
            frmKy.Height = 1700
        End If
  
        
        
'

'
'
'
'
         SetControlCaption Me, "frmPeriod"
'
'         cmbQuy.Visible = False
'         txtNgayDau.Visible = False
'         txtNgayCuoi.Visible = False
'
            
'    If intMonth = 0 Then
'        Me.Width = 3900 '3840
'        Me.Height = 1600
'        Frame1.Height = 720
        
'       lblYear.Top = 300
'        lblYearFormat.Top = 300
        
'        txtYear.Top = 255
        
'        cmdOk.Top = 1100
'        cmdClose.Top = 1100
'    Else
'        Me.Width = 3900 '3840
'        Me.Height = 1920
'        Frame1.Height = 1065
    
'        lblMonth.Top = 300
'        lblYear.Top = 615
'        lblYearFormat.Top = 615
        
'        txtMonth.Top = 255
'        txtYear.Top = 585
        
'        cmdOk.Top = 1410
'        cmdClose.Top = 1410
'    End If
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutNTNN", Err.Number, Err.Description
    
End Sub



Private Sub SetupLayoutTTDB()
    On Error GoTo ErrorHandle
    
    Me.Height = 3285
    Me.Width = 4905
    
    frmKy.Height = 1700
    Frame2.Top = 2000
    
    
    Set chkTkhaiThang.Container = frmKy
    chkTkhaiThang.Top = 200
    chkTkhaiThang.Left = 120
    chkTkhaiThang.value = 1
    chkTKLanPS.value = 0
    
    Set chkTKLanPS.Container = frmKy
    chkTKLanPS.Top = 180
    chkTKLanPS.Left = 1800
    
    Set lblNgay.Container = frmKy
    lblNgay.Top = 570
    lblNgay.Left = 120

    Set txtDay.Container = frmKy
    txtDay.Top = 540
    txtDay.Left = 700
    
    
    
    Set lblMonth.Container = frmKy
    lblMonth.Top = 570
    lblMonth.Left = 1360
    
    Set txtMonth.Container = frmKy
    txtMonth.Top = 540
    txtMonth.Left = 1930
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2710
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 3130
    
 
 
    
    If chkTkhaiThang.value = 1 Then
        frmKy.Height = 2400
        Frame2.Top = 2700
        
        
        Set OptChinhthuc.Container = frmKy
        OptChinhthuc.Top = 900
        OptChinhthuc.Left = 960
        
        Set OptBosung.Container = frmKy
        OptBosung.Top = 1200
        OptBosung.Left = 960
        
        Set lblSolan.Container = frmKy
        lblSolan.Top = 1200
        lblSolan.Left = 3000
        Set txtSolan.Container = frmKy
        txtSolan.Top = 1200
        txtSolan.Left = 3400
        
        lblSolan.Visible = False
        txtSolan.Visible = False
        

        Set lblNganhKD.Container = frmKy
        lblNganhKD.Top = 1550
        lblNganhKD.Left = 120
        
        Set cboNganhKD.Container = frmKy
        cboNganhKD.Top = 1850
        cboNganhKD.Left = 120
    Else
        frmKy.Height = 1700
        Frame2.Top = 2000

        Set lblNganhKD.Container = frmKy
        lblNganhKD.Top = 950
        lblNganhKD.Left = 120
        
        Set cboNganhKD.Container = frmKy
        cboNganhKD.Top = 1200
        cboNganhKD.Left = 120
    End If
    ' set gia tri nganh nghe kinh doanh cho combo
    SetValueToList "05"
    
    
    cmbQuy.Visible = False
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
    
    SetControlCaption Me, "frmPeriod"
        
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutTTDB", Err.Number, Err.Description
    
End Sub

' Set up layout cho to khai 02_TNDN
Private Sub SetupLayout02TNDN()
    On Error GoTo ErrorHandle
    
    Dim m, Y, d As Integer
    Dim dTem, dtem1, dtem2 As Date
    Dim varMenuId As String
    dtem2 = Date
    dTem = Date
    dtem1 = DateAdd("M", -1, Date)
    
    Me.Height = 3285
    Me.Width = 4905
    

    Set lblNgay.Container = frmKy
    lblNgay.Top = 250
    lblNgay.Left = 120
    lblNgay.Visible = True

    Set txtDay.Container = frmKy
    txtDay.Top = 220
    txtDay.Left = 700
    txtDay.Visible = True
    
    
    Set lblMonth.Container = frmKy
    lblMonth.Top = 250
    lblMonth.Left = 1360

    Set txtMonth.Container = frmKy
    txtMonth.Top = 220
    txtMonth.Left = 1930
    
       
    Set lblYear.Container = frmKy
    lblYear.Top = 250
    lblYear.Left = 2710
    
    Set txtYear.Container = frmKy
    txtYear.Top = 220
    txtYear.Left = 3130
        
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
    
     strLoaiTKThang_PS = "TK_LANPS"
     chkTKLanPS.value = "1"
     'strKieuKy = "D"
     OptChinhthuc.value = True
     lblSolan.Visible = False
     txtSolan.Visible = False
     fpsNgaykhaiBS.Visible = False
    
     frmKy.Height = 1400
         
     Set OptChinhthuc.Container = frmKy
     OptChinhthuc.Top = 600
     OptChinhthuc.Left = 960
         
     Set OptBosung.Container = frmKy
     OptBosung.Top = 900
     OptBosung.Left = 960
         
     Set lblSolan.Container = frmKy
     lblSolan.Top = 950
     lblSolan.Left = 3000
     Set txtSolan.Container = frmKy
     txtSolan.Top = 900
     txtSolan.Left = 3400
         
     lblSolan.Visible = False
     txtSolan.Visible = False
     
     m = month(dTem)
    Y = Year(dTem)
    d = Day(dTem)
    txtDay.Text = d
    txtMonth.Text = m
    txtYear.Text = Y
    If Len(txtDay.Text) = 1 Then
    txtDay.Text = "0" & txtDay.Text
    End If
    If Len(txtMonth.Text) = 1 Then
    txtMonth.Text = "0" & txtMonth.Text
    End If
                
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout02TNDN", Err.Number, Err.Description
    
End Sub



' Set up layout cho to khai 01_TBVMT
Private Sub SetupLayout01TBVMT()
    On Error GoTo ErrorHandle
    
    Me.Height = 3285
    Me.Width = 4905
    
    frmKy.Height = 1600
    Frame2.Top = 1700
    
    
    Set chkTkhaiThang.Container = frmKy
    chkTkhaiThang.Top = 200
    chkTkhaiThang.Left = 120
    chkTkhaiThang.value = 1
    chkTKLanPS.value = 0
    
    Set chkTKLanPS.Container = frmKy
    chkTKLanPS.Top = 180
    chkTKLanPS.Left = 1800
    
    Set lblNgay.Container = frmKy
    lblNgay.Top = 570
    lblNgay.Left = 120

    Set txtDay.Container = frmKy
    txtDay.Top = 540
    txtDay.Left = 700
    
    
    
    Set lblMonth.Container = frmKy
    lblMonth.Top = 570
    lblMonth.Left = 1360
    
    Set txtMonth.Container = frmKy
    txtMonth.Top = 540
    txtMonth.Left = 1930
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2710
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 3130
    
 
 
    
    If chkTkhaiThang.value = 1 Then
        frmKy.Height = 1600
        Frame2.Top = 2000
        
        Set OptChinhthuc.Container = frmKy
        OptChinhthuc.Top = 900
        OptChinhthuc.Left = 960
        
        Set OptBosung.Container = frmKy
        OptBosung.Top = 1200
        OptBosung.Left = 960
        
        Set lblSolan.Container = frmKy
        lblSolan.Top = 1200
        lblSolan.Left = 3000
        Set txtSolan.Container = frmKy
        txtSolan.Top = 1200
        txtSolan.Left = 3400
        
        lblSolan.Visible = False
        txtSolan.Visible = False
        

 
    Else
        frmKy.Height = 1700
        Frame2.Top = 2000

 
    End If
    
    cmbQuy.Visible = False
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
        
    SetControlCaption Me, "frmPeriod"
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout01TBVMT", Err.Number, Err.Description
    
End Sub

'******************************
'Description: SetValueToList procedure list subfolders in
'             the datafiles folder and add names to list.
'******************************
Private Sub SetValueToList(strId As String)
    On Error GoTo ErrHandle
    
    Dim fldList() As String
    Dim tempValue As Variant
    Dim strDataFileName As String
    Dim i As Integer
        
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_DM_NNKD.xml"))
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_NNKD.xml")
    ' Lay danh muc loai hoa don
    ' 15/11/2010
    i = 0
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
        cboNganhKD.Clear
        For Each xmlNode In xmlNodeListItems
            fldList = Split(GetAttribute(xmlNode, "Value"), "###")
            If strId = "11" Then
                If fldList(0) = "01A_TNDN" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "12" Then
                If fldList(0) = "01B_TNDN" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "01" Then
                If fldList(0) = "01_GTGT" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "05" Then
                If fldList(0) = "01_TTDB" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "03" Then
                If fldList(0) = "03_TNDN" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "73" Then
                If fldList(0) = "02_TNDN" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "98" Or strId = "92" Then
                If fldList(0) = "01A_TNDN_DK" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "93" Then
                If fldList(0) = "02_TNDN_DK" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            ElseIf strId = "89" Then
                If fldList(0) = "02_TAIN_DK" Then
                    cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                    cboNganhKD.ItemData(i) = Val(fldList(1))
                    i = i + 1
                End If
            End If
        Next
    End If
   
  
    If cboNganhKD.listcount > 0 Then _
        cboNganhKD.ListIndex = 0
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetValueToList", Err.Number, Err.Description
End Sub



' setup layout cho cac to khai 01A/TNDN, 01B/TNDN, 01_GTGT, 02_GTGT, 03_GTGT, 02/TAIN, 01/TD-GTGT
Private Sub SetupLayoutGTGT(strKieuKy As String, strIdToKhai As String)
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
'    frmKy.Height = 1800
'    Frame2.Top = 2100
    frmKy.Height = 1400
    Frame2.Top = 1700
    Select Case strKieuKy
        Case KIEU_KY_THANG
            Set lblMonth.Container = frmKy
            lblMonth.Top = 200
            lblMonth.Left = 960
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 150
            txtMonth.Left = 1530
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = 2310
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = 2730
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 600
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 950
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 950
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
            If strIdToKhai = "01" Or strIdToKhai = "11" Or strIdToKhai = "12" Then
                frmKy.Height = 2100
                Frame2.Top = 2400
                Set lblNganhKD.Container = frmKy
                lblNganhKD.Top = 1300
                lblNganhKD.Left = 120
                
                
                Set cboNganhKD.Container = frmKy
                cboNganhKD.Top = 1600
                cboNganhKD.Left = 120
                ' set gia tri nganh nghe kinh doanh cho combo
                SetValueToList strIdToKhai
            End If
'            Set fpsNgaykhaiBS.Container = frmKy
'            fpsNgaykhaiBS.Top = 1250
'            fpsNgaykhaiBS.Left = 960
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
        
        Case KIEU_KY_QUY
            Set lblQuy.Container = frmKy
            lblQuy.Top = 200
            lblQuy.Left = 1050
            
            Set cmbQuy.Container = frmKy
            cmbQuy.Top = 150
            cmbQuy.Left = 1440
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = 2220
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = 2640
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 650
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 900
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 850
            txtSolan.Left = 3400
            lblSolan.Visible = False
            txtSolan.Visible = False
            
    
            
            ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
            If strIdToKhai = "01" Or strIdToKhai = "11" Or strIdToKhai = "12" Then
                frmKy.Height = 2100
                Frame2.Top = 2400
                Set lblNganhKD.Container = frmKy
                lblNganhKD.Top = 1300
                lblNganhKD.Left = 120
                
                
                Set cboNganhKD.Container = frmKy
                cboNganhKD.Top = 1600
                cboNganhKD.Left = 120
                ' set gia tri nganh nghe kinh doanh cho combo
                SetValueToList strIdToKhai
            End If
                        
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
                SetControlCaption Me, "frmPeriodBCTC"
            Else
                SetControlCaption Me, "frmPeriodQuy"
            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            cmdOK.Top = 1500
            cmdClose.Top = cmdOK.Top
            Me.Height = 2040
            Me.Width = 4905
            
        Case KIEU_KY_NAM
            
            Set lblYear.Container = frmKy
            lblYear.Top = 200
            lblYear.Left = (frmKy.Width - txtYear.Width) / 2 - 100
            
            Set txtYear.Container = frmKy
            txtYear.Top = 150
            txtYear.Left = lblYear.Left + lblYear.Width + 50
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 650
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 950
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 900
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 850
            txtSolan.Left = 3400
            
            lblSolan.Visible = False
            txtSolan.Visible = False

            
'            If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
'                SetControlCaption Me, "frmPeriodBCTC"
'            Else
                SetControlCaption Me, "frmPeriodQuy"
'            End If
            'SetControlCaption Me, "frmPeriodQuy"
            
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            txtMonth.Visible = False
            cmbQuy.Visible = False
            
        Case KIEU_KY_NGAY_NAM
            Me.Height = 3285
            Me.Width = 5305
            frmKy.Height = 1665
            Frame2.Top = 1815
    
            Set lblYear.Container = frmKy
            lblYear.Top = 300
            lblYear.Left = 120
            
            Set txtYear.Container = frmKy
            txtYear.Top = 240
            txtYear.Left = 1000 '1200
            
            Set lblNgayDau.Container = frmKy
            lblNgayDau.Top = 630
            lblNgayDau.Left = 120
            
            Set txtNgayDau.Container = frmKy
            txtNgayDau.Top = 600
            txtNgayDau.Left = 1000 '1200
            
            Set lblNgayCuoi.Container = frmKy
            lblNgayCuoi.Top = 630
            lblNgayCuoi.Left = 2600 '2400
            
            Set txtNgayCuoi.Container = frmKy
            txtNgayCuoi.Top = 600
            txtNgayCuoi.Left = 3480
            
            
            Set OptChinhthuc.Container = frmKy
            OptChinhthuc.Top = 1000
            OptChinhthuc.Left = 960
            
            Set OptBosung.Container = frmKy
            OptBosung.Top = 1250
            OptBosung.Left = 960
            
            Set lblSolan.Container = frmKy
            lblSolan.Top = 1250
            lblSolan.Left = 3000
            Set txtSolan.Container = frmKy
            txtSolan.Top = 1250
            txtSolan.Left = 3480
            lblSolan.Visible = False
            txtSolan.Visible = False
            
            
            SetControlCaption Me, "frmPeriodQuy"
            
            txtMonth.Visible = False
            cmbQuy.Visible = False

        Case KIEU_KY_NGAY_THANG
            Set lblNgay.Container = frmKy
            lblNgay.Top = 480
            lblNgay.Left = 240
            
            Set txtDay.Container = frmKy
            txtDay.Top = 450
            txtDay.Left = 840
            
            Set lblMonth.Container = frmKy
            lblMonth.Top = 480
            lblMonth.Left = 1560
            
            Set txtMonth.Container = frmKy
            txtMonth.Top = 450
            txtMonth.Left = 2160
            
            Set lblYear.Container = frmKy
            lblYear.Top = 480
            lblYear.Left = 2880
            
            Set txtYear.Container = frmKy
            txtYear.Top = 450
            txtYear.Left = 3360
            
            SetControlCaption Me, "frmPeriod"
   
            cmbQuy.Visible = False
            txtNgayDau.Visible = False
            txtNgayCuoi.Visible = False
            
            
    End Select
    strKHBS = "TKCT"
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutGTGT", Err.Number, Err.Description
    
End Sub


Private Sub SetActiveValueKHBS()
'    Dim lCtrl As Long
    Dim varMenuId As String
    varMenuId = GetAttribute(TAX_Utilities_v1.NodeValidity.parentNode, "ID")
'    For lCtrl = 1 To fpSpread1.MaxRows
'        fpSpread1.Row = lCtrl
'        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(lCtrl - 1), "Active", "0"
'    Next lCtrl
    If varMenuId = "02" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(11), "Active", 1
    ElseIf varMenuId = "01" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(7), "Active", 1
    ElseIf varMenuId = "04" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(1), "Active", 1
'    ElseIf varMenuId = "95" Then
'        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(3), "Active", 1
    ElseIf varMenuId = "73" Or varMenuId = "88" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(1), "Active", 1
    ElseIf varMenuId = "71" Or varMenuId = "90" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(2), "Active", 1
    ElseIf varMenuId = "85" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(1), "Active", 1
    ElseIf varMenuId = "12" Or varMenuId = "11" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(3), "Active", 1
    ElseIf varMenuId = "06" Or varMenuId = "70" Or varMenuId = "77" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(1), "Active", 1
    ElseIf varMenuId = "05" Or varMenuId = "80" Or varMenuId = "89" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(3), "Active", 1
    ElseIf varMenuId = "86" Or varMenuId = "87" Or varMenuId = "72" Or varMenuId = "81" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(1), "Active", 1
    ElseIf varMenuId = "03" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(15), "Active", 1
    ElseIf varMenuId = "83" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(4), "Active", 1
    ElseIf varMenuId = "96" Or varMenuId = "94" Or varMenuId = "98" Or varMenuId = "99" Or varMenuId = "97" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(2), "Active", 1
    ElseIf varMenuId = "92" Or varMenuId = "93" Or varMenuId = "82" Then
        SetAttribute TAX_Utilities_v1.NodeValidity.childNodes(2), "Active", 1
    End If
End Sub


Private Sub formatPrefix(strDate As String, strarrdate() As String)
    Dim lstrdate As Variant
    Dim i As Integer
    i = 0
    strarrdate = Split(strDate, "/")
    For Each lstrdate In strarrdate
        If Len(Trim(lstrdate)) < 2 Then
                strarrdate(i) = "0" & Trim(lstrdate)
         End If
         i = i + 1
    Next
End Sub


Private Sub SetupLayout08TNCN()
    On Error GoTo ErrorHandle
    
    Me.Height = 3285
    Me.Width = 4905
    
    'frmKy.Height = 1300
    Set chkTKQuy.Container = frmKy
    chkTKQuy.Top = 200
    chkTKQuy.Left = 120
    chkTKQuy.value = 1
    chkTuThangDenThang.value = 0
    
    
    Set chkTuThangDenThang.Container = frmKy
    chkTuThangDenThang.Top = 200
    chkTuThangDenThang.Left = 2000
    
    
    
    
    Set lblQuy.Container = frmKy
    lblQuy.Top = 570
    lblQuy.Left = 1360
    
    Set cmbQuy.Container = frmKy
    cmbQuy.Top = 540
    cmbQuy.Left = 1930
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2710
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 3130
    
    txtMonth.Visible = False
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
    ' Set lai max lengh cho to khai tu thang den thang
    txtNgayDau.MaxLength = 7
    txtNgayCuoi.MaxLength = 7

    frmKy.Height = 1600
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 900
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1200
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1200
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1200
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    strKHBS = "TKCT"
    strQuy = "TK_QUY"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout08TNCN", Err.Number, Err.Description
    
End Sub


Private Sub SetLayoutToKhaiThangQuy()
    On Error GoTo ErrorHandle
    
    strLoaiSacThue = "ToKhaiGTGT"
    
    Me.Height = 3285
    Me.Width = 4905
    
    Frame2.Top = 1950
    
    'frmKy.Height = 1300
    Set chkTKQuy.Container = frmKy
    chkTKQuy.Top = 200
    chkTKQuy.Left = 2500
    chkTKQuy.value = 0
    chkTkhaiThang.value = 1
    
    
    Set chkTkhaiThang.Container = frmKy
    chkTkhaiThang.Top = 200
    chkTkhaiThang.Left = 120
    
    
    
    
    Set lblMonth.Container = frmKy
    lblMonth.Top = 570
    lblMonth.Left = 960
    
    Set txtMonth.Container = frmKy
    txtMonth.Top = 540
    txtMonth.Left = 1530
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2310
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 2730
        

    frmKy.Height = 1600
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 900
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1200
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1200
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1200
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    strKHBS = "TKCT"
    strQuy = "TK_THANG"
    
     ' to khai 01/GTGT co them danh muc nganh nghe kinh doanh
            If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
                frmKy.Height = 2400
                Frame2.Top = 2700
                Set lblNganhKD.Container = frmKy
                lblNganhKD.Top = 1600
                lblNganhKD.Left = 120
                
                
                Set cboNganhKD.Container = frmKy
                cboNganhKD.Top = 1900
                cboNganhKD.Left = 120
                ' set gia tri nganh nghe kinh doanh cho combo
                SetValueToList GetAttribute(TAX_Utilities_v1.NodeMenu, "ID")
            End If
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetLayoutToKhaiThangQuy", Err.Number, Err.Description
    
End Sub


Private Sub setValueDefault()
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "13" Or GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "18" Then
        chkTkhaiThang.value = 0
        chkTKQuy.value = 1
    Else
        chkTkhaiThang.value = 1
    End If
    chkTKLanPS.value = 0
End Sub


Private Sub SetupLayoutBC26()
    On Error GoTo ErrorHandle
    
    strLoaiSacThue = "BC26"
    
    Me.Height = 3500
    Me.Width = 4905

    frmKy.Height = 1100
    Set chkTKQuy.Container = frmKy
    chkTKQuy.Top = 200
    chkTKQuy.Left = 2500
'    chkTKQuy.value = 0
'    chkTkhaiThang.value = 1
'
    Set chkTkhaiThang.Container = frmKy
    chkTkhaiThang.Top = 200
    chkTkhaiThang.Left = 120
'
'    Set lblMonth.Container = frmKy
'    lblMonth.Top = 570
'    lblMonth.Left = 960
'
'    Set txtMonth.Container = frmKy
'    txtMonth.Top = 540
'    txtMonth.Left = 1530
'
'    Set lblYear.Container = frmKy
'    lblYear.Top = 570
'    lblYear.Left = 2310
'
'    Set txtYear.Container = frmKy
'    txtYear.Top = 540
'    txtYear.Left = 2730
'
    Set lblNgayDau.Container = frmKy
    lblNgayDau.Top = 930
    lblNgayDau.Left = 120

    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 900
    txtNgayDau.Left = 1000 '1200
    'txtNgayDau.Locked = True

    Set lblNgayCuoi.Container = frmKy
    lblNgayCuoi.Top = 930
    lblNgayCuoi.Left = 2600 '2400

    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 900
    txtNgayCuoi.Left = 3480
    strQuy = "TK_QUY"
    chkTkhaiThang.value = 0
    chkTKQuy.value = 1
            
    Set lblQuy.Container = frmKy
    lblQuy.Top = 570
    lblQuy.Left = 960
    
    Set cmbQuy.Container = frmKy
    cmbQuy.Top = 540
    cmbQuy.Left = 1530
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2310
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 2730
    
    SetControlCaption Me, "frmPeriod"

    cmbQuy.Visible = True
    lblQuy.Visible = True
    
    lblMonth.Visible = False
    txtMonth.Visible = False
    txtNgayDau.Visible = True
    txtNgayCuoi.Visible = True
    lblNgayDau.Visible = True
    lblNgayCuoi.Visible = True
    frmKy.Height = 1300
    Frame2.Top = 1600
    
'    Set lblQuy.Container = frmKy
'    lblQuy.Top = 300
'    lblQuy.Left = 120
'    lblQuy.caption = "Ky`"
'    SetControlCaption Me, "frmPeriodHY"
'    'lblCaption.caption = GetAttribute(GetMessageCellById("0183"), "Msg")
'    Set cmbQuy.Container = frmKy
'    cmbQuy.Top = 240
'    cmbQuy.Left = 1000
'
'    lblYear.Visible = False
'
'
'    Set txtYear.Container = frmKy
'    txtYear.Top = 240
'    txtYear.Left = 1600
'
'    Set lblNgayDau.Container = frmKy
'    lblNgayDau.Top = 630
'    lblNgayDau.Left = 120
'
'    Set txtNgayDau.Container = frmKy
'    txtNgayDau.Top = 600
'    txtNgayDau.Left = 1000 '1200
'    'txtNgayDau.Locked = True
'
'    Set lblNgayCuoi.Container = frmKy
'    lblNgayCuoi.Top = 630
'    lblNgayCuoi.Left = 2600 '2400
'
'    Set txtNgayCuoi.Container = frmKy
'    txtNgayCuoi.Top = 600
'    txtNgayCuoi.Left = 3480
'
'    'SetControlCaption Me, "frmPeriodQuy"
'
'    txtMonth.Visible = False
'    'cmbQuy.Visible = False
'     ' end

    SetControlCaption Me, "frmPeriod"
   
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2

    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutBC26", Err.Number, Err.Description
    
End Sub


' Set up layout cho to 02/BVMT
Private Sub SetupLayout02BVMT()
    On Error GoTo ErrorHandle
    
    Me.Height = 3385
    Me.Width = 4905
    frmKy.Height = 1840
    Frame2.Top = 2050

    Set lblYear.Container = frmKy
    lblYear.Top = 300
    lblYear.Left = 120
    
    Set txtYear.Container = frmKy
    txtYear.Top = 240
    txtYear.Left = 1000 '1200
    
    Set lblTuThang.Container = frmKy
    lblTuThang.Top = 630
    lblTuThang.Left = 120
    
    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 600
    txtNgayDau.Left = 1000 '1200
    
    Set lblDenThang.Container = frmKy
    lblDenThang.Top = 630
    lblDenThang.Left = 2600 '2400
    
    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 600
    txtNgayCuoi.Left = 3480
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 1050
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1400
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1400
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1400
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
        
    SetControlCaption Me, "frmPeriodQuy"
    
    txtMonth.Visible = False
    cmbQuy.Visible = False
        
    strKHBS = "TKCT"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2 - 400
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout02BVMT", Err.Number, Err.Description
    
End Sub

Private Sub SetupLayout01TTS()
    On Error GoTo ErrorHandle
    
     Dim m, Y, d As Integer
    Dim dTem, dtem1 As Date
    Dim q As Quy
    
    m = month(Date)
    Y = Year(Date)
    
    Me.Height = 3285
    Me.Width = 4905
    
    'frmKy.Height = 1300
    Set chkTKQuy.Container = frmKy
    chkTKQuy.Top = 200
    chkTKQuy.Left = 120
    chkTKQuy.value = 1
    chkTuThangDenThang.value = 0
    
     ' Set gia tri mac dinh cho Quy
    q = GetQuyHienTai(iNgayTaiChinh, iThangTaiChinh)

    If q.q = 1 Then
        q.q = 4
        q.Y = q.Y - 1
    Else
        q.q = q.q - 1
    End If

    cmbQuy.ListIndex = q.q - 1
    txtYear.Text = q.Y
    
    Set chkTuThangDenThang.Container = frmKy
    chkTuThangDenThang.Top = 200
    chkTuThangDenThang.Left = 2000
    
    
    
    
    Set lblQuy.Container = frmKy
    lblQuy.Top = 570
    lblQuy.Left = 1360
    
    Set cmbQuy.Container = frmKy
    cmbQuy.Top = 540
    cmbQuy.Left = 1930
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2710
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 3130
    
    txtMonth.Visible = False
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
    ' Set lai max lengh cho to khai tu thang den thang
    txtNgayDau.MaxLength = 7
    txtNgayCuoi.MaxLength = 7

    frmKy.Height = 1600
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 900
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1200
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1200
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1200
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    strKHBS = "TKCT"
    strQuy = "TK_QUY"
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout08TNCN", Err.Number, Err.Description
    
End Sub


Private Sub SetupLayout16TH()
    On Error GoTo ErrorHandle
    Me.Height = 3285
    Me.Width = 4905
    
    frmKy.Height = 1200
    
    Set lblYear.Container = frmKy
    lblYear.Top = 380
    lblYear.Left = (frmKy.Width - txtYear.Width) / 2 - 100
    
    Set txtYear.Container = frmKy
    txtYear.Top = 300
    txtYear.Left = lblYear.Left + lblYear.Width + 50 '1200
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 780
    lblSolan.Left = (frmKy.Width - txtYear.Width) / 2 - 100
   
    
    Set txtSolan.Container = frmKy
    txtSolan.Top = 780
    txtSolan.Left = lblYear.Left + lblYear.Width + 50 '1200
    
    
        
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "PopID") = "101" Then
        SetControlCaption Me, "frmPeriodBCTC"
    Else
        SetControlCaption Me, "frmPeriodQuy"
    End If
    'SetControlCaption Me, "frmPeriodQuy"
    
    txtNgayDau.Visible = False
    txtNgayCuoi.Visible = False
    txtMonth.Visible = False
    cmbQuy.Visible = False
    
    strKHBS = "TKCT"
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayout16TH", Err.Number, Err.Description
    
End Sub

' to khai 04/GTGT
Private Sub SetLayoutToKhaiThangQuyLanPS()
    On Error GoTo ErrorHandle
    
    strLoaiSacThue = "ToKhaiGTGT"
    
    Me.Height = 3285
    Me.Width = 4905
    
    Frame2.Top = 1950
    
    'frmKy.Height = 1300
    Set chkTKQuy.Container = frmKy
    chkTKQuy.Top = 200
    chkTKQuy.Left = 1500
    chkTKQuy.value = 0
    chkTkhaiThang.value = 1
    
    
    Set chkTkhaiThang.Container = frmKy
    chkTkhaiThang.Top = 200
    chkTkhaiThang.Left = 50
    
    Set chkTKLanPS.Container = frmKy
    chkTKLanPS.Top = 200
    chkTKLanPS.Left = 2850
    
    Set lblNgay.Container = frmKy
    lblNgay.Top = 570
    lblNgay.Left = 120
    lblNgay.Visible = False

    Set txtDay.Container = frmKy
    txtDay.Top = 540
    txtDay.Left = 700
    txtDay.Visible = False
    
    
    Set lblMonth.Container = frmKy
    lblMonth.Top = 570
    lblMonth.Left = 1360
    
    Set txtMonth.Container = frmKy
    txtMonth.Top = 540
    txtMonth.Left = 1930
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2710
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 3130
        

    frmKy.Height = 1600
    
    Set OptChinhthuc.Container = frmKy
    OptChinhthuc.Top = 900
    OptChinhthuc.Left = 960
    
    Set OptBosung.Container = frmKy
    OptBosung.Top = 1200
    OptBosung.Left = 960
    
    Set lblSolan.Container = frmKy
    lblSolan.Top = 1200
    lblSolan.Left = 3000
    Set txtSolan.Container = frmKy
    txtSolan.Top = 1200
    txtSolan.Left = 3400
    
    lblSolan.Visible = False
    txtSolan.Visible = False
    strKHBS = "TKCT"
    strQuy = "TK_THANG"
    
        
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetLayoutToKhaiThangQuyLanPS", Err.Number, Err.Description
    
End Sub



Private Sub SetupLayoutBC01()
    On Error GoTo ErrorHandle
    
    strLoaiSacThue = "BC01"
    
    Me.Height = 3500
    Me.Width = 4905

    frmKy.Height = 1100
    Set chkTKQuy.Container = frmKy
    chkTKQuy.Top = 200
    chkTKQuy.Left = 2500
    Set chkTKKy.Container = frmKy
    chkTKKy.Top = 200
    chkTKKy.Left = 120
    Set lblNgayDau.Container = frmKy
    lblNgayDau.Top = 930
    lblNgayDau.Left = 120

    Set txtNgayDau.Container = frmKy
    txtNgayDau.Top = 900
    txtNgayDau.Left = 1000 '1200
    'txtNgayDau.Locked = True

    Set lblNgayCuoi.Container = frmKy
    lblNgayCuoi.Top = 930
    lblNgayCuoi.Left = 2600 '2400

    Set txtNgayCuoi.Container = frmKy
    txtNgayCuoi.Top = 900
    txtNgayCuoi.Left = 3480
    strQuy = "TK_QUY"
    chkTKKy.value = 0
    chkTKQuy.value = 1
            
    Set lblQuy.Container = frmKy
    lblQuy.Top = 570
    lblQuy.Left = 960
    
    Set cmbQuy.Container = frmKy
    cmbQuy.Top = 540
    cmbQuy.Left = 1530
    
    Set lblYear.Container = frmKy
    lblYear.Top = 570
    lblYear.Left = 2310
    
    Set txtYear.Container = frmKy
    txtYear.Top = 540
    txtYear.Left = 2730
    
    SetControlCaption Me, "frmPeriod"

    cmbQuy.Visible = True
    lblQuy.Visible = True
    
    lblMonth.Visible = False
    txtMonth.Visible = False
    txtNgayDau.Visible = True
    txtNgayCuoi.Visible = True
    lblNgayDau.Visible = True
    lblNgayCuoi.Visible = True
    frmKy.Height = 1300
    Frame2.Top = 1600
    
    SetControlCaption Me, "frmPeriod"
   
    Me.Top = (frmSystem.ScaleHeight - Me.ScaleHeight) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2

    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "SetupLayoutBC01", Err.Number, Err.Description
    
End Sub

' loaiKyKK = 1, to khai thang, 0 to khai lan xuat ban
Private Sub SetValueToListDK(loaiKyKK As String)
    On Error GoTo ErrHandle
    
    Dim fldList() As String
    Dim tempValue As Variant
    Dim strDataFileName As String
    Dim i As Integer
        
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_DM_NNKD.xml"))
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_NNKD.xml")
    ' Lay danh muc loai hoa don
    ' 17/03/2014
    i = 0
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
        cboNganhKD.Clear
        For Each xmlNode In xmlNodeListItems
            fldList = Split(GetAttribute(xmlNode, "Value"), "###")
            If fldList(0) = "01A_TNDN_DK" Then
                If Trim$(loaiKyKK) = "1" Then
                    If Val(fldList(1)) = 2 Then
                        cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                        cboNganhKD.ItemData(i) = Val(fldList(1))
                        i = i + 1
                    End If
                Else
                    If Val(fldList(1)) <> 2 Then
                        cboNganhKD.AddItem TAX_Utilities_v1.Convert(fldList(2), UNICODE, TCVN)
                        cboNganhKD.ItemData(i) = Val(fldList(1))
                        i = i + 1
                    End If
                End If
            End If
        Next
    End If
   
  
    If cboNganhKD.listcount > 0 Then _
        cboNganhKD.ListIndex = 0
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetValueToListDK", Err.Number, Err.Description
End Sub

' tenFileTK: truyen datafile cua to khai
Private Function checkKyKKTrung(ByVal tenFileTK As String, ByVal tuThangKK As String, ByVal denThangKK As String, ByVal NamKK As String) As Boolean
    Dim isTrung As Boolean
    Dim lngIndex As Integer
    Dim arrTemp() As String
    Dim tuThang1 As String
    Dim denThang1 As String
    Dim chenhLech1 As Integer
    Dim chenhLech2 As Integer
    Dim chenhLech3 As Integer
    Dim chenhLech4 As Integer
    On Error GoTo ErrHandle
    ' load danh muc file trong folder
    LoadXMLFileNames
    
    For lngIndex = 0 To UBound(arrStrXMLFileNames)
        ' to khai chinh thuc
        If Len(arrStrXMLFileNames(lngIndex)) > 18 Then
        ' kiem tra 19: YYYY_MMYYYY_MMYYYY
            If tenFileTK = Mid$(arrStrXMLFileNames(lngIndex), 1, Len(arrStrXMLFileNames(lngIndex)) - 19) Then
                arrTemp = Split(Right(arrStrXMLFileNames(lngIndex), 13), "_")
                tuThang1 = arrTemp(0)
                denThang1 = arrTemp(1)
                ' kiem tra neu tu thang1 , den thang 1 = tu thang KK , den thang KK thi tra ve false
                ' truong hop tu thang, den than nam trong khoang
                chenhLech1 = DateDiff("M", format(tuThangKK, "mm/yyyy"), format(Left(tuThang1, 2) & "/" & Right(tuThang1, 4), "mm/yyyy"))
                chenhLech2 = DateDiff("M", format(Left(denThang1, 2) & "/" & Right(denThang1, 4), "mm/yyyy"), format(denThangKK, "mm/yyyy"))
                If chenhLech1 = 0 And chenhLech2 = 0 Then
                    isTrung = False
                    Exit For
                ElseIf chenhLech1 * chenhLech2 > 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 18), 4) = NamKK Then
                    ' kiem tra them nam ke khai trung thi moi bat
                    isTrung = True
                    Exit For
                End If
                
                
                ' truong hop tu thang 1 nam trong khoang
                chenhLech1 = DateDiff("M", format(tuThangKK, "mm/yyyy"), format(Left(tuThang1, 2) & "/" & Right(tuThang1, 4), "mm/yyyy"))
                chenhLech2 = DateDiff("M", format(tuThangKK, "mm/yyyy"), format(Left(denThang1, 2) & "/" & Right(denThang1, 4), "mm/yyyy"))
                If chenhLech1 * chenhLech2 <= 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 18), 4) = NamKK Then
                    isTrung = True
                    Exit For
                End If
                ' truong hop den thang nam trong khoang
                chenhLech1 = DateDiff("M", format(denThangKK, "mm/yyyy"), format(Left(tuThang1, 2) & "/" & Right(tuThang1, 4), "mm/yyyy"))
                chenhLech2 = DateDiff("M", format(denThangKK, "mm/yyyy"), format(Left(denThang1, 2) & "/" & Right(denThang1, 4), "mm/yyyy"))
                If chenhLech1 * chenhLech2 <= 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 18), 4) = NamKK Then
                    isTrung = True
                    Exit For
                End If
                
                
                
            End If
        End If
    Next lngIndex
    checkKyKKTrung = isTrung
    Exit Function
ErrHandle:
    checkKyKKTrung = False
    SaveErrorLog "frmPeriod", "checkKyKKTrung", Err.Number, Err.Description
End Function


Private Sub LoadXMLFileNames()
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As file
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v1.DataFolder)).Files
        If Right$(fle.Name, 4) = ".xml" Then
            ReDim Preserve arrStrXMLFileNames(lngIndex)
            arrStrXMLFileNames(lngIndex) = Mid$(fle.Name, 1, Len(fle.Name) - 4)
            lngIndex = lngIndex + 1
        End If
    Next
End Sub



' tenFileTK: truyen datafile cua to khai
Private Function checkKyKKTrungNgay(ByVal tenFileTK As String, ByVal tuNgayKK As String, ByVal denNgayKK As String, ByVal NamKK As String) As Boolean
    Dim isTrung As Boolean
    Dim lngIndex As Integer
    Dim arrTemp() As String
    Dim tuNgay1 As String
    Dim denNgay1 As String
    Dim chenhLech1 As Long
    Dim chenhLech2 As Long
    Dim chenhLech3 As Long
    Dim chenhLech4 As Long
    On Error GoTo ErrHandle
    ' load danh muc file trong folder
    LoadXMLFileNames

    For lngIndex = 0 To UBound(arrStrXMLFileNames)
        ' to khai chinh thuc
        If Len(arrStrXMLFileNames(lngIndex)) > 23 Then
        ' kiem tra 19: YYYY_MMYYYY_MMYYYY
            If tenFileTK = Mid$(arrStrXMLFileNames(lngIndex), 1, Len(arrStrXMLFileNames(lngIndex)) - 23) Then
                arrTemp = Split(Right(arrStrXMLFileNames(lngIndex), 17), "_")
                tuNgay1 = arrTemp(0)
                denNgay1 = arrTemp(1)
                ' kiem tra neu tu thang1 , den thang 1 = tu thang KK , den thang KK thi tra ve false
                ' truong hop tu thang, den than nam trong khoang
                chenhLech1 = DateDiff("D", format(tuNgayKK, "dd/mm/yyyy"), format(Left(tuNgay1, 2) & "/" & Mid(tuNgay1, 3, 2) & "/" & Right(tuNgay1, 4), "dd/mm/yyyy"))
                chenhLech2 = DateDiff("D", format(Left(denNgay1, 2) & "/" & Mid$(denNgay1, 3, 2) & "/" & Right(denNgay1, 4), "dd/mm/yyyy"), format(denNgayKK, "dd/mm/yyyy"))
                If chenhLech1 <> 0 Then
                    chenhLech1 = chenhLech1 / Abs(chenhLech1)
                End If
                If chenhLech2 <> 0 Then
                    chenhLech2 = chenhLech2 / Abs(chenhLech2)
                End If
                If chenhLech1 = 0 And chenhLech2 = 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 22), 4) = NamKK Then
                    isTrung = False
                    Exit For
                ElseIf chenhLech1 * chenhLech2 > 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 22), 4) = NamKK Then
                    isTrung = True
                    Exit For
                End If


                ' truong hop tu thang 1 nam trong khoang
                chenhLech1 = DateDiff("D", format(tuNgayKK, "dd/mm/yyyy"), format(Left(tuNgay1, 2) & "/" & Mid$(tuNgay1, 3, 2) & "/" & Right(tuNgay1, 4), "dd/mm/yyyy"))
                chenhLech2 = DateDiff("D", format(tuNgayKK, "dd/mm/yyyy"), format(Left(denNgay1, 2) & "/" & Mid$(denNgay1, 3, 2) & "/" & Right(denNgay1, 4), "dd/mm/yyyy"))
                If chenhLech1 <> 0 Then
                    chenhLech1 = chenhLech1 / Abs(chenhLech1)
                End If
                If chenhLech2 <> 0 Then
                    chenhLech2 = chenhLech2 / Abs(chenhLech2)
                End If
                If chenhLech1 * chenhLech2 <= 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 22), 4) = NamKK Then
                    isTrung = True
                    Exit For
                End If
                ' truong hop den thang nam trong khoang
                chenhLech1 = DateDiff("D", format(denNgayKK, "dd/mm/yyyy"), format(Left(tuNgay1, 2) & "/" & Mid$(tuNgay1, 3, 2) & "/" & Right(tuNgay1, 4), "dd/mm/yyyy"))
                chenhLech2 = DateDiff("D", format(denNgayKK, "dd/mm/yyyy"), format(Left(denNgay1, 2) & "/" & Mid$(denNgay1, 3, 2) & "/" & Right(denNgay1, 4), "dd/mm/yyyy"))
                If chenhLech1 <> 0 Then
                    chenhLech1 = chenhLech1 / Abs(chenhLech1)
                End If
                If chenhLech2 <> 0 Then
                    chenhLech2 = chenhLech2 / Abs(chenhLech2)
                End If
                If chenhLech1 * chenhLech2 <= 0 And Left$(Right$(arrStrXMLFileNames(lngIndex), 22), 4) = NamKK Then
                    isTrung = True
                    Exit For
                End If



            End If
        End If
    Next lngIndex
    checkKyKKTrungNgay = isTrung
    Exit Function
ErrHandle:
    checkKyKKTrungNgay = False
    SaveErrorLog "frmPeriod", "checkKyKKTrungNgay", Err.Number, Err.Description
End Function
