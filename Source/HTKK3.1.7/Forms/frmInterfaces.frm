VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmInterfaces 
   AutoRedraw      =   -1  'True
   Caption         =   "H� tr� k� khai - Phi�n b�n 2.5.0"
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "DS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterfaces.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmInterfaces"
   LockControls    =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   10800
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   17295
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   7560
         TabIndex        =   20
         Top             =   195
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.ComboBox Cb_seach 
         Height          =   315
         ItemData        =   "frmInterfaces.frx":164A
         Left            =   3240
         List            =   "frmInterfaces.frx":1657
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_Seach 
         Caption         =   "T�m ki�m"
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin FPUSpreadADO.fpSpread txt_Seach 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   3000
         _Version        =   458752
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   1
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterfaces.frx":1676
      End
      Begin VB.Label Lbload 
         Caption         =   "�ang x� l� ..."
         BeginProperty Font 
            Name            =   "DS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   6120
         TabIndex        =   19
         Top             =   130
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      Left            =   0
      TabIndex        =   8
      Top             =   405
      Width           =   11580
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   420
         Top             =   1980
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   6510
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   11385
         _Version        =   458752
         _ExtentX        =   20082
         _ExtentY        =   11483
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterfaces.frx":1993
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   0
      TabIndex        =   9
      Top             =   7230
      Width           =   11535
      Begin VB.CommandButton Command1 
         Caption         =   "T�ng h�p KHBS"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoadToKhai 
         Caption         =   "T�i t� kh&ai"
         Height          =   375
         HelpContextID   =   81212
         Left            =   2745
         TabIndex        =   13
         Top             =   210
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Th�m ph� l�c"
         Height          =   375
         HelpContextID   =   81212
         Left            =   2745
         TabIndex        =   1
         Top             =   210
         Width           =   1140
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "��&ng"
         Height          =   375
         HelpContextID   =   81212
         Left            =   9945
         TabIndex        =   7
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Ghi"
         Height          =   375
         HelpContextID   =   81212
         Left            =   5175
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   81212
         Left            =   6360
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&X�a"
         Height          =   375
         HelpContextID   =   81212
         Left            =   1575
         TabIndex        =   5
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&K�t xu�t"
         Height          =   375
         HelpContextID   =   81212
         Left            =   8730
         TabIndex        =   6
         Top             =   210
         Width           =   1140
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "N&h�p l�i"
         Height          =   375
         HelpContextID   =   81212
         Left            =   3960
         TabIndex        =   2
         Top             =   210
         Width           =   1140
      End
      Begin VB.CommandButton cmdKiemTra 
         Caption         =   "Ki�m t&ra"
         Height          =   375
         HelpContextID   =   81212
         Left            =   7560
         TabIndex        =   14
         Top             =   210
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   150
         Width           =   2145
         WordWrap        =   -1  'True
      End
   End
   Begin FPUSpreadADO.fpSpread fpSpread2 
      Height          =   3780
      Left            =   6390
      TabIndex        =   12
      Top             =   2790
      Visible         =   0   'False
      Width           =   4560
      _Version        =   458752
      _ExtentX        =   8043
      _ExtentY        =   6667
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxRows         =   10
      SpreadDesigner  =   "frmInterfaces.frx":1C45
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Nh�p t� khai"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   30
      Width           =   3975
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmInterfaces
' Descriptions      : Loading interface for user preview, modify or adding data
'                   : use this form for all business process
'                   : Step 1 -> loading interface template from MS Excel file
'                   : Step 2 -> filling data
'                   : Step 3 -> allow modify data
'                   : Step 4 -> allow insert/ delete row
'                   : Step 5 -> update data (checking business rule before update)
'                   : Step 6 -> calling printting process (frmReports)
' Start date        : 10/10/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : htphuong
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :

Option Explicit


Private xmlDocumentInit()   As MSXML.DOMDocument
Private xmlDocumentStatus   As MSXML.DOMDocument
Private mOnLoad             As Boolean              ' mOnLoad = True when Form_Load process
Private mOnSetupData        As Boolean
Private mHeaderSheet        As Integer              ' save value of Header sheet (last sheet)
Private objTaxBusiness      As Object               ' private business object (cls001, cls002, cls003, ...)
Private mAdjustData         As Boolean              ' mAdjustData = True when user adjust data on interface



Private checkSoCT As Integer  ' Check so chi tieu =1 thieu chi tieu, =2 thua chi tieu, =3 khac so luong chi tieu
''' UpdateData description
''' Save data to Data Files, using save method of DOM object for save data to file
Private Function UpdateData(Optional blnSaveSession As Boolean = True) As Boolean
    On Error GoTo ErrorHandle
    Dim fso As New FileSystemObject
    '*********************************
    Dim xmlDom As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode, xmlNodeList As MSXML.IXMLDOMNodeList
    Dim strTenDoanhNghiep As String
    Dim clsConverter As New clsUnicodeTCVNConverter
    strTenDoanhNghiep = ""
    '*********************************
    
    Dim lSheet As Integer, lErrNumber As Long
    Dim strDataFileName As String
    
    For lSheet = 0 To TAX_Utilities_New.xmlDataCount
        
     If strKHBS = "TKBS" Then
        If GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = "0" Then
            strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
        Else
            If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Day") <> "1" Then
                If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "95" _
                Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "36" Then
                    If strQuy = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                    End If
                Else
                    strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
                If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "75" Then
                ' To khai 08/TNCN co to khai tu thang va to khai quy
                    If strQuy = "TK_TU_THANG" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                    End If
                 Else
                    strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                 End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") <> "1" Then
                    'Data file contain Day from and to.
                    If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "82" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                        & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                        & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                    End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
                    'Data file contain Day.
                    strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                    & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
            Else
                    'Data file not contain Day from and to.
                    strDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                    & TAX_Utilities_New.Year & ".xml"
            '*********************************
            End If
        End If
      Else
        If GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = "0" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
        Else
            If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Day") <> "1" Then
                If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "95" _
                Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "36" Then
                    If strQuy = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                    End If
                Else
                    strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
                If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "75" Then
                ' To khai 08/TNCN co to khai tu thang va to khai quy
                    If strQuy = "TK_TU_THANG" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                    End If
                ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "73" Then
                    ' To khai 02/TNDN
                    If strLoaiTKThang_PS = "TK_LANPS" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                    End If
                ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "68" Then
                    ' BC26
                    If strQuy = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_T" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                    End If
                Else
                    strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") <> "1" Then
                    'Data file contain Day from and to.
                    If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "82" Then
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                        & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                        & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                    End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
                    'Data file contain Day.
                    strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                    & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
            Else
                    'Data file not contain Day from and to.
                    strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                    & TAX_Utilities_New.Year & ".xml"
            '*********************************
            End If
        End If
      End If
      
        
        If TAX_Utilities_New.DataChanged And blnSaveSession Then
            If intDataSession >= 999 Then
                intDataSession = 0
            Else
                intDataSession = intDataSession + 1
            End If
            If intPrintingSession >= 999 Then
                intPrintingSession = 0
            Else
                intPrintingSession = intPrintingSession + 1
            End If
            If SaveSessionValueToFile(TAX_Utilities_New.DataFolder & "Session.dat") Then
                TAX_Utilities_New.DataChanged = False
            Else
                Exit Function
            End If
        End If
        '*********************************
        
        'TAX_Utilities_New.Data(CLng(lSheet)).save strDataFileName
        '*********************************
        If Val(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "Active")) >= 1 Then
            If fso.FileExists(strDataFileName) Then
                fso.GetFile(strDataFileName).Attributes = Normal
            End If
            TAX_Utilities_New.Data(CLng(lSheet)).save strDataFileName
        End If
        '*********************************
    Next
    'mAdjustData = False
    ResetAdjustData
    
    '
    If fso.FolderExists(GetAbsolutePath("..\DataFiles\" & strTaxIdString)) Then
       ' Load data header to DOM
        xmlDom.Load (GetAbsolutePath("..\DataFiles\" & strTaxIdString)) & "\Header_01.xml"
        ' Get Cell nodes
        Set xmlNodeList = xmlDom.getElementsByTagName("Cell")
        Set xmlNode = xmlNodeList(13)
        strTenDoanhNghiep = GetAttribute(xmlNode, "Value")
    End If
    
    'Set tax id to system caption
    frmSystem.lblUserInfo.caption = Mid$(strTaxIdString, 1, 10) & _
        IIf(Len(strTaxIdString) = 13, " - " & Mid$(strTaxIdString, 11, 3), "") & " : " & clsConverter.Convert(strTenDoanhNghiep, TCVN, UNICODE)
    
    Set xmlDom = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    '**********************************
    
    UpdateData = True
    Set fso = Nothing
    Exit Function
    
ErrorHandle:
    lErrNumber = Err.Number
    SaveErrorLog Me.Name, "UpdateData", Err.Number, Err.Description
    If lErrNumber = -2147024784 Then _
        DisplayMessage "0037", msOKOnly, miCriticalError
End Function

Private Function ImportExcel(ByVal strFileName As String) As Boolean
    Dim Y As Boolean, z As Boolean
    Dim Var As Variant
    Dim X As Integer, listcount As Integer, Handle As Integer
    Dim List(10) As String

    
    Dim strContentOfFile As String
    
    On Error GoTo DialogError
    
        ' Check if file is an Excel file and set result to x
    X = fpSpread1.IsExcelFile(strFileName)

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If X = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        Y = fpSpread2.GetExcelSheetList(strFileName, List, listcount, "C:\ILOGFILE.TXT", Handle, True)
        ' If received sheet list, tell user, import file,
        ' and set result to z
        If Y = True Then
            'MsgBox "Got sheet list.", , "Status"
            z = fpSpread2.ImportExcelSheet(Handle, 0)
            '
            
            
            
'            Dim checkvl As String
'            fpSpread2.sheet = 1
'            fpSpread2.Col = 1
'            fpSpread2.Row = 5
'            checkvl = fpSpread2.Text
'            If checkvl = Right(strFileName, 17) And "04-1/TNCN" = fpSpread1.SheetName Then
'                ImportExcel = True
'            ElseIf checkvl = Right(strFileName, 15) And "PL 01-1/GTGT" = fpSpread1.SheetName Then
'                ImportExcel = True
'            ElseIf checkvl = Right(strFileName, 16) And "PL 01-2/GTGT" = fpSpread1.SheetName Then
'                ImportExcel = True
'            Else
'                ImportExcel = False
'                DisplayMessage "0106", msOKOnly, miInformation
'            End If
            
            ImportExcel = True
            
            ' Tell user result based on T/F value of z
'            If z = True Then
'                MsgBox "Import complete.", , "Result"
'            Else
'                MsgBox "Import did not succeed.", , "Result"
'            End If
        Else
            ' Tell user cannot obtain sheet list
            DisplayMessage "0105", msOKOnly, miInformation
            ImportExcel = False
        End If
    Else
    ' Tell user file is not Excel file or is locked
       ' MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
        DisplayMessage "0105", msOKOnly, miInformation
        ImportExcel = False
    End If
    
DialogError:
    Me.Show
    Exit Function
ErrHandle:
    Me.Show
End Function

Private Function loadToKhai() As Boolean
    Dim Y As Boolean, z As Boolean
    Dim Var As Variant
    Dim X As Integer, listcount As Integer, Handle As Integer
    Dim List(10) As String

    Dim strFileName As String
    
    Dim strContentOfFile As String
    
    Dim mstFile As Variant, mstUD As Variant
    
    On Error GoTo DialogError
    fpSpread2.SheetCount = 1
    With CommonDialog1
        .CancelError = True
        .Filter = "File (UNICODE Font) (*.xls)|*.xls|File (TCVN3 Font) (*.xls)|*.xls|File (VNI Font) (*.xls)|*.xls|File (VIQR Font) (*.xls)|*.xls|File (VISCII Font) (*.xls)|*.xls"
        .FilterIndex = 1
        .DialogTitle = "Chon to khai de load vao chuong trinh"
        .ShowOpen
        strFileName = .FileName
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
    
    ' Check if file is an Excel file and set result to x
    X = fpSpread2.IsExcelFile(strFileName)
        
    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If X = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        Y = fpSpread2.GetExcelSheetList(strFileName, List, listcount, "C:\ILOGFILE.TXT", Handle, True)
        ' If received sheet list, tell user, import file,
        ' and set result to z
        If Y = True Then
        Dim i As Integer
            'MsgBox "Got sheet list.", , "Status"
            fpSpread2.SheetCount = listcount
            For i = 1 To listcount
                fpSpread2.sheet = i
                z = fpSpread2.ImportExcelSheet(Handle, i - 1)
            Next
            
            fpSpread2.Visible = False
            loadToKhai = True
        Else
            ' Tell user cannot obtain sheet list
            DisplayMessage "0105", msOKOnly, miInformation
            loadToKhai = False
        End If
    Else
        DisplayMessage "0105", msOKOnly, miInformation
        loadToKhai = False
    End If
    
DialogError:
    Me.Show
    Exit Function
ErrHandle:
    Me.Show

End Function

Private Sub moveData()
Dim value As String
Dim xmlDocument As New MSXML.DOMDocument
Dim xmlNode As MSXML.IXMLDOMNode

Dim i, count, count1, count2 As Long
Dim inc As Boolean
Dim colStart As Integer
Dim varMenuId As String

Dim lRow2s As Long
Dim incSession As Integer

On Error GoTo ErrHandle

incSession = 0

fpSpread1.EventEnabled(EventAllEvents) = False
    ' Truong hop them du lieu va xoa du lieu da ton tai
    If themXoaDuLieu Then
        ResetData
        ResetDataAndForm mCurrentSheet
    End If
    
' Lay ID cua Menu
varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")

fpSpread2.Visible = False
ProgressBar1.Visible = True
ProgressBar1.max = fpSpread2.MaxRows
ProgressBar1.value = 0
If Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_05A_TNCN.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 3 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_05B_TNCN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 4 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_TNCN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "42" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_02A_TNCN_BH.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "43" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_02A_TNCN_XS.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "44" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_06D_TNCN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "01" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_1_GTGT.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "01" And fpSpread1.ActiveSheet = 3 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_2_GTGT.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "02" Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_02_1_GTGT.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "14" Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\HHDL_05_TNDN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "05" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_1_TTDB.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "05" And fpSpread1.ActiveSheet = 3 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_2_TTDB.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "59" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_06_1_TNCN.xml"))
    colStart = 4
End If

Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
   fpSpread1.EventEnabled(EventAllEvents) = False
   fpSpread1.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row1").Item(0).Text)
   fpSpread2.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row2").Item(0).Text)
   fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
   count1 = Conversion.CInt(xmlDocument.getElementsByTagName("count").Item(0).Text)
   
    
    ' Truong hop them tiep du lieu
    Dim xmlSecionNode As MSXML.IXMLDOMNode
    Dim currentRow As Long
    Dim varData1, varData2 As Variant
    If themDuLieu Then
        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(0)
        'fpSpread1.Visible = False
        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
            If (xmlSecionNode.childNodes.length = 1) Then
                fpSpread1.sheet = mCurrentSheet
                fpSpread1.GetText colStart, fpSpread1.Row, varData1
                fpSpread1.GetText colStart + 1, fpSpread1.Row, varData2
                If Trim(varData1) = vbNullString And Trim(varData2) = vbNullString Then
                    fpSpread1.Row = fpSpread1.Row
                Else
                    InsertNode colStart, currentRow - 1
                    fpSpread1.Row = currentRow
                End If
            Else
                InsertNode colStart, currentRow - 1
                fpSpread1.Row = currentRow
            End If
        End If
    End If
    ' Ket thuc truong hop them tiep du lieu
' Dat lai vi tri row cho phu luc 01-2 cua to 02 GTGT
If Trim(varMenuId) = "02" Then
    For i = 17 To fpSpread2.MaxRows
        fpSpread2.Col = 2
        fpSpread2.Row = i
        If Left(fpSpread2.Text, 2) = "4." Then
            fpSpread2.Row = i + 1
            Exit For
        End If
    Next
'ElseIf Trim(varMenuId) = "01" And fpSpread1.ActiveSheet = 3 Then
'    For i = 17 To fpSpread2.MaxRows
'        fpSpread2.Col = 2
'        fpSpread2.Row = i
'        If Left(fpSpread2.Text, 2) = "4." Then
'            fpSpread2.Row = i + 1
'            'Exit For
'        End If
'    Next
End If

Do While count < count1 And count2 < fpSpread2.MaxRows
DoEvents
Frame2.Enabled = False
ProgressBar1.value = fpSpread2.Row
'check next row
    fpSpread1.sheet = mCurrentSheet
    fpSpread2.Row = fpSpread2.Row + 1
    value = fpSpread2.value
    If ((Mid(value, 1, 1) = "T" Or Trim(value) = "" Or Trim(value) = vbNullString) And (Trim(varMenuId) = "01" Or Trim(varMenuId) = "02" Or Trim(varMenuId) = "14" Or Trim(varMenuId) = "05" Or Trim(varMenuId) = "59")) Or ((Trim(value) = "" Or Trim(value) = vbNullString) And (Trim(varMenuId) = "17" Or Trim(varMenuId) = "42" Or Trim(varMenuId) = "43" Or Trim(varMenuId) = "44")) Then
        count = count + 1
        inc = True
        ProgressBar1.value = fpSpread2.MaxRows
    ElseIf count = count1 And value = "" Then
        count = count + 1
    Else
        InsertNode colStart, fpSpread1.Row
        inc = False
        count2 = count2 + 1
    End If
        fpSpread2.Row = fpSpread2.Row - 1
    'insert cell
        For Each xmlNode In xmlNodeListMap
            fpSpread2.Col = Conversion.CInt(GetAttribute(xmlNode, "c2"))
            value = fpSpread2.value
           If value <> "" Or value <> vbNullString Then
            fpSpread1.Col = Conversion.CInt(GetAttribute(xmlNode, "c1"))
    'check type of cell
            If Conversion.CInt(GetAttribute(xmlNode, "type")) = 13 Then
                If fpSpread1.CellType = CellTypeNumber Then
                    fpSpread1.TypeNumberNegStyle = TypeNumberNegStyle1
                End If
                fpSpread1.value = value
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 12 Then
                fpSpread1.value = value
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 2 Then
    '            fpSpread2.CellType = CellTypeNumber
    '            fpSpread2.TypeNumberDecPlaces = 0
                fpSpread1.Text = Left(fpSpread2.Text, IIf(InStr(1, fpSpread2.Text, ".") <> 0, InStr(1, fpSpread2.Text, ".") - 1, Len(fpSpread2.Text)))
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 1 Then
              If IsDate(fpSpread2.Text) Then
                Dim arrStr() As String
                Dim sDate As String
                'fpSpread2.CellType = CellTypeDate
                'fpSpread2.TypeDateFormat = TypeDateFormatDDMMYY
                'Dim objCvt As New DateUtils
                'fpSpread2.Text = CStr(objCvt.ToDate(fpSpread2.Text, "DD/MM/YYYY"))
                If InStr(1, fpSpread2.Text, "-") <> 0 Then
                    arrStr = Split(fpSpread2.Text, "-")
                Else
                    arrStr = Split(fpSpread2.Text, "/")
                End If
                
                sDate = Right("00" & arrStr(0), 2) & "/" & Right("00" & arrStr(1), 2) & "/" & Right("20" & arrStr(2), 4)
                
                fpSpread1.Text = sDate
    
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
             End If
            Else
                Select Case strfileFont
                   Case "TCVN"
                      fpSpread1.Text = TAX_Utilities_New.Convert(value, TCVN, UNICODE)
                   Case "VNI"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VNI, UNICODE)
                   Case "VIQR"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VIQR, UNICODE)
                   Case "VISCII"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VISCII, UNICODE)
                   Case Else
                    fpSpread1.Text = value
                End Select
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
            End If
            
          End If
        Next
    'next row
        If inc = True Then
            'have 2 hidden row
'            If themDuLieu Then
'                currentRow = 0
'                If Trim(varMenuId) = "01" Then
'                    incSession = incSession + 1
'                    ' Trong truong hop thuoc to khai 01/GTGT thi session 2 chinh la thue suat  chinh la 0%
'                    If incSession = 1 Then
'                        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(1)
'                        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
'                            currentRow = xmlSecionNode.childNodes.length
'                        End If
'                    End If
'                    ' Trong truong hop thuoc to khai 01/GTGT thi session 3 chinh la thue suat  chinh la 5%
'                    If incSession = 2 Then
'                        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(2)
'                        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
'                            currentRow = xmlSecionNode.childNodes.length
'                        End If
'                    End If
'                    ' Trong truong hop thuoc to khai 01/GTGT thi session 4 chinh la thue suat  chinh la 10%
'                    If incSession = 3 Then
'                        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(3)
'                        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
'                            currentRow = xmlSecionNode.childNodes.length
'                        End If
'                    End If
'                End If
'                fpSpread1.Row = fpSpread1.Row + 5 + currentRow
'                fpSpread2.Row = fpSpread2.Row + 3
'            Else




                Dim temp As Variant
               Dim temp1 As Double
                fpSpread1.Row = fpSpread1.Row + 5
                fpSpread2.Row = fpSpread2.Row + 3
                If count = 3 And Trim(varMenuId) = "01" And fpSpread1.ActiveSheet = 3 Then
                        Do
                            fpSpread2.Col = fpSpread2.ColLetterToNumber("B")
                            temp1 = temp1 + 1
                            temp = fpSpread2.value
                            fpSpread2.Row = fpSpread2.Row + 1
                        Loop Until (Mid(temp, 1, 1) = "T")
                        fpSpread1.Row = fpSpread1.Row + 5
                        fpSpread2.Row = fpSpread2.Row + 1
                        count = count + 1
                End If
'            End If
            'test
              If themDuLieu Then
                Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(count)
                'fpSpread1.Visible = False
                If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
                    currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
                    If (xmlSecionNode.childNodes.length = 1) Then
                        fpSpread1.sheet = mCurrentSheet
                        fpSpread1.GetText colStart, fpSpread1.Row, varData1
                        fpSpread1.GetText colStart + 1, fpSpread1.Row, varData2
                        If Trim(varData1) = vbNullString And Trim(varData2) = vbNullString Then
                            fpSpread1.Row = fpSpread1.Row
                        Else
                            InsertNode colStart, currentRow - 1
                            fpSpread1.Row = currentRow
                        End If
                    Else
                        InsertNode colStart, currentRow - 1
                        fpSpread1.Row = currentRow
                    End If
                End If
            End If
            
            ' end test
        Else
            fpSpread1.Row = fpSpread1.Row + 1
            fpSpread2.Row = fpSpread2.Row + 1
        End If
            fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
            value = fpSpread2.value
    Loop
 ProgressBar1.Visible = False
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True
 If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
 
 Exit Sub
ErrHandle:
 DisplayMessage "0122", msOKOnly, miCriticalError
 ProgressBar1.Visible = False
 ResetData
 ResetDataAndForm mCurrentSheet
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True

End Sub

Private Sub moveData5A()
    Dim value As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim i, count, count1, count2 As Long, countRow As Integer
    
    Dim colStart As Integer
    Dim rowStart As Long
    Dim rowStartSpread2 As Long
    
    Dim varMenuId As String
    
    On Error GoTo ErrHandle
    
    fpSpread1.EventEnabled(EventAllEvents) = False
        ' Truong hop them du lieu va xoa du lieu da ton tai
        If themXoaDuLieu Then
            ResetData
            ResetDataAndForm mCurrentSheet
        End If
        
    ' Lay ID cua Menu
    varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")
    
    fpSpread1.Visible = False
    
    fpSpread2.Visible = True
    ProgressBar1.Visible = True
    ProgressBar1.max = fpSpread2.MaxRows
    ProgressBar1.value = 0
    If Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 2 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_05A_TNCN.xml"))
        colStart = 4
        rowStart = 22
        rowStartSpread2 = 5
    ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 3 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_05B_TNCN.xml"))
        colStart = 3
        rowStart = 22
        rowStartSpread2 = 4
    End If
    
    fpSpread1.Row = rowStart
    
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
   
   fpSpread1.EventEnabled(EventAllEvents) = False
    
    ' Truong hop them tiep du lieu
    Dim xmlSecionNode As MSXML.IXMLDOMNode
    Dim currentRow As Long
    Dim varData1, varData2 As Variant
    Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(0)
    If themDuLieu Then
        'fpSpread1.Visible = False
        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
        End If
    End If
    ' Ket thuc truong hop them tiep du lieu
    Frame2.Enabled = False
    
    fpSpread1.sheet = mCurrentSheet
    
    Dim lRowCtrl, lrowCount, pRow As Long
    Dim varTemp, varTemp1  As Variant
    ' Kiem tra tu dong maxrow len, neu gap bat ky mot dong nao bat dau co du lieu thi se lay do la maxrow luon
    For lrowCount = fpSpread2.MaxRows To 0 Step -1
       fpSpread2.GetText fpSpread2.ColLetterToNumber("B"), lrowCount, varTemp
       fpSpread2.GetText fpSpread2.ColLetterToNumber("F"), lrowCount, varTemp1
       If (Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "") And (Trim(varTemp1) <> vbNullString Or Trim(varTemp1) <> "") Then
            ' Tru tiep 4 dong header dau tien thi se duoc tong so dong can import vao
            If mCurrentSheet = 2 Then
                lrowCount = lrowCount - 4
            ElseIf mCurrentSheet = 3 Then
                lrowCount = lrowCount - 3
            End If
            Exit For
       End If
    Next
    ' Ca hai bang ke thi dong du lieu bat dau = dong du lieu bat dau - 1
    rowStartSpread2 = rowStartSpread2 - 1
    ' Ca hai bang ke trong to quyet toan 5A bat dau tu dong 22, 5B bat dau tu dong 21
    If themDuLieu Then
        rowStart = currentRow - 3
    Else
        rowStart = rowStart - 2
    End If
        
    With fpSpread1
        
        Dim blockRow, stepRow  As Integer
        
        
        blockRow = 50
        stepRow = 1
        
        .MaxRows = .MaxRows + lrowCount
        
'        Debug.Print "start: " & Time
        Lbload.Visible = True
        If themDuLieu And xmlSecionNode.childNodes.length > 1 Then
            lrowCount = lrowCount + 1
        End If
        For lRowCtrl = 1 To lrowCount - 1

            DoEvents

            ProgressBar1.value = lRowCtrl

            pRow = lRowCtrl + 1

            If (pRow <= blockRow + 1) Or ((blockRow * stepRow) + 1 > lrowCount) Then
                'dhdang sua ngay 27/05
                If mCurrentSheet = 2 Then
                    'dntai 03/02/2012 set thanh 2 mang de xu ly cho cot 16 mat formula tu tinh
                    ReDim fparray(0, 8) As Variant
                    ReDim fparray1(0, 1) As Variant
                ElseIf mCurrentSheet = 3 Then
                    ReDim fparray(0, 7) As Variant
                End If
                'lay gia tri, bo qua cot 16 de set lai gia tri
                fpSpread2.GetArray 2, lRowCtrl + rowStartSpread2, fparray
                ' set BK 05A/TNCN
                If mCurrentSheet = 2 Then
                    fpSpread2.GetArray 12, lRowCtrl + rowStartSpread2, fparray1
                    'lam tron so cac cell tren 05A_TNCN
                    For countRow = 0 To UBound(fparray)
                        If fparray(countRow, 4) <> vbNullString Then
                            fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                        End If
                        If fparray(countRow, 5) <> vbNullString Then
                            fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                        End If
                        If fparray(countRow, 7) <> vbNullString Then
                            fparray(countRow, 7) = Round(fparray(countRow, 7), 0)
                        End If
                        If fparray(countRow, 8) <> vbNullString Then
                            fparray(countRow, 8) = Round(fparray(countRow, 8), 0)
                        End If
                        If fparray1(countRow, 0) <> vbNullString Then
                            fparray1(countRow, 0) = Round(fparray1(countRow, 0), 0)
                        End If
                    Next
                    countRow = 0
                ElseIf mCurrentSheet = 3 Then
                    'lam tron cac cell tren 05_TNCN
                    For countRow = 0 To UBound(fparray)
                        If fparray(countRow, 4) <> vbNullString Then
                            fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                        End If
                        If fparray(countRow, 5) <> vbNullString Then
                            fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                        End If
                        If fparray(countRow, 6) <> vbNullString Then
                            fparray(countRow, 6) = Round(fparray(countRow, 6), 0)
                        End If
                    Next
                    countRow = 0
                End If
                    
                .InsertRows (rowStart + 1 + pRow), 1
                .CopyRowRange rowStart + pRow, rowStart + pRow, (rowStart + 1) + pRow
                
                If xmlSecionNode.childNodes.length > 1 Then
                    .SetArray colStart, rowStart + pRow + 1, fparray
                    ' Chi bang ke 05A/TNCN moi set lai
                    If mCurrentSheet = 2 Then
                    .SetArray .ColLetterToNumber("N"), rowStart + pRow + 1, fparray1
                    End If
                Else
                    .SetArray colStart, rowStart + pRow, fparray
                    ' Chi bang ke 05A/TNCN moi set lai
                    If mCurrentSheet = 2 Then
                    .SetArray .ColLetterToNumber("N"), rowStart + pRow, fparray1
                    End If
                End If

                If pRow = blockRow + 1 Then stepRow = 2
                If (lRowCtrl = lrowCount - 1) And (stepRow > 1) Then
                    If xmlSecionNode.childNodes.length = 1 Then
                        fpSpread2.GetArray 2, lrowCount + rowStartSpread2, fparray
                        
                        'lam tron so tren cac cell truoc khi set value
                        If mCurrentSheet = 2 Then
                            'lam tron so cac cell tren 05A_TNCN
                            For countRow = 0 To UBound(fparray)
                                If fparray(countRow, 4) <> vbNullString Then
                                    fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                                End If
                                If fparray(countRow, 5) <> vbNullString Then
                                    fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                                End If
                                If fparray(countRow, 7) <> vbNullString Then
                                    fparray(countRow, 7) = Round(fparray(countRow, 7), 0)
                                End If
                                If fparray(countRow, 8) <> vbNullString Then
                                    fparray(countRow, 8) = Round(fparray(countRow, 8), 0)
                                End If
                            Next
                            countRow = 0
                        ElseIf mCurrentSheet = 3 Then
                            'lam tron cac cell tren 05_TNCN
                            For countRow = 0 To UBound(fparray)
                                If fparray(countRow, 4) <> vbNullString Then
                                    fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                                End If
                                If fparray(countRow, 5) <> vbNullString Then
                                    fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                                End If
                                If fparray(countRow, 6) <> vbNullString Then
                                    fparray(countRow, 6) = Round(fparray(countRow, 6), 0)
                                End If
                            Next
                            countRow = 0
                        End If
                        'end
                        
                        .SetArray colStart, rowStart + pRow + 1, fparray
                        ' Chi bang ke 05A/TNCN moi set lai
                        If mCurrentSheet = 2 Then
                        fpSpread2.GetArray 12, lrowCount + rowStartSpread2, fparray1
                        
                        'lam trong so truoc khi set
                            For countRow = 0 To UBound(fparray1)
                                If fparray1(countRow, 0) <> vbNullString Then
                                    fparray1(countRow, 0) = Round(fparray1(countRow, 0), 0)
                                End If
                            Next
                            countRow = 0
                        'end
                        .SetArray .ColLetterToNumber("N"), rowStart + pRow + 1, fparray1
                        End If
                    Else
                        fpSpread2.GetArray 2, lrowCount + rowStartSpread2 - 1, fparray
                        
                        'lam tron so tren cac cell truoc khi set value
                        If mCurrentSheet = 2 Then
                            'lam tron so cac cell tren 05A_TNCN
                            For countRow = 0 To UBound(fparray)
                                If fparray(countRow, 4) <> vbNullString Then
                                    fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                                End If
                                If fparray(countRow, 5) <> vbNullString Then
                                    fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                                End If
                                If fparray(countRow, 7) <> vbNullString Then
                                    fparray(countRow, 7) = Round(fparray(countRow, 7), 0)
                                End If
                                If fparray(countRow, 8) <> vbNullString Then
                                    fparray(countRow, 8) = Round(fparray(countRow, 8), 0)
                                End If
                            Next
                            countRow = 0
                        ElseIf mCurrentSheet = 3 Then
                            'lam tron cac cell tren 05_TNCN
                            For countRow = 0 To UBound(fparray)
                                If fparray(countRow, 4) <> vbNullString Then
                                     fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                                 End If
                                 If fparray(countRow, 5) <> vbNullString Then
                                     fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                                 End If
                                 If fparray(countRow, 6) <> vbNullString Then
                                     fparray(countRow, 6) = Round(fparray(countRow, 6), 0)
                                 End If
                            Next
                            countRow = 0
                        End If
                        'end
                        
                        
                        .SetArray colStart, rowStart + pRow + 1, fparray
                        ' Chi bang ke 05A/TNCN moi set lai
                        If mCurrentSheet = 2 Then
                        fpSpread2.GetArray 12, lrowCount + rowStartSpread2 - 1, fparray1
                        
                        
                        'lam trong so truoc khi set
                            For countRow = 0 To UBound(fparray1)
                                If fparray1(countRow, 0) <> vbNullString Then
                                    fparray1(countRow, 0) = Round(fparray1(countRow, 0), 0)
                                End If
                            Next
                            countRow = 0
                        'end
                        
                        
                        .SetArray .ColLetterToNumber("N"), rowStart + pRow + 1, fparray1
                        End If
                    End If
                ElseIf (lRowCtrl = lrowCount - 1) And xmlSecionNode.childNodes.length = 1 Then
                    fpSpread2.GetArray 2, lrowCount + rowStartSpread2, fparray
                    
                    'lam tron so tren cac cell truoc khi set value
                    If mCurrentSheet = 2 Then
                        'lam tron so cac cell tren 05A_TNCN
                        For countRow = 0 To UBound(fparray)
                            If fparray(countRow, 4) <> vbNullString Then
                                fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                            End If
                            If fparray(countRow, 5) <> vbNullString Then
                                fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                            End If
                            If fparray(countRow, 7) <> vbNullString Then
                                fparray(countRow, 7) = Round(fparray(countRow, 7), 0)
                            End If
                            If fparray(countRow, 8) <> vbNullString Then
                                fparray(countRow, 8) = Round(fparray(countRow, 8), 0)
                            End If
                        Next
                        countRow = 0
                    ElseIf mCurrentSheet = 3 Then
                        'lam tron cac cell tren 05_TNCN
                        For countRow = 0 To UBound(fparray)
                            If fparray(countRow, 4) <> vbNullString Then
                                fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                            End If
                            If fparray(countRow, 5) <> vbNullString Then
                                fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                            End If
                            If fparray(countRow, 6) <> vbNullString Then
                                fparray(countRow, 6) = Round(fparray(countRow, 6), 0)
                            End If
                        Next
                        countRow = 0
                    End If
                    'end
                    
                    
                    .SetArray colStart, rowStart + pRow + 1, fparray
                    ' Chi bang ke 05A/TNCN moi set lai
                    If mCurrentSheet = 2 Then
                    fpSpread2.GetArray 12, lrowCount + rowStartSpread2, fparray1
                    
                    'lam trong so truoc khi set
                        For countRow = 0 To UBound(fparray1)
                                If fparray1(countRow, 0) <> vbNullString Then
                                    fparray1(countRow, 0) = Round(fparray1(countRow, 0), 0)
                                End If
                        Next
                        countRow = 0
                    'end
                    
                    .SetArray .ColLetterToNumber("N"), rowStart + pRow + 1, fparray1
                    End If
                End If
            ElseIf pRow = (blockRow * stepRow) + 1 Then
                'dhdang sua ngay 27/05
                If mCurrentSheet = 2 Then
                    ReDim fparray(50, 8) As Variant
                    ReDim fparray1(50, 1) As Variant
                ElseIf mCurrentSheet = 3 Then
                    ReDim fparray(50, 7) As Variant
                End If
                fpSpread2.GetArray 2, (blockRow * (stepRow - 1) + rowStartSpread2 + 1), fparray
                
                'lam tron so tren cac cell truoc khi set value
                If mCurrentSheet = 2 Then
                    'lam tron so cac cell tren 05A_TNCN
                    For countRow = 0 To UBound(fparray)
                        If fparray(countRow, 4) <> vbNullString Then
                            fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                        End If
                        If fparray(countRow, 5) <> vbNullString Then
                            fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                        End If
                        If fparray(countRow, 7) <> vbNullString Then
                            fparray(countRow, 7) = Round(fparray(countRow, 7), 0)
                        End If
                        If fparray(countRow, 8) <> vbNullString Then
                            fparray(countRow, 8) = Round(fparray(countRow, 8), 0)
                        End If
                    Next
                    countRow = 0
                ElseIf mCurrentSheet = 3 Then
                    'lam tron cac cell tren 05_TNCN
                    For countRow = 0 To UBound(fparray)
                        If fparray(countRow, 4) <> vbNullString Then
                            fparray(countRow, 4) = Round(fparray(countRow, 4), 0)
                        End If
                        If fparray(countRow, 5) <> vbNullString Then
                            fparray(countRow, 5) = Round(fparray(countRow, 5), 0)
                        End If
                        If fparray(countRow, 6) <> vbNullString Then
                            fparray(countRow, 6) = Round(fparray(countRow, 6), 0)
                        End If
                    Next
                    countRow = 0
                End If
                'end
                
                ' set cho BK 05A/TNCN
                If mCurrentSheet = 2 Then
                fpSpread2.GetArray 12, (blockRow * (stepRow - 1) + rowStartSpread2 + 1), fparray1
                
                'lam trong so truoc khi set
                    For countRow = 0 To UBound(fparray1)
                        If fparray1(countRow, 0) <> vbNullString Then
                            fparray1(countRow, 0) = Round(fparray1(countRow, 0), 0)
                        End If
                    Next
                    countRow = 0
                'end
                
                End If
                
                stepRow = stepRow + 1
                
                .InsertRows (rowStart + 2) + (blockRow * (stepRow - 2)) + 1, blockRow
                
                '.CopyRowRange (rowStart + 2), (rowStart + 2) + blockRow - 1, (rowStart + 2) + (blockRow * (stepRow - 2)) + 1
                .CopyRowRange 22, 71, (rowStart + 2) + (blockRow * (stepRow - 2)) + 1
                
                .SetArray colStart, (rowStart + 2) + (blockRow * (stepRow - 2)), fparray
                ' set BK 05A/TNCN
                If mCurrentSheet = 2 Then
                .SetArray .ColLetterToNumber("N"), (rowStart + 2) + (blockRow * (stepRow - 2)), fparray1
                End If
            End If

        Next
        'dhdang and nvhai edit convert font to unicode
        'Debug.Print "begin" & Time
        If strfileFont <> "UNICODE" Then
            For lRowCtrl = 1 To lrowCount
                fpSpread1.Col = fpSpread1.ColLetterToNumber("D")
                fpSpread1.Row = lRowCtrl + rowStart + 1
                Select Case strfileFont
                   Case "TCVN"
                      fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, TCVN, UNICODE)
                   Case "VNI"
                    fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VNI, UNICODE)
                   Case "VIQR"
                    fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VIQR, UNICODE)
                   Case "VISCII"
                    fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VISCII, UNICODE)
                   Case Else
                    fpSpread1.Text = fpSpread1.Text
                End Select
                UpdateCell fpSpread1.ColLetterToNumber("D"), fpSpread1.Row, fpSpread1.Text
            Next
        End If
         'Debug.Print "end" & Time
        .Col = colStart
        .Row = rowStart + 2
        UpdateCell colStart, rowStart + 2, .Text
        .Col = colStart + 1
        UpdateCell colStart + 1, rowStart + 2, .Text
        .Col = colStart + 2
        UpdateCell colStart + 2, rowStart + 2, .Text
        .Col = colStart + 3
        UpdateCell colStart + 3, rowStart + 2, .Text
        
'        Debug.Print "end: " & Time
        Lbload.Visible = False
        .ReDraw = True
        
    End With
    
    fpSpread1.Visible = True
    
    ProgressBar1.Visible = False
    Frame2.Enabled = True
    fpSpread1.EventEnabled(EventAllEvents) = True
    If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
    'If Not objTaxBusiness Is Nothing Then objTaxBusiness.finish
'    If mCurrentSheet = 2 Then
'        If Not objTaxBusiness Is Nothing Then objTaxBusiness.InsertNode 22, lRowCount, mCurrentSheet
'    ElseIf mCurrentSheet = 3 Then
'        If Not objTaxBusiness Is Nothing Then objTaxBusiness.InsertNode 21, lRowCount, mCurrentSheet
'    End If
    Exit Sub
ErrHandle:
    DisplayMessage "0122", msOKOnly, miCriticalError
    ProgressBar1.Visible = False
    ResetData
    ResetDataAndForm mCurrentSheet
    Frame2.Enabled = True
    fpSpread1.EventEnabled(EventAllEvents) = True

End Sub

Private Sub moveDataToKhai()
Dim value As String
Dim xmlDocument As New MSXML.DOMDocument
Dim xmlNode As MSXML.IXMLDOMNode
Dim i, count, count1, count2 As Long
Dim inc As Boolean
Dim colStart As Integer
Dim varMenuId As String

On Error GoTo ErrHandle

'Delete data exit
     fpSpread1.EventEnabled(EventAllEvents) = False
     ResetData
     ResetDataAndForm mCurrentSheet

' Lay ID cua Menu
varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")

fpSpread2.Visible = False
ProgressBar1.Visible = True
fpSpread2.sheet = mCurrentSheet
ProgressBar1.max = fpSpread2.MaxRows
ProgressBar1.value = 0
If Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_05A_TNCN.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 3 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_05B_TNCN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 4 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_PL_01_TNCN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "42" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_02A_TNCN_BH.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "43" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_02A_TNCN_XS.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "44" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_06D_TNCN.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "41" And fpSpread1.ActiveSheet = 4 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_PL_09C_TNCN.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "59" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_PL_06_1_TNCN.xml"))
    colStart = 4
ElseIf Trim(varMenuId) = "76" And fpSpread1.ActiveSheet = 1 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_08B_TNCN.xml"))
    colStart = 3
End If


Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
   fpSpread1.EventEnabled(EventAllEvents) = False
   fpSpread1.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row1").Item(0).Text)
   fpSpread2.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row2").Item(0).Text)
   fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
   count1 = Conversion.CInt(xmlDocument.getElementsByTagName("count").Item(0).Text)
   
Do While count < count1 And count2 < fpSpread2.MaxRows
DoEvents
Frame2.Enabled = False
ProgressBar1.value = fpSpread2.Row
'check next row
    fpSpread1.sheet = mCurrentSheet
    fpSpread2.sheet = mCurrentSheet
    fpSpread2.Row = fpSpread2.Row + 1
    value = fpSpread2.value
    
    If Trim(value) = "" Or Trim(value) = vbNullString Or Trim(value) = "aa" Then
        count = count + 1
        inc = True
        ProgressBar1.value = fpSpread2.MaxRows
    ElseIf count = count1 And value = "" Then
        count = count + 1
    Else
        InsertNode colStart, fpSpread1.Row
        inc = False
        count2 = count2 + 1
    End If
        fpSpread2.Row = fpSpread2.Row - 1
'insert cell
    For Each xmlNode In xmlNodeListMap
        fpSpread2.Col = Conversion.CInt(GetAttribute(xmlNode, "c2"))
        value = fpSpread2.value
       If value <> "" Or value <> vbNullString Then
        fpSpread1.Col = Conversion.CInt(GetAttribute(xmlNode, "c1"))
'check type of cell
        If Conversion.CInt(GetAttribute(xmlNode, "type")) = 13 Then
            If fpSpread1.CellType = CellTypeNumber Then
                fpSpread1.TypeNumberNegStyle = TypeNumberNegStyle1
            End If
            fpSpread1.value = Round(value, 0)
            UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
        ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 12 Then
            fpSpread1.value = value
            UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
        ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 2 Then
'            fpSpread2.CellType = CellTypeNumber
'            fpSpread2.TypeNumberDecPlaces = 0
            fpSpread1.Text = Left(fpSpread2.Text, IIf(InStr(1, fpSpread2.Text, ".") <> 0, InStr(1, fpSpread2.Text, ".") - 1, Len(fpSpread2.Text)))
            UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
        ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 1 Then
          If IsDate(fpSpread2.Text) Then
            Dim arrStr() As String
            Dim sDate As String
            'fpSpread2.CellType = CellTypeDate
            'fpSpread2.TypeDateFormat = TypeDateFormatDDMMYY
            'Dim objCvt As New DateUtils
            'fpSpread2.Text = CStr(objCvt.ToDate(fpSpread2.Text, "DD/MM/YYYY"))
            If InStr(1, fpSpread2.Text, "-") <> 0 Then
                arrStr = Split(fpSpread2.Text, "-")
            Else
                arrStr = Split(fpSpread2.Text, "/")
            End If
            
            sDate = Right("00" & arrStr(0), 2) & "/" & Right("00" & arrStr(1), 2) & "/" & Right("20" & arrStr(2), 4)
            
            fpSpread1.Text = sDate

            UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
         End If
        Else
            Select Case strfileFont
               Case "TCVN"
                  fpSpread1.Text = TAX_Utilities_New.Convert(value, TCVN, UNICODE)
               Case "VNI"
                fpSpread1.Text = TAX_Utilities_New.Convert(value, VNI, UNICODE)
               Case "VIQR"
                fpSpread1.Text = TAX_Utilities_New.Convert(value, VIQR, UNICODE)
               Case "VISCII"
                fpSpread1.Text = TAX_Utilities_New.Convert(value, VISCII, UNICODE)
               Case Else
                fpSpread1.Text = value
            End Select
            UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
        End If
        
      End If
    Next
    'next row
        If inc = True Then
            'have 2 hidden row
            fpSpread1.Row = fpSpread1.Row + 5
            fpSpread2.Row = fpSpread2.Row + 3
        Else
            fpSpread1.Row = fpSpread1.Row + 1
            fpSpread2.Row = fpSpread2.Row + 1
        End If
            fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
            value = fpSpread2.value
    Loop
 ProgressBar1.Visible = False
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True
 If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
 
 Exit Sub
ErrHandle:
 DisplayMessage "0122", msOKOnly, miCriticalError
 ProgressBar1.Visible = False
 ResetData
 ResetDataAndForm mCurrentSheet
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True

End Sub

Private Sub moveDataToKhai5A()
    Dim value As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim i, count, count1, count2 As Long
    
    Dim colStart As Integer
    Dim rowStart As Long
    Dim rowStartSpread2 As Long
    
    Dim varMenuId As String
    
    On Error GoTo ErrHandle
    
    fpSpread1.EventEnabled(EventAllEvents) = False
        
    ResetData
    ResetDataAndForm mCurrentSheet
     
    ' Lay ID cua Menu
    varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")
    Frame3.Visible = True
    fpSpread2.Visible = True
    ProgressBar1.Visible = True
    fpSpread2.sheet = mCurrentSheet
    'ProgressBar1.max = fpSpread2.MaxRows
    ProgressBar1.value = 0
    If Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 2 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_05A_TNCN.xml"))
        colStart = 4
        rowStart = 22
        rowStartSpread2 = 22
    ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 3 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_05B_TNCN.xml"))
        colStart = 3
        rowStart = 22
        rowStartSpread2 = 22
    End If
    
    fpSpread1.Row = rowStart
    
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
   
   fpSpread1.EventEnabled(EventAllEvents) = False
    
    ' Truong hop them tiep du lieu
    Dim xmlSecionNode As MSXML.IXMLDOMNode
    Dim currentRow As Long
    Dim varData1, varData2 As Variant
    Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(0)
    If themDuLieu Then
        'fpSpread1.Visible = False
        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
        End If
    End If
    ' Ket thuc truong hop them tiep du lieu
    Frame2.Enabled = False
    
    fpSpread1.sheet = mCurrentSheet
    fpSpread2.sheet = mCurrentSheet
    
    Dim lRowCtrl, lrowCount, pRow As Long
    Dim varTemp, varTemp1  As Variant
    ' Kiem tra tu dong maxrow len, neu gap bat ky mot dong nao bat dau co du lieu thi se lay do la maxrow luon
    For lrowCount = fpSpread2.MaxRows To 0 Step -1
       If mCurrentSheet = 2 Then
            fpSpread2.GetText fpSpread2.ColLetterToNumber("D"), lrowCount, varTemp
            fpSpread2.GetText fpSpread2.ColLetterToNumber("H"), lrowCount, varTemp1
       ElseIf mCurrentSheet = 3 Then
            fpSpread2.GetText fpSpread2.ColLetterToNumber("C"), lrowCount, varTemp
            fpSpread2.GetText fpSpread2.ColLetterToNumber("G"), lrowCount, varTemp1
       End If
       If (Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "") And (Trim(varTemp1) <> vbNullString Or Trim(varTemp1) <> "") Then
            ' Tru tiep 4 dong header dau tien thi se duoc tong so dong can import vao
            If mCurrentSheet = 2 Then
                lrowCount = lrowCount - 21
                If lrowCount >= 0 Then
                    ProgressBar1.max = lrowCount
                End If
            ElseIf mCurrentSheet = 3 Then
                lrowCount = lrowCount - 21
                If lrowCount >= 0 Then
                    ProgressBar1.max = lrowCount
                End If
                    'ProgressBar1.max = lrowCount
            End If
            Exit For
       End If
    Next
    ' Ca hai bang ke thi dong du lieu bat dau = dong du lieu bat dau - 1
    rowStartSpread2 = rowStartSpread2 - 1
    ' Ca hai bang ke trong to quyet toan 5A bat dau tu dong 22, 5B bat dau tu dong 21
    If themDuLieu Then
        rowStart = currentRow - 3
    Else
        rowStart = rowStart - 2
    End If
        
    With fpSpread1
        
        Dim blockRow, stepRow  As Integer
        
        
        blockRow = 50
        stepRow = 1
        
        .MaxRows = .MaxRows + lrowCount
        
'        Debug.Print "start: " & Time
        Lbload.Visible = True
        If themDuLieu And xmlSecionNode.childNodes.length > 1 Then
            lrowCount = lrowCount + 1
        End If
        For lRowCtrl = 1 To lrowCount - 1

            DoEvents

            ProgressBar1.value = lRowCtrl

            pRow = lRowCtrl + 1

            If (pRow <= blockRow + 1) Or ((blockRow * stepRow) + 1 > lrowCount) Then
                'dhdang sua ngay 27/05
                If mCurrentSheet = 2 Then
                    ReDim fparray(0, 10) As Variant
                ElseIf mCurrentSheet = 3 Then
                    ReDim fparray(0, 5) As Variant
                End If
                
                fpSpread2.GetArray colStart, lRowCtrl + rowStartSpread2, fparray
                
                .InsertRows (rowStart + 1 + pRow), 1
                .CopyRowRange rowStart + pRow, rowStart + pRow, (rowStart + 1) + pRow
                
                If xmlSecionNode.childNodes.length > 1 Then
                    .SetArray colStart, rowStart + pRow + 1, fparray
                Else
                    .SetArray colStart, rowStart + pRow, fparray
                End If

                If pRow = blockRow + 1 Then stepRow = 2
                If (lRowCtrl = lrowCount - 1) And (stepRow > 1) Then
                    If xmlSecionNode.childNodes.length = 1 Then
                        fpSpread2.GetArray colStart, lrowCount + rowStartSpread2, fparray
                        .SetArray colStart, rowStart + pRow + 1, fparray
                    Else
                        fpSpread2.GetArray colStart, lrowCount + rowStartSpread2 - 1, fparray
                        .SetArray colStart, rowStart + pRow + 1, fparray
                    End If
                ElseIf (lRowCtrl = lrowCount - 1) And xmlSecionNode.childNodes.length = 1 Then
                    fpSpread2.GetArray colStart, lrowCount + rowStartSpread2, fparray
                    .SetArray colStart, rowStart + pRow + 1, fparray
                End If
            ElseIf pRow = (blockRow * stepRow) + 1 Then
                'dhdang sua ngay 27/05
                If mCurrentSheet = 2 Then
                    ReDim fparray(50, 10) As Variant
                ElseIf mCurrentSheet = 3 Then
                    ReDim fparray(50, 5) As Variant
                End If
                fpSpread2.GetArray colStart, (blockRow * (stepRow - 1) + rowStartSpread2 + 1), fparray
                
                stepRow = stepRow + 1
                
                .InsertRows (rowStart + 2) + (blockRow * (stepRow - 2)) + 1, blockRow
                
                '.CopyRowRange (rowStart + 2), (rowStart + 2) + blockRow - 1, (rowStart + 2) + (blockRow * (stepRow - 2)) + 1
                .CopyRowRange 22, 71, (rowStart + 2) + (blockRow * (stepRow - 2)) + 1
                
                .SetArray colStart, (rowStart + 2) + (blockRow * (stepRow - 2)), fparray

            End If

        Next
        'dhdang and nvhai edit convert font to unicode
        'Debug.Print "begin" & Time
        If strfileFont <> "UNICODE" Then
            For lRowCtrl = 1 To lrowCount
                fpSpread1.Col = fpSpread1.ColLetterToNumber("D")
                fpSpread1.Row = lRowCtrl + rowStart + 1
                Select Case strfileFont
                   Case "TCVN"
                      fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, TCVN, UNICODE)
                   Case "VNI"
                    fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VNI, UNICODE)
                   Case "VIQR"
                    fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VIQR, UNICODE)
                   Case "VISCII"
                    fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VISCII, UNICODE)
                   Case Else
                    fpSpread1.Text = fpSpread1.Text
                End Select
                UpdateCell fpSpread1.ColLetterToNumber("D"), fpSpread1.Row, fpSpread1.Text
            Next
        End If
         'Debug.Print "end" & Time
        .Col = colStart
        .Row = rowStart + 2
        UpdateCell colStart, rowStart + 2, .Text
        .Col = colStart + 1
        UpdateCell colStart + 1, rowStart + 2, .Text
        .Col = colStart + 2
        UpdateCell colStart + 2, rowStart + 2, .Text
        .Col = colStart + 3
        UpdateCell colStart + 3, rowStart + 2, .Text
        
'        Debug.Print "end: " & Time
        Lbload.Visible = False
        .ReDraw = True
        
    End With
    
    ProgressBar1.Visible = False
    Frame2.Enabled = True
    fpSpread1.EventEnabled(EventAllEvents) = True
    If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
   
'    If mCurrentSheet = 2 Then
''        If Not objTaxBusiness Is Nothing Then objTaxBusiness.InsertNode 23, 23, mCurrentSheet
'        If Not objTaxBusiness Is Nothing Then objTaxBusiness.fnis
'    ElseIf mCurrentSheet = 3 Then
''        If Not objTaxBusiness Is Nothing Then objTaxBusiness.InsertNode 22, 22, mCurrentSheet
'        'delNullRow 3
'    End If
    
    Exit Sub
   
ErrHandle:
    DisplayMessage "0122", msOKOnly, miCriticalError
    ProgressBar1.Visible = False
    ResetData
    ResetDataAndForm mCurrentSheet
    Frame2.Enabled = True
    fpSpread1.EventEnabled(EventAllEvents) = True

End Sub


Private Sub Cmd_Seach_Click()
Dim seach  As Variant
Dim Option_Seach  As String
Dim ret2 As Long
Dim Y As Long
    Static so, curenrow, curencol As Long
    If so = 0 Then
        so = 1
    End If
    txt_Seach.sheet = 1
    txt_Seach.Col = 1
    txt_Seach.Row = 1
    seach = txt_Seach.Text
    
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            
               If .sheet = 1 Then
                    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "45" Then
                        If Trim(seach) = "" Then
                               DisplayMessage "0160", msOKOnly, miInformation
                               txt_Seach.SetFocus
                        Else
                            'If Col = .ColLetterToNumber("E") And Row = 11 Then
                            'T�m kiem c�c ban ghi giong nhau
                             If curenrow <= 21 Or curenrow > 21 + .MaxRows Then
                               curenrow = 21
                             End If
                             
                             'seach = Trim(txt_seach.Text)
                             Option_Seach = Cb_seach.ListIndex
                             If Option_Seach = 0 Then
                                    '.Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                     ret2 = .SearchCol(.ColLetterToNumber("D"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                     curenrow = ret2
                                     curencol = .ColLetterToNumber("D")
                             ElseIf Option_Seach = 1 Then
                                    '.Sort .ColLetterToNumber("E"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                    ret2 = .SearchCol(.ColLetterToNumber("E"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                    curenrow = ret2
                                    curencol = .ColLetterToNumber("E")
                             ElseIf Option_Seach = 2 Then
                                    '.Sort .ColLetterToNumber("F"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                    ret2 = .SearchCol(.ColLetterToNumber("F"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                    curenrow = ret2
                                    curencol = .ColLetterToNumber("F")
                             End If
                             'Select cell
                             If ret2 > -1 Then
                                 
                                .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("X"), ret2
                                '.BackColor = vbGreen
                                .Refresh
                             Else
                                ret2 = .SearchCol(curencol, 21, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                   If ret2 > -1 Then
                                     .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("X"), ret2
                                     curenrow = ret2
                                     '.BackColor = vbGreen
                                     .Refresh
                                   Else
                                         'MsgBox "Khong co ban ghi nay.", vbInformation
                                       DisplayMessage "0160", msOKOnly, miInformation
                                   End If
                             End If
                    End If
                End If
            End If
            
            ' Note: Truong hop Sheet2 va Sheet3 trong cung mot dieu kien re nhanh, Sheet1 la doc lap
            ' Nho sua lai sau
            
            If .sheet = 2 Then
               ' to khai 05KK-TNCN
               If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Then
                If Trim(seach) = "" Then
                       DisplayMessage "0160", msOKOnly, miInformation
                       'txt_Seach.Text = ""
                       txt_Seach.SetFocus
                       Else
                             'If Col = .ColLetterToNumber("E") And Row = 11 Then
                                'T�m kiem c�c ban ghi giong nhau
                                 If curenrow <= 21 Or curenrow > 21 + .MaxRows Then
                                   curenrow = 21
                                 End If
                                 
                                 Option_Seach = Cb_seach.ListIndex
                                 If Option_Seach = 0 Then
                                    '.Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                     ret2 = .SearchCol(.ColLetterToNumber("D"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                     curenrow = ret2
                                     curencol = .ColLetterToNumber("D")
                                 Else
                                   If Option_Seach = 1 Then
                                         '.Sort .ColLetterToNumber("E"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                         ret2 = .SearchCol(.ColLetterToNumber("E"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                         curenrow = ret2
                                         curencol = .ColLetterToNumber("E")
                                         Else
                                             If Option_Seach = 2 Then
                                                 '.Sort .ColLetterToNumber("F"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                                 ret2 = .SearchCol(.ColLetterToNumber("F"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                                 curenrow = ret2
                                                 curencol = .ColLetterToNumber("F")
                                             End If
                                   End If
                                  End If
                                 'Select cell
                                 If ret2 > -1 Then
                                     
                                    .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("AG"), ret2
                                    '.BackColor = vbGreen
                                    .Refresh
                                 Else
                                   
                                    ret2 = .SearchCol(curencol, 21, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                       If ret2 > -1 Then
                                         .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("AG"), ret2
                                         curenrow = ret2
                                         '.BackColor = vbGreen
                                         .Refresh
                                       Else
                                           txt_Seach.SetFocus
                                             'MsgBox "Khong co ban ghi nay.", vbInformation
                                           DisplayMessage "0160", msOKOnly, miInformation
                                       End If
                                 End If
                       End If
            Else
                If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "42" Then
                    If Trim(seach) = "" Then
                       DisplayMessage "0160", msOKOnly, miInformation
                       'txt_Seach.Text = ""
                       txt_Seach.SetFocus
                       Else
                             'If Col = .ColLetterToNumber("E") And Row = 11 Then
                                'T�m kiem c�c ban ghi giong nhau
                                 If curenrow <= 21 Or curenrow > 21 + .MaxRows Then
                                   curenrow = 21
                                 End If
                                 
                                 'seach = Trim(txt_seach.Text)
                                 Option_Seach = Cb_seach.ListIndex
                                 If Option_Seach = 0 Then
                                    '.Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                     ret2 = .SearchCol(.ColLetterToNumber("D"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                     curenrow = ret2
                                     curencol = .ColLetterToNumber("D")
                                 Else
                                   If Option_Seach = 1 Then
                                         '.Sort .ColLetterToNumber("E"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                         ret2 = .SearchCol(.ColLetterToNumber("E"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                         curenrow = ret2
                                         curencol = .ColLetterToNumber("E")
                                         Else
                                             If Option_Seach = 2 Then
                                                 '.Sort .ColLetterToNumber("F"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                                 ret2 = -1
                                                 curenrow = ret2
                                                 curencol = .ColLetterToNumber("F")
                                             End If
                                   End If
                                  End If
                                 'Select cell
                                 If ret2 > -1 Then
                                     
                                    .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("J"), ret2
                                    '.BackColor = vbGreen
                                    .Refresh
                                 Else
                                   
                                    ret2 = .SearchCol(curencol, 21, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                       If ret2 > -1 Then
                                         .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("J"), ret2
                                         curenrow = ret2
                                         '.BackColor = vbGreen
                                         .Refresh
                                       Else
                                             'MsgBox "Khong co ban ghi nay.", vbInformation
                                           DisplayMessage "0160", msOKOnly, miInformation
                                       End If
                                 End If
                       End If
                       
                 Else
                 If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "43" Then
                    If Trim(seach) = "" Then
                       DisplayMessage "0160", msOKOnly, miInformation
                       
                       txt_Seach.SetFocus
                       Else
                             'If Col = .ColLetterToNumber("E") And Row = 11 Then
                                'T�m kiem c�c ban ghi giong nhau
                                 If curenrow <= 21 Or curenrow > 21 + .MaxRows Then
                                   curenrow = 21
                                 End If
                                 
                                 'seach = Trim(txt_seach.Text)
                                 Option_Seach = Cb_seach.ListIndex
                                 If Option_Seach = 0 Then
                                    '.Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                     ret2 = .SearchCol(.ColLetterToNumber("D"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                     curenrow = ret2
                                     curencol = .ColLetterToNumber("D")
                                 Else
                                   If Option_Seach = 1 Then
                                         '.Sort .ColLetterToNumber("E"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                         ret2 = .SearchCol(.ColLetterToNumber("E"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                         curenrow = ret2
                                         curencol = .ColLetterToNumber("E")
                                         Else
                                             If Option_Seach = 2 Then
                                                 '.Sort .ColLetterToNumber("F"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                                 ret2 = .SearchCol(.ColLetterToNumber("F"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                                 curenrow = ret2
                                                 curencol = .ColLetterToNumber("F")
                                             End If
                                   End If
                                  End If
                                 'Select cell
                                 If ret2 > -1 Then
                                     
                                    .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("J"), ret2
                                    '.BackColor = vbGreen
                                    .Refresh
                                 Else
                                   
                                    ret2 = .SearchCol(curencol, 21, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                       If ret2 > -1 Then
                                         .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("J"), ret2
                                         curenrow = ret2
                                         '.BackColor = vbGreen
                                         .Refresh
                                       Else
                                             'MsgBox "Khong co ban ghi nay.", vbInformation
                                           DisplayMessage "0160", msOKOnly, miInformation
                                       End If
                                 End If
                       End If
                       
                       
                   End If
                   
                 End If
                       
                 End If
                 
               ' To khai 06KK-TNCN
               If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "59" Then
                   If Trim(seach) = "" Then
                       DisplayMessage "0160", msOKOnly, miInformation
                       'txt_Seach.Text = ""
                       txt_Seach.SetFocus
                   Else
                        'If Col = .ColLetterToNumber("E") And Row = 11 Then
                           'T�m kiem c�c ban ghi giong nhau
                            If curenrow <= 21 Or curenrow > 21 + .MaxRows Then
                              curenrow = 21
                            End If

                            'seach = Trim(txt_seach.Text)
                            Option_Seach = Cb_seach.ListIndex
                            If Option_Seach = 0 Then
                               '.Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                ret2 = .SearchCol(.ColLetterToNumber("D"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                curenrow = ret2
                                curencol = .ColLetterToNumber("D")
                            ElseIf Option_Seach = 1 Then
                                    '.Sort .ColLetterToNumber("E"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                    ret2 = .SearchCol(.ColLetterToNumber("E"), curenrow, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                    curenrow = ret2
                                    curencol = .ColLetterToNumber("E")
                            ElseIf Option_Seach = 2 Then
                                            '.Sort .ColLetterToNumber("F"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                                            ret2 = -1
                                            curenrow = ret2
                                            curencol = .ColLetterToNumber("F")
                            End If
                            'Select cell
                            If ret2 > -1 Then

                               .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("I"), ret2
                               '.BackColor = vbGreen
                               .Refresh
                            Else
                                ret2 = .SearchCol(curencol, 21, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                                If ret2 > -1 Then
                                  .SetSelection .ColLetterToNumber("D"), ret2, .ColLetterToNumber("I"), ret2
                                  curenrow = ret2
                                  '.BackColor = vbGreen
                                  .Refresh
                                Else
                                      'MsgBox "Khong co ban ghi nay.", vbInformation
                                    DisplayMessage "0160", msOKOnly, miInformation
                                End If
                           End If
                   End If
               End If
            Else
      If .sheet = 3 Then
          If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Then
                If Trim(seach) = "" Then
                       DisplayMessage "0160", msOKOnly, miInformation
                       txt_Seach.SetFocus
                Else
                    'If Col = .ColLetterToNumber("E") And Row = 11 Then
                    'T�m kiem c�c ban ghi giong nhau
                     If curenrow <= 21 Or curenrow > 21 + .MaxRows Then
                       curenrow = 21
                     End If
                     
                     'seach = Trim(txt_seach.Text)
                     Option_Seach = Cb_seach.ListIndex
                     If Option_Seach = 0 Then
                            '.Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                             ret2 = .SearchCol(.ColLetterToNumber("C"), curenrow, (20 + .MaxRows), seach, SearchFlagsPartialMatch)
                             curenrow = ret2
                             curencol = .ColLetterToNumber("C")
                     ElseIf Option_Seach = 1 Then
                            '.Sort .ColLetterToNumber("E"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                            ret2 = .SearchCol(.ColLetterToNumber("D"), curenrow, (20 + .MaxRows), seach, SearchFlagsPartialMatch)
                            curenrow = ret2
                            curencol = .ColLetterToNumber("D")
                     ElseIf Option_Seach = 2 Then
                            '.Sort .ColLetterToNumber("F"), 22, .ColLetterToNumber("AG"), (21 + totalRow5A), SortByRow, Y, so
                            ret2 = .SearchCol(.ColLetterToNumber("E"), curenrow, (20 + .MaxRows), seach, SearchFlagsPartialMatch)
                            curenrow = ret2
                            curencol = .ColLetterToNumber("E")
                     End If
                     'Select cell
                     If ret2 > -1 Then
                         
                        .SetSelection .ColLetterToNumber("C"), ret2, .ColLetterToNumber("J"), ret2
                        '.BackColor = vbGreen
                        .Refresh
                     Else
                        ret2 = .SearchCol(curencol, 21, (21 + .MaxRows), seach, SearchFlagsPartialMatch)
                           If ret2 > -1 Then
                             .SetSelection .ColLetterToNumber("C"), ret2, .ColLetterToNumber("J"), ret2
                             curenrow = ret2
                             '.BackColor = vbGreen
                             .Refresh
                           Else
                                 'MsgBox "Khong co ban ghi nay.", vbInformation
                               DisplayMessage "0160", msOKOnly, miInformation
                           End If
                     End If
                    End If
                End If
            End If
       End If
      .EventEnabled(EventAllEvents) = True
    End With
End Sub

''' cmdClear_Click description
''' reset data in DOM and interface
''' No parameter
Private Sub cmdClear_Click()
    On Error GoTo ErrorHandle
    
    Dim strDataFileName As String
    Dim lResult As VbMsgBoxResult
    Dim lSheet As Long, lCol As Long, lRow As Long
    Dim loFile As New Scripting.FileSystemObject
    'test
    Dim lSheetActive, iCount As Integer
    Dim rowStart As Long
    Dim totalRow5A As Long
    
    Dim sumRowDel, ctlRow As Long
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    
    'end test
    lResult = DisplayMessage(IIf(mCurrentSheet = 1, "0030", "0035"), msYesNo, miQuestion)
    If lResult = mrYes Then
        ' check to quyet toan 05 thi khi nhap lai se xoa sheet active va them sheet do, cac to khac giu nguyen
        If Trim(GetAttribute(TAX_Utilities_New.NodeMenu, "ID")) = "17" And (fpSpread1.ActiveSheet = 2 Or fpSpread1.ActiveSheet = 3) Then
            lSheetActive = fpSpread1.ActiveSheet
            If fpSpread1.ActiveSheet = 2 Then
                rowStart = 22
            ElseIf fpSpread1.ActiveSheet = 3 Then
                rowStart = 22
            End If
            ' Xoa sheet active
            ' Set active attribute on xmlDoc menu
'            SetAttribute TAX_Utilities_New.NodeValidity.childNodes(lSheetActive - 1), "Active", "0"
'            ' Delete data file
'            DeleteSheet mCurrentSheet - 1
'            TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = False
'
'            ' Invisible current sheet
'            fpSpread1.SheetVisible = False
'
'            ' Them sheet active
'            fpSpread1.sheet = lSheetActive
'            If Not fpSpread1.SheetVisible Then
'                ResetDataAndForm lSheetActive
'                fpSpread1.SheetVisible = True
'                SetAttribute TAX_Utilities_New.NodeValidity.childNodes(lSheetActive - 1), "Active", "1"
'                TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
'            End If
'             ' Set active sheet
'            fpSpread1.ActiveSheet = lSheetActive
'            fpSpread1.sheet = lSheetActive
'            fpSpread1.Refresh

            With fpSpread1
                .EventEnabled(EventChange) = False
                .ReDraw = False
                '.Visible = False
                .sheet = lSheetActive
                iCount = 1
                Do
                    .Col = .ColLetterToNumber("B")
                    .Row = iCount + rowStart
                    totalRow5A = iCount
                    iCount = iCount + 1
                Loop Until .Text = "aa"
                
                .DeleteRows rowStart + 1, totalRow5A - 1
                .MaxRows = .MaxRows - totalRow5A + 1
                '.Visible = True
                .ReDraw = True
                
                
                

         
            
                sumRowDel = TAX_Utilities_New.Data(lSheetActive - 1).getElementsByTagName("Cells").length
                For ctlRow = 1 To sumRowDel - 1
                    If lSheetActive = 2 Then
                        Set xmlNodeCells = TAX_Utilities_New.Data(lSheetActive - 1).nodeFromID("D" & "_" & 22 + ctlRow & "").parentNode
                        xmlNodeCells.parentNode.removeChild xmlNodeCells
                    ElseIf lSheetActive = 3 Then
                        Set xmlNodeCells = TAX_Utilities_New.Data(lSheetActive - 1).nodeFromID("C" & "_" & 21 + ctlRow & "").parentNode
                        xmlNodeCells.parentNode.removeChild xmlNodeCells
                    End If
                Next
                
                ResetData
                SetActiveFirstCell lSheet, lCol, lRow
                ResetErrorCells
                .EventEnabled(EventChange) = True
            End With
        ElseIf Trim(GetAttribute(TAX_Utilities_New.NodeMenu, "ID")) = "59" And fpSpread1.ActiveSheet = 2 Then
            lSheetActive = fpSpread1.ActiveSheet
            If fpSpread1.ActiveSheet = 2 Then
                rowStart = 22
            End If
            With fpSpread1
                .EventEnabled(EventChange) = False
                .ReDraw = False
                '.Visible = False
                .sheet = lSheetActive
                iCount = 1
                Do
                    .Col = .ColLetterToNumber("B")
                    .Row = iCount + rowStart
                    totalRow5A = iCount
                    iCount = iCount + 1
                Loop Until .Text = "aa"
                
                .DeleteRows rowStart + 1, totalRow5A - 1
                .MaxRows = .MaxRows - totalRow5A + 1
                '.Visible = True
                .ReDraw = True
                
            
                sumRowDel = TAX_Utilities_New.Data(lSheetActive - 1).getElementsByTagName("Cells").length
                For ctlRow = 1 To sumRowDel - 1
                    If lSheetActive = 2 Then
                        Set xmlNodeCells = TAX_Utilities_New.Data(lSheetActive - 1).nodeFromID("D" & "_" & 22 + ctlRow & "").parentNode
                        xmlNodeCells.parentNode.removeChild xmlNodeCells
                    End If
                Next
                
                ResetData
                SetActiveFirstCell lSheet, lCol, lRow
                ResetErrorCells
                .EventEnabled(EventChange) = True
            End With
 
        Else
            fpSpread1.EventEnabled(EventAllEvents) = False
            ResetData
            SetActiveFirstCell lSheet, lCol, lRow
            ResetErrorCells
            fpSpread1.EventEnabled(EventAllEvents) = True
        End If
    End If
    fpSpread1_EditMode lCol, lRow, 0, False
    fpSpread1.Refresh
    fpSpread1.SetFocus
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "cmdClear_Click", Err.Number, Err.Description
End Sub

Private Sub DeleteSheet(pIndex As Integer)
    On Error GoTo ErrorHandle
    
    Dim strDataFileName As String
    Dim loFile As New Scripting.FileSystemObject
    ' TO khai TTDB va NTNN, 02/TNDN xu ly xoa lan phat sinh
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "05" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "70" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "91" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "64" Then
        If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And TAX_Utilities_New.Day = "" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
        End If
        ' end
    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "73" Then
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" And TAX_Utilities_New.Day = "" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
        End If
    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "74" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "75" Then
        If strQuy = "TK_TU_THANG" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
        Else
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
        End If
    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "68" Then
        'BC26
        If strQuy = "TK_THANG" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_T" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
        ElseIf strQuy = "TK_QUY" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
        End If
    Else
        If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
            If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "95" _
            Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "36" Then
                If strQuy = "TK_THANG" Then
                    strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                ElseIf strQuy = "TK_QUY" Then
                    strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_Q0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                End If
            Else
                strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
            End If
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
            strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" Then
            If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "80" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "82" Then
                strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & _
                Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
            Else
                strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" & TAX_Utilities_New.Year _
                & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
            End If
        Else
                strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(pIndex), "DataFile") & "_" _
                    & TAX_Utilities_New.Year & ".xml"
        '**********************************
        End If
    End If
    
    If loFile.FileExists(strDataFileName) = True Then
        loFile.DeleteFile strDataFileName, True
    End If
    
    Set loFile = Nothing
    '*******************
    'Remove Active attribute
    SetAttribute TAX_Utilities_New.NodeValidity.childNodes(pIndex), "Active", "0"
    '*******************
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "DeleteSheet", Err.Number, Err.Description
    If Err.Number = 70 Then
        DisplayMessage "0038", msOKOnly, miCriticalError
    Else
    End If
End Sub

Private Sub DeleteKHBS()
    On Error GoTo ErrorHandle
    Dim lSheet As Long
    Dim strKHBSDataFileName As String
    Dim strSheetKHBSDataFileName As String
    Dim loFile As New Scripting.FileSystemObject
    'tinh datafile sheet strSheetKHBSDataFileName
    If strKHBS = "TKBS" Then
                If GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = "0" Then
                    strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                    'Ten Sheet KHBS
                    strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                Else
                    If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "0" Then
                        If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "95" _
                            Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "36" Then
                                If strQuy = "TK_THANG" Then
                                    strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                                    strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                                ElseIf strQuy = "TK_QUY" Then
                                    strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                                    strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                                End If
                        End If
                    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
                        If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "75" Then
                                    ' To khai 08/TNCN co to khai tu thang va to khai quy
                                    If strQuy = "TK_TU_THANG" Then
                                        strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                                        strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                                    Else
                                        strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                                        strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                                    End If
                        Else
                            strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                            strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & ".xml"
                        End If
                    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "0" Then
                            ' To khai 02/NTNN, 04/NTNN
                            If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "82" Then
                                 strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                                 strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                            Else
                                 strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                                
                                strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & ".xml"
                            End If
                    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
                            strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                            
                            strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & ".xml"
                    Else
                            strKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Year & ".xml"
                            
                            strSheetKHBSDataFileName = TAX_Utilities_New.DataFolder & "bs" & strSolanBS & "_KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Year & ".xml"
                    '*********************************
                    End If
                End If
    Else
                If GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = "0" Then
                    strKHBSDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                Else
                    If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "0" Then
                        strKHBSDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
                        strKHBSDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "0" Then
                             strKHBSDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
                            strKHBSDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                    Else
                            strKHBSDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                    '*********************************
                    End If
                End If
    End If
    
    
    If loFile.FileExists(strKHBSDataFileName) = True Then
        loFile.DeleteFile strKHBSDataFileName, True
    End If
    'Xoa Sheet KHBS cac TK BS (GTGT,TTDB,TAIN,01A-01B/TNDN)
    Dim varMenuId1 As Variant
    varMenuId1 = TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    If strKHBS <> "TKBS" Then
        strKHBSDataFileName = Replace(strKHBSDataFileName, "KHBS", "KHBS1")
        If loFile.FileExists(strKHBSDataFileName) = True Then
            loFile.DeleteFile strKHBSDataFileName, True
        End If
    ElseIf strKHBS = "TKBS" And (varMenuId1 = "02" Or varMenuId1 = "01" Or varMenuId1 = "04" Or varMenuId1 = "11" Or varMenuId1 = "12" Or varMenuId1 = "05" Or varMenuId1 = "06" Or varMenuId1 = "86" Or varMenuId1 = "87" _
    Or varMenuId1 = "89" Or varMenuId1 = "71" Or varMenuId1 = "72" Or varMenuId1 = "77" Or varMenuId1 = "03" Or varMenuId1 = "73" Or varMenuId1 = "80" Or varMenuId1 = "81" Or varMenuId1 = "70" Or varMenuId1 = "82" Or varMenuId1 = "83" _
    Or varMenuId1 = "85") Then
        If loFile.FileExists(strSheetKHBSDataFileName) = True Then
            loFile.DeleteFile strSheetKHBSDataFileName, True
        End If
    End If
    Set loFile = Nothing
    '*******************
    'Remove Active attribute
    'SetAttribute TAX_Utilities_New.NodeValidity.childNodes(pIndex), "Active", "0"
    '*******************
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "DeleteKHBS", Err.Number, Err.Description
    If Err.Number = 70 Then
        DisplayMessage "0038", msOKOnly, miCriticalError
    Else
    End If
End Sub



''' cmdDelete_Click description
''' Delete data file on disk, exit function after delete
''' No parameter
Private Sub cmdDelete_Click()
    On Error GoTo ErrorHandle

    Dim lResult As VbMsgBoxResult
    Dim lSheet As Integer
    Dim lCol As Long, lRow As Long
    
    If strKHBS = "frmKHBS_BS" Or strKHBS = "TKBS" Then
        lResult = DisplayMessage("0012", msYesNo, miQuestion, , mrNo)
        If lResult = mrYes Then
            ' Delete data files
                DeleteKHBS
            strInterfaceUnloadEventName = "Delete"
            Unload Me
            Exit Sub
        End If
    End If
    ' Truong hop doi voi cac to khai quyet toan thi khi chon xoa phai xoa tat ca to khai luon
    'vtt sua them ID 59
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "41" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "42" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "43" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "59" Then
        mCurrentSheet = 1
        lResult = DisplayMessage("0012", msYesNo, miQuestion, , mrNo)
        If lResult = mrYes Then
            ' Delete data files
            For lSheet = 0 To TAX_Utilities_New.xmlDataCount
                DeleteSheet lSheet
            Next
            strInterfaceUnloadEventName = "Delete"
            Unload Me
        End If
        Exit Sub
    End If
    
    If mCurrentSheet = 1 Then
    'If GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(mCurrentSheet - 1), "Active") = "2" Then
    '**************************************
        lResult = DisplayMessage("0012", msYesNo, miQuestion, , mrNo)
        If lResult = mrYes Then
            ' Delete data files
            For lSheet = 0 To TAX_Utilities_New.xmlDataCount
                DeleteSheet lSheet
            Next
            strInterfaceUnloadEventName = "Delete"
            Unload Me
            Exit Sub
        End If
    Else
        lResult = DisplayMessage("0032", msYesNo, miQuestion)
        If lResult = mrYes Then
            '****************************
            ' added
            'Reset all of error status in screen
            ResetErrorCells
            '****************************

            ' Reset data on screen and in xmlDoc object
            ResetData

            ' Set active attribute on xmlDoc menu
            SetAttribute TAX_Utilities_New.NodeValidity.childNodes(mCurrentSheet - 1), "Active", "0"

            ' Delete data file
            DeleteSheet mCurrentSheet - 1
            TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = False
            
            ' Invisible current sheet
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.SheetVisible = False
            
            ' Set active sheet
            fpSpread1.ActiveSheet = 1
            fpSpread1.sheet = 1
            fpSpread1.Refresh
            
            If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
                If Not objTaxBusiness Is Nothing Then
                     objTaxBusiness.LockCellBySheet
                End If
            End If
            ' To khai 03_TNDN
            If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "03" Then
                If Not objTaxBusiness Is Nothing Then
                         objTaxBusiness.unLockCellPL (objTaxBusiness.strloaitk)
                End If
            End If
            
            ' To khai 04_NTNN
            If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "82" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "80" Then
                If Not objTaxBusiness Is Nothing Then
                         objTaxBusiness.updateSomeCell
                End If
            End If
                   ' TK 01 TD
            If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "83" Then
                If Not objTaxBusiness Is Nothing Then
                    objTaxBusiness.LockCellBySheet
                End If
            End If
            Exit Sub
        End If
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "cmdDelete_Click", Err.Number, Err.Description
End Sub

''' cmdExit_Click description
''' If user change data, ask for save data -> exit
''' No parameter
Private Sub cmdExit_Click()
    On Error GoTo ErrorHandle
    Dim mr As Integer
    
    ' Neu la cac mau in tong hop tu to quyet toan 05TNCN->09TNCN va cac chung tu cua TNCN thi thoat luon!
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "45" Then
        Unload Me
        strInterfaceUnloadEventName = "Exit"
        Exit Sub
    End If
    
    CallFinish True
    If IsAdjustData = True Then
        mr = DisplayMessage("0013", msYesNoCancel, miQuestion)
        If mr = mrYes Then
            If Not objTaxBusiness Is Nothing Then objTaxBusiness.finish
             If strKHBS = "frmKHBS_BS" Then
                    saveKHBS
                    Unload Me
                    Exit Sub
             End If
            If CheckValidData = False Then
                If DisplayMessage("0014", msYesNo, miQuestion) = mrNo Then
                    If UpdateData Then _
                        Unload Me
                Else
                    UpdateData
                End If
            Else
                If UpdateData Then _
                    Unload Me
            End If
        ElseIf mr = mrNo Then
            If Not TAX_Utilities_New.DataKHBS Is Nothing Then TAX_Utilities_New.DataKHBS = Nothing
            Unload Me
        ElseIf mr = mrCancel Then
        
        End If
    Else
        If Not TAX_Utilities_New.DataKHBS Is Nothing Then TAX_Utilities_New.DataKHBS = Nothing
        Unload Me
    End If
    
    strInterfaceUnloadEventName = "Exit"
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "cmdExit_Click", Err.Number, Err.Description
End Sub

''' cmdExport_Click description
''' Export data to compress string data file
''' File name:  FunctionID(len = 3) & Period(len = 2) & Year(len = 4)
'''             & TaxID(len = 10 or 13) & ".xml" (ext file name)
''' This function convert from unicode to TCVN, compress string data before export to file
''' No parameter
Private Sub cmdExport_Click()
'    On Error GoTo ErrorHandle
'    Dim strFolder As String
'    Dim strDataFileName As String
'    Dim strTaxId As String
'    Dim strAllData As String
'    Dim i As Long, lErrNumber As Long
'    Dim loFile As New Scripting.FileSystemObject
'    Dim loTextStream As Scripting.TextStream
'    Dim strArrActive() As String
'
'    '*****************************
'    'Backup node validity
'    For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'        ReDim Preserve strArrActive(i)
'        strArrActive(i) = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(i), "Active")
'    Next i
'    'Set active sheet property
'    If Not objTaxBusiness Is Nothing Then
'        'For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'            Call objTaxBusiness.SetActiveSheet '(TAX_Utilities_New.NodeValidity.childNodes(i))
'        'Next i
'    End If
'    '*****************************
'
'    CallFinish
'    If CheckValidData = True Then
'        ' Get folder to export
'        strFolder = frmBrowser.getPath
'    Else
'        DisplayMessage "0039", msOKOnly, miInformation
'        Exit Sub
'    End If
'
'    If strFolder = vbNullString Then Exit Sub
'    ' Get datafile name only
'    strDataFileName = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
'
'    ' Get period
'    If Val(TAX_Utilities_New.month) <> 0 Then
'        ' Get month
'        strDataFileName = strDataFileName & TAX_Utilities_New.month & TAX_Utilities_New.Year
'    Else
'        ' Get threemonths
'        strDataFileName = strDataFileName & "0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year
'    End If
'
'    For i = 0 To 12
'        ' Get Tax code
'        strTaxId = strTaxId & IIf(GetAttribute(xmlHeaderData.getElementsByTagName("Cell")(i), "Value") = "", " ", GetAttribute(xmlHeaderData.getElementsByTagName("Cell")(i), "Value"))
'    Next
'
'    CreateExcelBook
'    strDataFileName = strDataFileName & Trim(strTaxId)
'    For i = 0 To UBound(strDataBarcode)
'        strAllData = strAllData & strDataBarcode(i)
'    Next
'    If Val(TAX_Utilities_New.month) <> 0 Then
'        strAllData = GetAttribute(TAX_Utilities_New.NodeMenu, "ID") & strTaxId & TAX_Utilities_New.month & TAX_Utilities_New.Year & strAllData
'    Else
'        strAllData = GetAttribute(TAX_Utilities_New.NodeMenu, "ID") & strTaxId & "0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & strAllData
'    End If
'
'    strDataFileName = strFolder & "\" & strDataFileName & ".txt"
'
'
'    Set loTextStream = loFile.CreateTextFile(strDataFileName, True, True)
'    ' Create data string
'    ' Convert unicode to TCVN (user ABC font)
'    loTextStream.WriteLine strAllData 'TAX_Utilities_New.Convert(strAllData, UNICODE, TCVN) 'TAX_Utilities_New.Compress(strAllData)
'    loTextStream.Close
'
'    Set loTextStream = Nothing
'    Set loFile = Nothing
'
'    DisplayMessage "0022", msOKOnly, miInformation
'
'    '*****************************
'    ' added
'    'Modify date: 05/12/2005
'
'    'Restore active properties of node validity
'    For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(i), "Active", strArrActive(i)
'    Next i
'    '*****************************
'    Exit Sub
'
'ErrorHandle:
'
'    '*****************************
'    ' added
'    'Modify date: 05/12/2005
'    lErrNumber = Err.Number
'    On Error GoTo ErrExit
'    'Restore active properties of node validity
'    For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(i), "Active", strArrActive(i)
'    Next i
'    '*****************************
'
'    Set loFile = Nothing
'    Set loTextStream = Nothing
'    If lErrNumber = 70 Then
'        DisplayMessage "0036", msOKOnly, miCriticalError
'        Err.Clear
'        cmdExport_Click
'    ElseIf lErrNumber = -2147024784 Then
'        DisplayMessage "0037", msOKOnly, miCriticalError
'        Err.Clear
'        cmdExport_Click
'    Else
'        SaveErrorLog Me.Name, "cmdExport_Click", lErrNumber, Err.Description
'    End If
'    Exit Sub
'ErrExit:
'    SaveErrorLog Me.Name, "cmdExport_Click", Err.Number, Err.Description
    Dim strFileName As String
    Dim strValue As Variant
    Dim cFolder As New Scripting.FileSystemObject
    Dim nFolder As String
    Dim nExcelFile As String
    
    Dim idToKhaiKHBS As String
    
    
    On Error GoTo DialogError
    
    flgloadToKhai = False
    ' Doi voi cac to quyet toan TNCN thi dat co flgloadToKhai = false
    ' Muc dich la trong truong hop load bang ke thi ko tong hop lai du lieu
    ' Khi Ghi, In, Ket xuat thi dat lai trang thai
    If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        objTaxBusiness.flgloadToKhai = flgloadToKhai
    End If
    ' To khai quyet toan thue TNCN
    If GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "05_TNCN" _
        Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "06_TNCN10" _
            Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "09_TNCN" _
                Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_BH" _
                    Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_XS" _
                        Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "08B_TNCN" Then
        ' Doi voi to khai quyet toan thue TNCN thi export ra thu muc C:\TNCN-Temp
        nExcelFile = prepareFileName(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile"))
        If (cFolder.FolderExists("C:\TNCN-Temp")) = False Then
            nFolder = "C:\TNCN-Temp"
            cFolder.CreateFolder nFolder
        Else
            nFolder = "C:\TNCN-Temp"
        End If
    Else
        nExcelFile = prepareFileName(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile"))
        ' Doi voi cac to khai khac khi export thi ra thu muc C:\HTTK-Temp
        If (cFolder.FolderExists("C:\HTTK-Temp")) = False Then
            nFolder = "C:\HTTK-Temp"
            cFolder.CreateFolder nFolder
        Else
            nFolder = "C:\HTTK-Temp"
        End If
    End If
    CallFinish
    
    ' nkhoan: 02/TNDN
    If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73") Then
        If objTaxBusiness.iflag = True Then
            DisplayMessage "0240", msOKOnly, miCriticalError
            Exit Sub
       End If
    End If
    
    If CheckValidData = True Then
        With CommonDialog1
            .CancelError = True
            .InitDir = nFolder
            .Filter = "Excel file (*.xls)|*.xls"
            .FilterIndex = 1
            .DialogTitle = "File excel export to " & nFolder
            .FileName = nExcelFile
            .ShowSave
            If Right$(.FileName, 4) <> ".xls" Then
                strFileName = .FileName & ".xls"
            Else
                strFileName = .FileName
            End If
        End With
    
        On Error GoTo ErrHandle
        
        If GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "01_GTGT" Then
            With fpSpread1
                fpSpread1.EventEnabled(EventAllEvents) = False
                .GetText .ColLetterToNumber("I"), 23, strValue
                If strValue = 1 Then
                    .Col = .ColLetterToNumber("I")
                    .Row = 23
                    .Text = "X"
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .Col = .ColLetterToNumber("I")
                    .Row = 23
                    .Text = ""
                End If
                
                
                fpSpread1.EventEnabled(EventAllEvents) = True
            End With
        End If
        
        If GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "05_TNCN" Then
            With fpSpread1
                .EventEnabled(EventAllEvents) = False
                .sheet = 2
                .Row = 22
                Do
                    .GetText .ColLetterToNumber("G"), .Row, strValue
                    If Trim(strValue) = "1" Or Trim(strValue) = "x" Then
                        ' start
                        .Col = .ColLetterToNumber("Q")
                        .Formula = ""
                        ' end
                        
                        .Col = .ColLetterToNumber("G")
                        .Text = "x"
                        .TypeHAlign = TypeHAlignCenter
                    Else
                        ' start
                        .Col = .ColLetterToNumber("Q")
                        .Formula = ""
                        ' end
                        .Col = .ColLetterToNumber("G")
                        .Text = ""
                    End If
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "aa"
                .Col = .ColLetterToNumber("C")
                .ColHidden = True
                
                .Col = .ColLetterToNumber("D")
                .Row = 5
                .Text = ""
                
                .sheet = 3
                .Row = 22
                Do
                    .GetText .ColLetterToNumber("F"), .Row, strValue
                    If Trim(strValue) = "1" Or Trim(strValue) = "x" Then
                        .Col = .ColLetterToNumber("F")
                        .Text = "x"
                        .TypeHAlign = TypeHAlignCenter
                    Else
                        .Col = .ColLetterToNumber("F")
                        .Text = ""
                    End If
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "aa"
                
                .Col = .ColLetterToNumber("Y")
                .ColHidden = True
                
                .Col = .ColLetterToNumber("C")
                .Row = 4
                .Text = ""
                
                
                .EventEnabled(EventAllEvents) = True
            End With
        ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_BH" Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_XS" Then
            With fpSpread1
                .EventEnabled(EventAllEvents) = False
                .sheet = 2
                .Row = 22
                
                .Col = .ColLetterToNumber("C")
                .ColHidden = True
                        
                .Col = .ColLetterToNumber("D")
                .Row = 3
                .Text = ""
                        
                .EventEnabled(EventAllEvents) = True
            End With
        ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "BC26_AC" Then
            With fpSpread1
                .EventEnabled(EventAllEvents) = False
                .sheet = 1
                .GetText .ColLetterToNumber("B"), 14, strValue
                If strValue = 1 Or UCase$(CStr(strValue)) = "X" Then
                    .Col = .ColLetterToNumber("B")
                    .Row = 14
                    .Text = "X"
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .Col = .ColLetterToNumber("B")
                    .Row = 14
                    .Text = ""
                End If
                
                .GetText .ColLetterToNumber("G"), 14, strValue
                If strValue = 1 Or UCase$(CStr(strValue)) = "X" Then
                    .Col = .ColLetterToNumber("G")
                    .Row = 14
                    .Text = "X"
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .Col = .ColLetterToNumber("G")
                    .Row = 14
                    .Text = ""
                End If
                
                .EventEnabled(EventAllEvents) = True
            End With
        ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "01_AC" Then
            With fpSpread1
                fpSpread1.EventEnabled(EventAllEvents) = False
                fpSpread1.sheet = 1
                .GetText .ColLetterToNumber("B"), 15, strValue
                If strValue = 1 Or UCase$(CStr(strValue)) = "X" Then
                    .Col = .ColLetterToNumber("B")
                    .Row = 15
                    .Text = "X"
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .Col = .ColLetterToNumber("B")
                    .Row = 15
                    .Text = ""
                End If
                
                fpSpread1.EventEnabled(EventAllEvents) = True
            End With
        End If
        
        ' Kiem tra neu khac to khai co KHBS thi moi xoa
        idToKhaiKHBS = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
        If idToKhaiKHBS <> "01" And idToKhaiKHBS <> "02" And idToKhaiKHBS <> "03" And idToKhaiKHBS <> "04" And idToKhaiKHBS <> "05" _
        And idToKhaiKHBS <> "06" And idToKhaiKHBS <> "08" And idToKhaiKHBS <> "11" And idToKhaiKHBS <> "12" And idToKhaiKHBS <> "86" And idToKhaiKHBS <> "87" _
        And idToKhaiKHBS <> "89" And idToKhaiKHBS <> "71" And idToKhaiKHBS <> "72" And idToKhaiKHBS <> "77" And idToKhaiKHBS <> "03" And idToKhaiKHBS <> "73" _
        And idToKhaiKHBS <> "80" And idToKhaiKHBS <> "81" And idToKhaiKHBS <> "70" And idToKhaiKHBS <> "82" And idToKhaiKHBS <> "83" And idToKhaiKHBS <> "85" Then
                fpSpread1.sheet = fpSpread1.SheetCount - 1
                If fpSpread1.SheetName = "KHBS" Then
                    fpSpread1.DeleteSheets fpSpread1.SheetCount - 1, 1
                End If
                
                ' Khac ky tinh thue nam 2012 se khong ket xuat pl MT 26,27
                ' To khai 09/KK-TNCN
                If idToKhaiKHBS = "41" And TAX_Utilities_New.Year <> "2012" Then
                    fpSpread1.sheet = 5
                    If fpSpread1.SheetName = "26MT-TNCN" Then
                        fpSpread1.DeleteSheets 5, 1
                    End If
                End If
                ' To khai 05/KK-TNCN
                
                 If idToKhaiKHBS = "17" And TAX_Utilities_New.Year <> "2012" Then
                    fpSpread1.sheet = 4
                    If fpSpread1.SheetName = "27MT-TNCN" Then
                        fpSpread1.DeleteSheets 4, 1
                    End If
                 End If
                 
                 ' chi ket xuat dong tong
                 If idToKhaiKHBS = "17" And TAX_Utilities_New.Year = "2012" Then
                    fpSpread1.EventEnabled(EventAllEvents) = False
                    fpSpread1.sheet = 4
                    Dim countRowDel As Integer
                    countRowDel = 0
                    fpSpread1.Row = 22
                    fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                    Do
                        countRowDel = countRowDel + 1
                        fpSpread1.Row = countRowDel + 22
                    Loop Until fpSpread1.Text = "aa"
                    
                    Dim arrData5A(1, 5) As Variant
                    fpSpread1.Row = fpSpread1.Row + 1
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("F"), fpSpread1.Row, arrData5A(1, 0)
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("G"), fpSpread1.Row, arrData5A(1, 1)
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("H"), fpSpread1.Row, arrData5A(1, 2)
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("I"), fpSpread1.Row, arrData5A(1, 3)
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("J"), fpSpread1.Row, arrData5A(1, 4)
                    
                    fpSpread1.DeleteRows 22, countRowDel
                    fpSpread1.MaxRows = fpSpread1.MaxRows - countRowDel + 1

                    fpSpread1.sheet = 4
'                    fpSpread1.Row = 22
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("C"), fpSpread1.Row, ""
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("D"), fpSpread1.Row, ""
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("E"), fpSpread1.Row, ""
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("F"), fpSpread1.Row, "0"
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("G"), fpSpread1.Row, "0"
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("H"), fpSpread1.Row, "0"
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("I"), fpSpread1.Row, "0"
'                    fpSpread1.SetText fpSpread1.ColLetterToNumber("J"), fpSpread1.Row, "0"
                    
                    fpSpread1.Row = 23
                    fpSpread1.SetText fpSpread1.ColLetterToNumber("F"), fpSpread1.Row, arrData5A(1, 0)
                    fpSpread1.SetText fpSpread1.ColLetterToNumber("G"), fpSpread1.Row, arrData5A(1, 1)
                    fpSpread1.SetText fpSpread1.ColLetterToNumber("H"), fpSpread1.Row, arrData5A(1, 2)
                    fpSpread1.SetText fpSpread1.ColLetterToNumber("I"), fpSpread1.Row, arrData5A(1, 3)
                    fpSpread1.SetText fpSpread1.ColLetterToNumber("J"), fpSpread1.Row, arrData5A(1, 4)
                    fpSpread1.EventEnabled(EventAllEvents) = True
                End If
                
                'ngay 04/03/2011
                'sua lai gia tri sheet count de check validate
                mHeaderSheet = fpSpread1.SheetCount
                ' end test
        End If
        
        fpSpread1.ExportExcelBookEx strFileName, vbNullString, ExcelSaveFlagNoFormulas 'App.path & "\ExportLog.log"
        
        ' chi ket xuat dong tong
         If idToKhaiKHBS = "17" And TAX_Utilities_New.Year = "2012" Then
             fpSpread1.sheet = 4
             fpSpread1.EventEnabled(EventAllEvents) = False
             fpSpread1.InsertRows 22, 1
             fpSpread1.CopyRowRange 21, 21, 22
             'fpSpread1.MaxRows = fpSpread1.MaxRows + 1
             fpSpread1.Row = 22
             ' set cong thuc cho chi tieu 12
             fpSpread1.Col = fpSpread1.ColLetterToNumber("H")
             fpSpread1.Formula = "ROUND(L22-M22,0)"
             ' set cong thuc cho chi tieu 13
             fpSpread1.Col = fpSpread1.ColLetterToNumber("I")
             fpSpread1.Formula = "H22 -J22"
             ' set cong thuc cho col K
             fpSpread1.Col = fpSpread1.ColLetterToNumber("K")
             fpSpread1.Formula = "F22/12"
             ' set cong thuc cho col L
             fpSpread1.Col = fpSpread1.ColLetterToNumber("L")
             fpSpread1.Formula = "12*IF(K22>80000000,((K22-80000000)*0.35+18150000),IF(AND(K22>52000000,K22<=80000000),((K22-52000000)*0.3+9750000),IF(AND(K22>32000000,K22<=52000000),((K22-32000000)*0.25+4750000),IF(AND(K22>18000000,K22<=32000000),((K22-18000000)*0.2+1950000),IF(AND(K22>10000000,K22<=18000000),((K22-10000000)*0.15+750000),IF(AND(K22>5000000,K22<=10000000),((K22-5000000)*0.1+250000),(K22*0.05)))))))"
             ' set cong thuc cho col M
             fpSpread1.Col = fpSpread1.ColLetterToNumber("M")
             fpSpread1.Formula = "IF(N22>0,ROUND(L22*O22/N22/2,0),0)"
             fpSpread1.RowHidden = False
             
             fpSpread1.EventEnabled(EventAllEvents) = True
             If objTaxBusiness Is Nothing Then
                 Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_New.NodeValidity, "Class"))
             End If
             'objTaxBusiness.CheckDynamicSheet1
             objTaxBusiness.fillDataPL27
         End If
        
       
        
'        ' Chu y khi ket xuat bo Mau Excel Bang Ke, Chu y la co the bo cho nay
'        With fpSpread1
'            .EventEnabled(EventAllEvents) = False
'                If GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_BH" Or GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "05_TNCN" Then
'                    .sheet = 2
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 5
'                    .Text = TAX_Utilities_New.Convert(GetMessageCellById("0174"), TCVN, UNICODE)
'
'                    .sheet = 3
'                    .Col = .ColLetterToNumber("C")
'                    .Row = 4
'                    .Text = TAX_Utilities_New.Convert(GetMessageCellById("0174"), TCVN, UNICODE)
'                End If
'            .EventEnabled(EventAllEvents) = True
'        End With
        
        
        DisplayMessage "0153", msOKOnly, miInformation
    Else
        DisplayMessage "0140", msOKOnly, miInformation
    End If

    Exit Sub
DialogError:
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdExport_Click", Err.Number, Err.Description
End Sub

Private Function prepareFileName(ByVal loaiToKhai As String) As String
    
    Dim nExcelFile As String
    Dim taxOfficeName As Variant
    Dim taxId As Variant
    Dim kyHieuToKhai As Variant
    Dim ctBs As Variant
    Dim lanBS As Variant
    ' Setup ten file theo tung to khai
    If UCase(Trim(loaiToKhai)) = "05_TNCN" _
            Or UCase(Trim(loaiToKhai)) = "02_TNCN_XS" _
                Or UCase(Trim(loaiToKhai)) = "02_TNCN_BH" _
                   Or UCase(Trim(loaiToKhai)) = "06_TNCN10" _
                      Or UCase(Trim(loaiToKhai)) = "09_TNCN" _
                        Or UCase(Trim(loaiToKhai)) = "08B_TNCN" Then
        With fpSpread1
            .EventEnabled(EventAllEvents) = False
            .sheet = 1
            .Col = .ColLetterToNumber("D")
            .Row = 10
            ' Lay MST cua NNT
            .GetText .Col, .Row, taxId
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(taxId))) = 10 Then
                taxId = Left(taxId, 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(Trim(taxId)) = 14 Then
                taxId = Replace(Trim(taxId), "-", "")
            End If
            ' Lay ky hieu to khai theo quy dinh cua PIT
            If UCase(Trim(loaiToKhai)) = "05_TNCN" Then
                kyHieuToKhai = "05TL"
            ElseIf UCase(Trim(loaiToKhai)) = "02_TNCN_BH" Then
                kyHieuToKhai = "02BH"
            ElseIf UCase(Trim(loaiToKhai)) = "02_TNCN_XS" Then
                kyHieuToKhai = "02XS"
            ElseIf UCase(Trim(loaiToKhai)) = "06_TNCN10" Then
                kyHieuToKhai = "06KK"
            ElseIf UCase(Trim(loaiToKhai)) = "09_TNCN" Then
                kyHieuToKhai = "09CN"
            ElseIf UCase(Trim(loaiToKhai)) = "08B_TNCN" Then
                kyHieuToKhai = "08BN"
            End If
            ' Lay trang thai va so lan cua to khai
            .Col = .ColLetterToNumber("E")
            .Row = 6
            .GetText .Col, .Row, ctBs
            If UCase(Trim(ctBs)) = "[X]" Then
                ctBs = "L00"
            Else
                .Col = .ColLetterToNumber("I")
                .Row = 6
                .GetText .Col, .Row, lanBS
                If Len(lanBS) = 1 Then
                    ctBs = "L" & "0" & lanBS
                Else
                    ctBs = "L" & lanBS
                End If
            End If
            ' Lay ma co quan thue cap cuc
            .Col = .ColLetterToNumber("D")
            If UCase(Trim(loaiToKhai)) = "05_TNCN" Then
                    .Row = 30
            ElseIf UCase(Trim(loaiToKhai)) = "09_TNCN" Then
                    .Row = 36
            Else
                     .Row = 34
            End If
            .GetText .Col, .Row, taxOfficeName
            taxOfficeName = Left$(taxOfficeName, 3)
            
            ' Ghep cac thong tin tren vao lam cau truc file chuan cho PIT
            nExcelFile = taxOfficeName & "-" & taxId & "-" & kyHieuToKhai & "-" & "Y" & TAX_Utilities_New.Year & "-" & ctBs
            .EventEnabled(EventAllEvents) = True
        End With
        prepareFileName = nExcelFile
    ElseIf UCase(Trim(loaiToKhai)) = "BC26_AC" Or UCase(Trim(loaiToKhai)) = "01_AC" Then
        With fpSpread1
            .EventEnabled(EventAllEvents) = False
            .sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 9
            ' Lay MST cua NNT
            .GetText .Col, .Row, taxId
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(taxId))) = 10 Then
                taxId = Left(taxId, 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(Trim(taxId)) = 14 Then
                taxId = Replace(Trim(taxId), "-", "")
            End If
            
            If UCase(Trim(loaiToKhai)) = "BC26_AC" Then
                kyHieuToKhai = "BC26AC"
            ElseIf UCase(Trim(loaiToKhai)) = "01_AC" Then
                kyHieuToKhai = "BC01AC"
            End If
            ' Lay ma co quan thue cap cuc
'            .sheet = .SheetCount
'            .Col = .ColLetterToNumber("R")
'            .Row = 8
'            .GetText .Col, .Row, taxOfficeName
'            taxOfficeName = Left$(taxOfficeName, 3)
            
            If UCase(Trim(loaiToKhai)) = "BC26_AC" Or UCase(Trim(loaiToKhai)) = "01_AC" Then
                nExcelFile = taxId & "-" & kyHieuToKhai & "-0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year
            End If
            
            .EventEnabled(EventAllEvents) = True
            
        End With
        prepareFileName = nExcelFile
    Else
        ' Doi voi cac to khai khac khi export ma la to khai thang
        If Val(TAX_Utilities_New.month) > 0 Then
            prepareFileName = loaiToKhai & "_" & TAX_Utilities_New.Year
        ' Doi voi cac to khai khac khi export ma la to khai quy
        ElseIf Val(TAX_Utilities_New.ThreeMonths) > 0 Then
            prepareFileName = loaiToKhai & "_" & TAX_Utilities_New.Year
        Else ' Truong hop la quyet toan hoac theo nam tai chinh
            prepareFileName = loaiToKhai & "_" & TAX_Utilities_New.Year
        End If
        
    End If
End Function

Private Sub cmdInsert_Click()
'**********************************************
'Noi dung thay doi: Thuc hien chuc nang them phu
'                   luc khi dang ke khai.
'**********************************************
    Dim intCtrl As Integer
    Dim strAddedSheet As String
    Dim strSheets As String, strSelectedSheets As String
    
    Dim blCheck_S4A As Boolean
    
    With fpSpread1
        For intCtrl = 1 To .SheetCount - 2
            .sheet = intCtrl
            strSheets = strSheets & "," & .SheetName
            If .SheetVisible Then
                strSelectedSheets = strSelectedSheets & "," & .SheetName
            End If
        Next intCtrl
            
        If strSheets <> "" Then
            strSheets = Mid$(strSheets, 2)
        End If
        
        If strSelectedSheets <> "" Then
            strSelectedSheets = Mid$(strSelectedSheets, 2)
        End If
        
        strAddedSheet = strSelectedSheets
        strAddedSheet = frmAddSheet.SheetSelections(strSheets, strSelectedSheets)
        
        If strAddedSheet = strSelectedSheets Then
            Exit Sub
        End If
        
        strAddedSheet = "," & strAddedSheet & ","
        For intCtrl = 1 To .SheetCount
            .sheet = intCtrl
            If InStr(1, strAddedSheet, "," & .SheetName & ",") <> 0 Then
                If Not .SheetVisible Then
                    ResetDataAndForm intCtrl
                    .SheetVisible = True
                    SetAttribute TAX_Utilities_New.NodeValidity.childNodes(intCtrl - 1), "Active", "1"
                    TAX_Utilities_New.AdjustData(intCtrl - 1) = True
                    
                End If
            End If
        Next intCtrl
        
        ' Them phu luc tren to khai 01_GTGT
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
            blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S4A = True Then
                If Not objTaxBusiness Is Nothing Then
                     objTaxBusiness.update_01_4A
                     objTaxBusiness.reset_01_4A
                End If
            End If
        End If
        
        ' Them phu luc tren to khai 03_TNDN
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "03" Then
            If Not objTaxBusiness Is Nothing Then
                     objTaxBusiness.unLockCellPL (objTaxBusiness.strloaitk)
                     objTaxBusiness.tongHopPL05
            End If
        End If
        ' TO khai 04_NTNN
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "82" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "80" Then
            If Not objTaxBusiness Is Nothing Then
                objTaxBusiness.updateSomeCell
            End If
        End If
        ' TK 01 TD
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "83" Then
            If Not objTaxBusiness Is Nothing Then
                objTaxBusiness.LockCellBySheet
            End If
        End If
        
    End With
End Sub

Private Sub cmdKiemTra_Click()
Dim blFinish As Boolean
    Lbload.Visible = True
    If (Not objTaxBusiness Is Nothing) And blFinish = False Then
        objTaxBusiness.flgloadToKhai = flgloadToKhai
        ' objTaxBusiness.kiemTraDuLieuImport
        ' objTaxBusiness.finish
    End If
    CallFinish
    Lbload.Visible = False
    If CheckValidData = True Then
        DisplayMessage "0157", msOKOnly, miInformation
    Else
        DisplayMessage "0158", msOKOnly, miInformation
    End If
End Sub

Private Sub cmdLoadToKhai_Click()
    Dim checkLoadToKhai As Boolean
    Dim varMenuId As String
    Dim loaiToKhai As Variant, mstFile As Variant, mstUD As Variant, kyKeKhai As Variant
        
    checkLoadToKhai = False
    checkLoadToKhai = loadToKhai
    If checkLoadToKhai = False Then Exit Sub
    flgloadToKhai = True
    
    varMenuId = TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    With fpSpread2
        .sheet = 1
        If varMenuId = "17" Then
            ' Lay ky ke khai
            .GetText .ColLetterToNumber("H"), 4, kyKeKhai
            ' Lay MST
            .GetText .ColLetterToNumber("D"), 10, mstFile
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(mstFile))) = 10 Then
                mstFile = Left(LTrim(RTrim(mstFile)), 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(LTrim(Trim(mstFile))) > 10 Then
                mstFile = Left(LTrim(RTrim(mstFile)), 10) & Right(LTrim(RTrim(mstFile)), 3)
            End If
        ElseIf varMenuId = "42" Or varMenuId = "43" Or varMenuId = "59" Or varMenuId = "41" Or varMenuId = "76" Then
            ' Lay ky ke khai
            .GetText .ColLetterToNumber("H"), 4, kyKeKhai
            ' Lay MST
            .GetText .ColLetterToNumber("D"), 10, mstFile
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(mstFile))) = 10 Then
                mstFile = Left(LTrim(RTrim(mstFile)), 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(LTrim(Trim(mstFile))) > 10 Then
                mstFile = Left(LTrim(RTrim(mstFile)), 10) & Right(LTrim(RTrim(mstFile)), 3)
            End If
        Else
            ' Lay ky ke khai
            .GetText .ColLetterToNumber("H"), 6, kyKeKhai
            ' Lay MST
            .GetText .ColLetterToNumber("D"), 8, mstFile
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(mstFile))) = 10 Then
                mstFile = Left(LTrim(RTrim(mstFile)), 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(LTrim(Trim(mstFile))) > 10 Then
                mstFile = Left(LTrim(RTrim(mstFile)), 10) & Right(LTrim(RTrim(mstFile)), 3)
            End If
        End If
        ' Lay loai to khai
        If Trim(varMenuId) = "17" Or Trim(varMenuId) = "41" Or Trim(varMenuId) = "44" Or Trim(varMenuId) = "42" Or Trim(varMenuId) = "43" Or Trim(varMenuId) = "59" Or Trim(varMenuId) = "76" Then
            .GetText .ColLetterToNumber("O"), 1, loaiToKhai
        End If
    End With
    
    objTaxBusiness.flgloadToKhai = flgloadToKhai
    
    ' Kiem tra voi MST dang nhap vao UD xem co dong nhat ko?.
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
        .sheet = 1
        If varMenuId = "17" Or varMenuId = "59" Or varMenuId = "42" Or varMenuId = "43" Or varMenuId = "41" Or varMenuId = "76" Then
            ' Lay MST cua NNT
            .GetText .ColLetterToNumber("D"), 10, mstUD
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(mstUD))) = 10 Then
                mstUD = Left(LTrim(RTrim(mstUD)), 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(LTrim(Trim(mstUD))) > 10 Then
                mstUD = Left(LTrim(RTrim(mstUD)), 10) & Right(LTrim(RTrim(mstUD)), 3)
            End If
        Else
            ' Lay MST cua NNT
            .GetText .ColLetterToNumber("D"), 8, mstUD
            ' Neu la ma 10 so thi them chuoi 000 vao sau
            If Len(LTrim(RTrim(mstUD))) = 10 Then
                mstUD = Left(LTrim(RTrim(mstUD)), 10) & "000"
                ' Neu la ma 13 so thi giu nguyen cau truc
            ElseIf Len(LTrim(Trim(mstUD))) > 10 Then
                mstUD = Left(LTrim(RTrim(mstUD)), 10) & Right(LTrim(RTrim(mstUD)), 3)
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
    If Trim(mstUD) <> Trim(mstFile) Then
        DisplayMessage "0141", msOKOnly, miInformation
        Exit Sub
    End If
    
    If Val(kyKeKhai) < 2009 Then
        DisplayMessage "0142", msOKOnly, miInformation
        Exit Sub
    End If
    ' To khai 05/KK-TNCN
    If (varMenuId = "17" And UCase(loaiToKhai) <> "05/KK-TNCN") Then
        DisplayMessage "0143", msOKOnly, miInformation
        Exit Sub
    End If
    ' To khai 09/KK-TNCN
    If (varMenuId = "41" And UCase(loaiToKhai) <> "09/KK-TNCN") Then
        DisplayMessage "0144", msOKOnly, miInformation
        Exit Sub
    End If
    ' To khai 06/KK-TNCN
    If (varMenuId = "44" And UCase(loaiToKhai) <> "06/KK-TNCN") Then
        DisplayMessage "0145", msOKOnly, miInformation
        Exit Sub
    End If
    ' To khai 06/KK-TNCN10
    If (varMenuId = "59" And UCase(loaiToKhai) <> "06/KK-TNCN") Then
        DisplayMessage "0145", msOKOnly, miInformation
        Exit Sub
    End If

    
    ' To khai 02/KK-BH
    If (varMenuId = "42" And UCase(loaiToKhai) <> "02/KK-BH") Then
        DisplayMessage "0146", msOKOnly, miInformation
        Exit Sub
    End If
    ' To khai 02/KK-XS
    If (varMenuId = "43" And UCase(loaiToKhai) <> "02/KK-XS") Then
        DisplayMessage "0147", msOKOnly, miInformation
        Exit Sub
    End If
    ' Lay du lieu cua to khai 05/KK
    If varMenuId = "17" Then
        convertData05KK
    End If
    ' Lay du lieu cua to khai 02/BH
    If varMenuId = "42" Then
        convertData02BH
    End If
    ' Lay du lieu cua to khai 02/XS
    If varMenuId = "43" Then
        convertData02XS
    End If
    ' Lay du lieu cua to khai 09/KK
    If varMenuId = "41" Then
        convertData09KK
    End If
    ' Lay du lieu cua to khai 06/KK
    If varMenuId = "44" Then
        convertData06KK
    End If
    ' Lay du lieu cua to khai 06/KK 10
    If varMenuId = "59" Then
        convertData06KK10
    End If
    ' lay du lieu cua to 08B_TNCN
    
    If varMenuId = "76" Then
        convertData08B
    End If
    
End Sub

Private Sub convertData05KK()
    Dim varTemp As Variant
    'Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    Dim varTemp3 As Variant
    Dim varTemp4 As Variant
    Dim idx As Integer
    
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            'fpSpread2.GetText .ColLetterToNumber("G"), 4, varTemp1
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp2
            fpSpread2.GetText .ColLetterToNumber("D"), 10, varTemp3
            fpSpread1.GetText .ColLetterToNumber("D"), 10, varTemp4
            
            If (Trim(varTemp) = "" Or Trim(varTemp) = vbNullString) And (Trim(varTemp2) = "" Or Trim(varTemp2) = vbNullString) Then
                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
'            If (UCase(Trim(varTemp)) = "[X]") And (UCase(Trim(varTemp1)) = "X") Then
'                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
'                Exit Sub
'            End If
            If (UCase(Trim(varTemp)) = "[X]") And (Trim(varTemp2) <> "" Or Trim(varTemp2) <> vbNullString) Then
                DisplayMessage "0172", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            If (Trim(varTemp3) <> Trim(varTemp4)) Then
                DisplayMessage "0173", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            If UCase(varTemp) = "[X]" Then
                .Col = .ColLetterToNumber("C")
                .Row = 67
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 67
                .Text = "0"
                UpdateCell .Col, .Row, .value
            End If
'            fpSpread2.GetText .ColLetterToNumber("G"), 4, varTemp
'            If UCase(varTemp) = "X" Then
'                .Col = .ColLetterToNumber("F")
'                .Row = 41
'                .Text = "1"
'                UpdateCell .Col, .Row, .value
'            Else
'                .Col = .ColLetterToNumber("F")
'                .Row = 41
'                .Text = "0"
'                UpdateCell .Col, .Row, .value
'            End If
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp
            
            .Col = .ColLetterToNumber("I")
            .Row = 67
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            .Col = .ColLetterToNumber("I")
            
            ' set cac gia tri cua cac chi tieu tu 21 den 42
            ' Nghia vu khau tru thue
            For idx = 36 To 57
                ' Chi tieu idx - 15
                .Row = idx
                fpSpread2.GetText .Col, .Row, varTemp
                .Text = Round(varTemp, 0)
                UpdateCell .Col, .Row, .value
            Next idx
            
            ' Nghia vu quyet toan thay
            For idx = 61 To 65
                ' Chi tieu idx - 18
                .Row = idx
                fpSpread2.GetText .Col, .Row, varTemp
                .Text = Round(varTemp, 0)
                UpdateCell .Col, .Row, .value
            Next idx
           
            
            .Col = .ColLetterToNumber("M")
            
            ' Nguoi Ky
            .Row = 69
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
                        
            ' Ngay Ky
            .Row = 71
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            .Col = .ColLetterToNumber("D")
            ' Nhan vien dai ly thue
            .Row = 69
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            ' chung chi so
            .Row = 71
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            .sheet = 2
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai5A
            'dhdang edit
            'date 08-06-2010
            'Turning Load to khai
            'CallFinish

            .sheet = 3
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai5A
            'dhdang edit
            'date 08-06-2010
            'Turning Load to khai
            'CallFinish

            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub convertData02BH()
    Dim varTemp As Variant
    Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    Dim varTemp3 As Variant
    Dim varTemp4 As Variant
    
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
'            fpSpread2.GetText .ColLetterToNumber("G"), 4, varTemp1
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp2
            fpSpread2.GetText .ColLetterToNumber("D"), 10, varTemp3
            fpSpread1.GetText .ColLetterToNumber("D"), 10, varTemp4
            
            If (Trim(varTemp) = "" Or Trim(varTemp) = vbNullString) And (Trim(varTemp2) = "" Or Trim(varTemp2) = vbNullString) Then
                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
'            If (UCase(Trim(varTemp)) = "X") Then
'                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
'                Exit Sub
'            End If
            If (UCase(Trim(varTemp)) = "X") And (Trim(varTemp2) <> "" Or Trim(varTemp2) <> vbNullString) Then
                DisplayMessage "0172", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            If (Trim(varTemp3) <> Trim(varTemp4)) Then
                DisplayMessage "0173", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            If UCase(varTemp) = "[X]" Then
                .Col = .ColLetterToNumber("C")
                .Row = 48
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 48
                .Text = ""
                UpdateCell .Col, .Row, .value
            End If
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp
            
            .Col = .ColLetterToNumber("I")
            .Row = 48
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
                        
            .Col = .ColLetterToNumber("I")
            
            ' Chi tieu 21
            .Row = 39
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 22
            .Row = 40
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 23
            .Row = 41
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 24
            .Row = 42
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 25
            .Row = 43
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
                        
                        
            .sheet = 2
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai

            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub convertData02XS()
    Dim varTemp As Variant
    Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    Dim varTemp3 As Variant
    Dim varTemp4 As Variant
    
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
'            fpSpread2.GetText .ColLetterToNumber("G"), 4, varTemp1
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp2
            fpSpread2.GetText .ColLetterToNumber("D"), 10, varTemp3
            fpSpread1.GetText .ColLetterToNumber("D"), 10, varTemp4
            
            If (Trim(varTemp) = "" Or Trim(varTemp) = vbNullString) And (Trim(varTemp2) = "" Or Trim(varTemp2) = vbNullString) Then
                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
'            If (UCase(Trim(varTemp)) = "X") Then
'                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
'                Exit Sub
'            End If
            If (UCase(Trim(varTemp)) = "X") And (Trim(varTemp2) <> "" Or Trim(varTemp2) <> vbNullString) Then
                DisplayMessage "0172", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            If (Trim(varTemp3) <> Trim(varTemp4)) Then
                DisplayMessage "0173", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            If UCase(varTemp) = "[X]" Then
                .Col = .ColLetterToNumber("C")
                .Row = 51
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 51
                .Text = ""
                UpdateCell .Col, .Row, .value
            End If
            fpSpread2.GetText .ColLetterToNumber("I"), 51, varTemp
            
            .Col = .ColLetterToNumber("I")
            .Row = 51
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            .Col = .ColLetterToNumber("I")
            
            ' Chi tieu 21
            .Row = 40
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 22
            .Row = 41
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 23
            .Row = 42
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 24
            .Row = 43
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
            ' Chi tieu 25
            .Row = 44
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = Round(varTemp, 0)
            UpdateCell .Col, .Row, .value
                        
            .sheet = 2
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai

            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub convertData06KK()
    Dim varTemp As Variant
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            fpSpread2.GetText .ColLetterToNumber("E"), 4, varTemp
            If UCase(varTemp) = "X" Then
                .Col = .ColLetterToNumber("C")
                .Row = 45
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 45
                .Text = "0"
                UpdateCell .Col, .Row, .value
            End If
            fpSpread2.GetText .ColLetterToNumber("G"), 4, varTemp
            If UCase(varTemp) = "X" Then
                .Col = .ColLetterToNumber("F")
                .Row = 45
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("F")
                .Row = 45
                .Text = "0"
                UpdateCell .Col, .Row, .value
            End If
            fpSpread2.GetText .ColLetterToNumber("I"), 4, varTemp
            
            .Col = .ColLetterToNumber("I")
            .Row = 45
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            .Col = .ColLetterToNumber("I")
            
            ' Chi tieu 17
            .Row = 37
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 18
            .Row = 38
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 19
            .Row = 39
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            ' Chi tieu 20
            .Row = 41
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 21
            .Row = 42
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 22
            .Row = 43
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
                                    
            .sheet = 2
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai

            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub convertData06KK10()
    Dim varTemp As Variant
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            If UCase(varTemp) = "X" Then
                .Col = .ColLetterToNumber("C")
                .Row = 61
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 61
                .Text = ""
                UpdateCell .Col, .Row, .value
            End If
            
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp
            
            .Col = .ColLetterToNumber("I")
            .Row = 61
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            .Col = .ColLetterToNumber("I")
            
            ' Chi tieu 21
            .Row = 41
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 22
            .Row = 42
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 25
            .Row = 47
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            ' Chi tieu 27
            .Row = 50
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 29
            .Row = 53
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            ' Chi tieu 30
            .Row = 54
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
                                    
            ' Chi tieu 31
            .Row = 55
            fpSpread2.GetText .Col, .Row, varTemp
            .Text = varTemp
            UpdateCell .Col, .Row, .value
                                    
            .sheet = 2
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai

            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub convertData09KK()
    Dim varTemp As Variant
    Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    Dim varTemp3 As Variant
    Dim varTemp4 As Variant
    Dim i As Integer
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
'            fpSpread2.GetText .ColLetterToNumber("G"), 4, varTemp1
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp2
            fpSpread2.GetText .ColLetterToNumber("D"), 10, varTemp3
            fpSpread1.GetText .ColLetterToNumber("D"), 10, varTemp4
            
            If (Trim(varTemp) = "" Or Trim(varTemp) = vbNullString) And (Trim(varTemp2) = "" Or Trim(varTemp2) = vbNullString) Then
                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
'            If (UCase(Trim(varTemp)) = "X") Then
'                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
'                Exit Sub
'            End If
            If (UCase(Trim(varTemp)) = "X") And (Trim(varTemp2) <> "" Or Trim(varTemp2) <> vbNullString) Then
                DisplayMessage "0172", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            If (Trim(varTemp3) <> Trim(varTemp4)) Then
                DisplayMessage "0173", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            If UCase(varTemp) = "[X]" Then
                .Col = .ColLetterToNumber("C")
                .Row = 70
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 70
                .Text = "0"
                UpdateCell .Col, .Row, .value
            End If
                        
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp
            .Col = .ColLetterToNumber("I")
            .Row = 70
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            
            ' Lay gia tri tu thang
            fpSpread2.GetText .ColLetterToNumber("L"), 4, varTemp
            .Col = .ColLetterToNumber("L")
            .Row = 4
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            ' Lay gia tri den thang
            fpSpread2.GetText .ColLetterToNumber("N"), 4, varTemp1
            .Col = .ColLetterToNumber("N")
            .Row = 4
            .Text = varTemp1
            UpdateCell .Col, .Row, .value
            
            ' Tinh tong so thang giua tu thang den thang
            .Row = 4
            .Col = .ColLetterToNumber("P")
            .Text = DateDiff("M", format(Trim(varTemp), "mm/yyyy"), format(Trim(varTemp1), "mm/yyyy")) + 1
            
            
            ' Lay gia tri so tai khoan ngan hang
'            fpSpread2.GetText .ColLetterToNumber("D"), 20, varTemp
'            .Col = .ColLetterToNumber("D")
'            .Row = 20
'            .Text = varTemp
'            UpdateCell .Col, .Row, .value
            
            ' Lay gia tri ten ngan hang mo tai khoan
'            fpSpread2.GetText .ColLetterToNumber("M"), 20, varTemp
'            .Col = .ColLetterToNumber("M")
'            .Row = 20
'            .Text = varTemp
'            UpdateCell .Col, .Row, .value
            
            .Col = .ColLetterToNumber("I")
            
            ' Cac chi tieu trong to khai 09
            For i = 42 To 61
                .Row = i
                fpSpread2.GetText .Col, .Row, varTemp
                .Text = varTemp
                UpdateCell .Col, .Row, .value
            Next
            
            .sheet = 2
            fpSpread2.sheet = .sheet
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            .Col = .ColLetterToNumber("I")
            
            ' Cac chi tieu trong phu luc 09A
            For i = 16 To 24
                .Row = i
                fpSpread2.GetText .Col, .Row, varTemp
                .Text = varTemp
                UpdateCell .Col, .Row, .value
            Next
                                   
            .sheet = 3
            fpSpread2.sheet = .sheet
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            .Col = .ColLetterToNumber("I")
            
            ' Cac chi tieu trong phu luc 09B
            For i = 16 To 31
                .Row = i
                fpSpread2.GetText .Col, .Row, varTemp
                .Text = varTemp
                UpdateCell .Col, .Row, .value
            Next
            
            ' Cac chi tieu trong phu luc 09C
            .sheet = 4
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            moveDataToKhai
            
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
Private Sub convertData08B()
    Dim varTemp As Variant
    Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    Dim varTemp3 As Variant
    Dim varTemp4 As Variant
    Dim i As Integer
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
            
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp2
            fpSpread2.GetText .ColLetterToNumber("D"), 10, varTemp3
            fpSpread1.GetText .ColLetterToNumber("D"), 10, varTemp4
            
            If (Trim(varTemp) = "" Or Trim(varTemp) = vbNullString) And (Trim(varTemp2) = "" Or Trim(varTemp2) = vbNullString) Then
                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
'            If (UCase(Trim(varTemp)) = "X") Then
'                DisplayMessage "0171", msOKOnly, miInformation, "Tai bang ke"
'                Exit Sub
'            End If
            If (UCase(Trim(varTemp)) = "X") And (Trim(varTemp2) <> "" Or Trim(varTemp2) <> vbNullString) Then
                DisplayMessage "0172", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            If (Trim(varTemp3) <> Trim(varTemp4)) Then
                DisplayMessage "0173", msOKOnly, miInformation, "Tai bang ke"
                Exit Sub
            End If
            
            fpSpread2.GetText .ColLetterToNumber("E"), 6, varTemp
            If UCase(varTemp) = "[X]" Then
                .Col = .ColLetterToNumber("C")
                .Row = 38
                .Text = "1"
                UpdateCell .Col, .Row, .value
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 38
                .Text = "0"
                UpdateCell .Col, .Row, .value
            End If
                        
            fpSpread2.GetText .ColLetterToNumber("I"), 6, varTemp
            .Col = .ColLetterToNumber("I")
            .Row = 38
            .Text = varTemp
            UpdateCell .Col, .Row, .value
            
            ' Cac chi tieu trong to khai 08
            .Col = .ColLetterToNumber("I")
            For i = 40 To 51
                .Row = i
                fpSpread2.GetText .Col, .Row, varTemp
                .Text = varTemp
                UpdateCell .Col, .Row, .value
            Next
            
            ' Cac chi tieu trong phan II
            moveDataToKhai08B
             
            
            .sheet = 1
            mCurrentSheet = .sheet
            .ActiveSheet = .sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
''' cmdSave_Click description
''' Checking business error but user can save it anyway
''' No parameter
''' cmdSave_Click description
''' Checking business error but user can save it anyway
''' No parameter
Private Sub cmdSave_Click()
    On Error GoTo ErrorHandle
    Dim blnValid As Boolean
    
    'Debug.Print "Bat dau ghi" & Time
    Lbload.Visible = True
    
    flgloadToKhai = False
    Dim varMenuId As String
    varMenuId = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")

    If strKHBS = "TKBS" And (varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "05" Or varMenuId = "06" Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "89" Or varMenuId = "71" _
    Or varMenuId = "72" Or varMenuId = "77" Or varMenuId = "03" Or varMenuId = "73" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" Or varMenuId = "85") Then
        TonghopKHBS
    End If
  ' Save KHBS
            If strKHBS = "frmKHBS_BS" Then
                Call objTaxBusiness.UpdateChangeKHBS
                saveKHBS
                Exit Sub
            End If

    ' Doi voi cac to quyet toan TNCN thi dat co flgloadToKhai = false
    ' Muc dich la trong truong hop load bang ke thi ko tong hop lai du lieu
    ' Khi Ghi, In, Ket xuat thi dat lai trang thai
    If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        objTaxBusiness.flgloadToKhai = flgloadToKhai
    End If
  

  'CallFinish
             CallFinish
            
            
            Dim intCtrl As Integer
            Dim strArrActive() As String
            
            'Backup node validity
            For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
                ReDim Preserve strArrActive(intCtrl)
                strArrActive(intCtrl) = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(intCtrl), "Active")
            Next intCtrl
            If Not objTaxBusiness Is Nothing Then
                'For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
                    Call objTaxBusiness.SetActiveSheet '(TAX_Utilities_New.NodeValidity.childNodes(intCtrl))
                'Next intCtrl
            End If
            
            
            blnValid = CheckValidData
            If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "36") Then
                If objTaxBusiness.iflag = True Then
                    DisplayMessage "0225", msOKOnly, miInformation
                End If
            End If
             
             If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73") Then
                If objTaxBusiness.iflag = True Then
                    DisplayMessage "0240", msOKOnly, miCriticalError
                End If
            End If
            
            'Restore active properties of node validity
            For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
                SetAttribute TAX_Utilities_New.NodeValidity.childNodes(intCtrl), "Active", strArrActive(intCtrl)
            Next intCtrl
            '****************************
            Lbload.Visible = False
                
    If Not blnValid And (checkSoCT = 1 Or checkSoCT = 2 Or checkSoCT = 3 Or checkSoCT = 4) Then
        If DisplayMessage("0184", msYesNo, miQuestion) = mrYes Then
            If UpdateData Then DisplayMessage "0002", msOKOnly, miInformation
        End If

    ElseIf Not blnValid And checkSoCT = 0 Then

        If DisplayMessage("0015", msYesNo, miQuestion) = mrYes Then
            If UpdateData Then DisplayMessage "0002", msOKOnly, miInformation
        End If

    Else

        If UpdateData Then DisplayMessage "0002", msOKOnly, miInformation
    End If

    ' Set lai co isNewDataBS sau khi bam nut ghi
    If strKHBS = "TKBS" And (varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "05" Or varMenuId = "06" _
    Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "89" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "77" Or varMenuId = "03" Or varMenuId = "73" Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" _
    Or varMenuId = "85") Then
        isNewdataBS = False
    End If
            
            fpSpread1.sheet = fpSpread1.ActiveSheet
            SetStatus fpSpread1.ActiveCol, fpSpread1.ActiveRow
            fpSpread1.SetFocus
    
   'Debug.Print "Ket thuc ghi ghi" & Time
    
    Exit Sub
    
ErrorHandle:
    '****************************
    ' added
    'Restore active properties of node validity
    For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(intCtrl), "Active", strArrActive(intCtrl)
    Next intCtrl
    '****************************
    SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
End Sub


'this sub is called to execute objTaxBusiness.Finish
'and solve the hotkey problem
Private Sub CallFinish(Optional blFinish As Boolean)
    '*****************************
    ' Xu ly cho truong hop bam phim nong.
    DoEvents
    '*****************************
    
    On Error GoTo ErrorHandle
        
    Dim iSheet As Integer, iActiveSheet As Integer
    Dim lActiveCol As Long, lActiveRow As Long
    Dim lCol As Long, lRow As Long
    Dim i As Long, arrLActiveCol() As String
    Dim arrLActiveRow() As Long
    Dim arrStrPositions As Variant, arrStrPosition() As String
    
    lblStatus.Visible = False
    With fpSpread1
     If blFinish = False Then
        For i = 1 To .SheetCount - 1
            .sheet = i
            If .SheetVisible Then
                If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "17" Then
                    delNullRowOn05 i - 1
                ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "59" Then
                    delNullRowOn06 i - 1
                ' dntai sua phan del rownull 16022012
                ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Then
                    If i = 2 Or i = 3 Then
                        delNullRowOn01 i - 1
                    Else
                        delNullRow i - 1
                    End If
                Else
                    delNullRow i - 1
                End If
            End If
        Next
    End If
    
        .Visible = False
        .ReDraw = False
        .EditMode = False
        iActiveSheet = .ActiveSheet
        lActiveCol = .ActiveCol
        lActiveRow = .ActiveRow
        
        ReDim arrLActiveCol(.SheetCount)
        ReDim arrLActiveRow(.SheetCount)
        
'***************************************
  'Xoa cac canh bao tren form
        .EventEnabled(EventAllEvents) = False
        arrStrPositions = arrErrCells.Keys
        For i = 1 To arrErrCells.count
            arrStrPosition = Split(CStr(arrStrPositions(i - 1)), "_")
            .sheet = CLng(arrStrPosition(0))
            .Col = .ColLetterToNumber(arrStrPosition(1))
            .Row = CLng(arrStrPosition(2))
            .CellNote = ""
            .BackColor = arrErrCells.Item(arrStrPositions(i - 1))
        Next
        arrErrCells.RemoveAll
        .EventEnabled(EventAllEvents) = True
'***************************************
        
        For i = 1 To .SheetCount
            .ActiveSheet = i
            .sheet = .ActiveSheet
            arrLActiveCol(i) = .ActiveCol
            arrLActiveRow(i) = .ActiveRow
            
            .Row = 1
            .Col = 1
            .Lock = False
            .SetActiveCell 1, 1
            .EditMode = True
'            .EditMode = False
'            .SetActiveCell arrLActiveCol(i), arrLActiveRow(i)
'            .Lock = True
        Next
        
        'Reset all of error satatus on form
        ResetErrorCells
        
        If (Not objTaxBusiness Is Nothing) And blFinish = False Then
            objTaxBusiness.finish
        End If
        
        For i = 1 To .SheetCount
            .ActiveSheet = i
            .sheet = .ActiveSheet
            .Row = 1
            .Col = 1
            .Lock = True
            .EditMode = False
            .SetActiveCell arrLActiveCol(i), arrLActiveRow(i)
        Next

        .ActiveSheet = iActiveSheet
        .sheet = iActiveSheet
        .Col = lActiveCol
        .Row = lActiveRow
        .EditMode = True
        .SetActiveCell lActiveCol, lActiveRow
        .ReDraw = True
        .Visible = True
        
    End With
    
    lblStatus.Visible = True
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "CallFinish", Err.Number, Err.Description
    lblStatus.Visible = True
End Sub

''' cmdPrint_Click description
''' Show print report form
''' No parameter
Private Sub cmdPrint_Click()
    On Error GoTo ErrorHandle
    Dim varTemp As Variant
    
    flgloadToKhai = False
    
    ' Trong truong hop in bia thi ko check gi ca, in luon
    If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "52" Then
        frmReports.Show 1
        Exit Sub
    End If
    
    '****************************
    ' added
    'Modify date: 13/12/2005
    ' Neu la cac mau in tong hop tu to quyet toan 05TNCN->09TNCN va cac chung tu cua TNCN thi hien thi phan In luon!
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "45" Then
        Call objTaxBusiness.prepareDataPrinter
        frmReports.Show 1
        Exit Sub
    End If
    
    ' Doi voi cac to quyet toan TNCN thi dat co flgloadToKhai = false
    ' Muc dich la trong truong hop load bang ke thi ko tong hop lai du lieu
    ' Khi Ghi, In, Ket xuat thi dat lai trang thai
    If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
        objTaxBusiness.flgloadToKhai = flgloadToKhai
    End If
    
    Dim intCtrl As Integer
    Dim strArrActive() As String
    ' Print KHBS
            If strKHBS = "frmKHBS_BS" Then
                Call objTaxBusiness.UpdateChangeKHBS
                
                ' Doi voi to khai khau tru, trong truong hop dieu chinh lam giam so thue phai nop
                ' Tuc la chi tieu [41] > 0 hoac chi tieu [43] > 0 thi:
                ' 1. Khong cho in ra to khai bo sung
                ' 2. Thong bao dieu chinh thue vao ky ke khai hien tai va dieu chinh vao phu luc 03_GTGT
                
'                If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Then
'                    fpSpread1.sheet = 1
'                    fpSpread1.Col = fpSpread1.ColLetterToNumber("L")
'                    fpSpread1.Row = 31
'                    If fpSpread1.value > 0 Then
'                        DisplayMessage "0138", msOKOnly, miInformation
'                        fpSpread1.sheet = 1
'                        fpSpread1.SetFocus
'                        fpSpread1.SetActiveCell fpSpread1.ColLetterToNumber("L"), 31
'                        Exit Sub
'                    End If
'                    fpSpread1.sheet = 1
'                    fpSpread1.Col = fpSpread1.ColLetterToNumber("L")
'                    fpSpread1.Row = 33
'                    If fpSpread1.value > 0 Then
'                        DisplayMessage "0138", msOKOnly, miInformation
'                        fpSpread1.sheet = 1
'                        fpSpread1.SetFocus
'                        fpSpread1.SetActiveCell fpSpread1.ColLetterToNumber("L"), 33
'                        Exit Sub
'                    End If
'                    frmReports.Show 1
'                    Exit Sub
'                Else ' Cac truong hop con lai sau khi Update data KHBS la cho in luon.
'                    frmReports.Show 1
'                    Exit Sub
'                End If
                frmReports.Show 1
                Exit Sub
                
            End If
    
    CallFinish
    
    If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73") Then
                If objTaxBusiness.iflag = True Then
                    DisplayMessage "0240", msOKOnly, miCriticalError
                    Exit Sub
                End If
            End If
    'Backup node validity
    For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
        ReDim Preserve strArrActive(intCtrl)
        strArrActive(intCtrl) = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(intCtrl), "Active")
    Next intCtrl
    If Not objTaxBusiness Is Nothing Then
        'For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
            Call objTaxBusiness.SetActiveSheet '(TAX_Utilities_New.NodeValidity.childNodes(intCtrl))
        'Next intCtrl
    End If
    '****************************
        
        If CheckValidData = True Then
            ' Trong truong hop in dieu chinh thi cac to khai quyet toan TNCN hien thi o check "In thong tin dieu chinh"
            Dim varMenuId As String
            varMenuId = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
            If Trim(varMenuId) = "17" Or Trim(varMenuId) = "42" Or Trim(varMenuId) = "43" Or Trim(varMenuId) = "59" Then
                Dim countInBoSung As Integer
                flgPrintBoSung = False
                countInBoSung = 1
                fpSpread1.sheet = 1
                fpSpread1.Col = fpSpread1.ColLetterToNumber("I")
                fpSpread1.Row = 6
                fpSpread1.GetText fpSpread1.Col, fpSpread1.Row, varTemp
                If varTemp = "x" Or Trim(varTemp) <> "" Then
                    With fpSpread1
                        If Trim(varMenuId) = "17" Then
                            .sheet = 2
                            .Row = 22
                            Do
                                .Col = .ColLetterToNumber("C")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    flgPrintBoSung = True
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                            .Row = 22
                            ReDim listInBoSung5A(countInBoSung - 1) As String
                            countInBoSung = 1
                            Do
                                .Col = .ColLetterToNumber("C")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    .Col = .ColLetterToNumber("B")
                                    listInBoSung5A(countInBoSung - 1) = .Text
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                            
                            .sheet = 3
                            countInBoSung = 1
                            .Row = 22
                            Do
                                .Col = .ColLetterToNumber("Y")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    flgPrintBoSung = True
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                            .Row = 22
                            ReDim listInBoSung5B(countInBoSung - 1) As String
                            countInBoSung = 1
                            Do
                                .Col = .ColLetterToNumber("Y")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    .Col = .ColLetterToNumber("B")
                                    listInBoSung5B(countInBoSung - 1) = .Text
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                        End If
                        
                        If flgPrintBoSung = True Then
                            frmReports.chkDieuChinh.Visible = True
                            frmReports.chkDieuChinh.value = 1
                        Else
                            frmReports.chkDieuChinh.Visible = True
                            frmReports.chkDieuChinh.value = 0
                        End If
                        
                        ' To quyet toan 02BH, 02SX
                        countInBoSung = 1
                        If Trim(varMenuId) = "42" Or Trim(varMenuId) = "43" Then
                            .sheet = 2
                            .Row = 22
                            Do
                                .Col = .ColLetterToNumber("C")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    flgPrintBoSung = True
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                            .Row = 22
                            ReDim listInBoSung2A(countInBoSung - 1) As String
                            countInBoSung = 1
                            Do
                                .Col = .ColLetterToNumber("C")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    .Col = .ColLetterToNumber("B")
                                    listInBoSung2A(countInBoSung - 1) = .Text
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                        End If
                        
                        ' to khai 06KK-TNCN
                         If Trim(varMenuId) = "59" Then
                            .sheet = 2
                            .Row = 22
                            Do
                                .Col = .ColLetterToNumber("C")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    flgPrintBoSung = True
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                            .Row = 22
                            ReDim listInBoSung6B(countInBoSung - 1) As String
                            countInBoSung = 1
                            Do
                                .Col = .ColLetterToNumber("C")
                                .GetText .Col, .Row, varTemp
                                If varTemp = "1" Or varTemp = "x" Then
                                    .Col = .ColLetterToNumber("B")
                                    listInBoSung6B(countInBoSung - 1) = .Text
                                    countInBoSung = countInBoSung + 1
                                End If
                                .Row = .Row + 1
                                .Col = .ColLetterToNumber("B")
                            Loop Until .Text = "aa"
                       End If
                        
                        If flgPrintBoSung = True Then
                            frmReports.chkDieuChinh.Visible = True
                            frmReports.chkDieuChinh.value = 1
                        Else
                            frmReports.chkDieuChinh.Visible = True
                            frmReports.chkDieuChinh.value = 0
                        End If
                    End With
                End If
                
            End If
            ' End truong hop in dieu chinh thi cac to khai quyet toan TNCN
            
            frmReports.Show 1
            
        Else
            DisplayMessage "0016", msOKOnly, miInformation
        End If
    
    '****************************
    ' added
    'Modify date: 13/12/2005
    'Restore active properties of node validity
    For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(intCtrl), "Active", strArrActive(intCtrl)
    Next intCtrl
    '****************************
    Exit Sub
    
ErrorHandle:
    'Restore active properties of node validity
    For intCtrl = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(intCtrl), "Active", strArrActive(intCtrl)
    Next intCtrl
    '****************************
    SaveErrorLog Me.Name, "cmdPrint_Click", Err.Number, Err.Description
End Sub


Private Sub Command1_Click()
    Dim strTemp As String
    Dim strOldValue As String
    Dim strDieuChinhTangGiam() As String
    Dim arrDieuChinhGiam() As String
    Dim arrDieuChinhTang() As String
    Dim arrDieuChinh4043() As String
    Dim arrValue() As String ' Luu cac cell cua mot row
    Dim numRowI, numRowII, numRowIII, j As Integer
    Dim tempCurrSheet As Integer
    
    Dim flagTang, flagGiam, flag4043 As Boolean
    
    Dim strTongOld, strTongCurr As String ' Luu gia tri tong dieu chinh
    
    Dim countDel As Long
    numRowI = 0
    numRowII = 0
    numRowIII = 0
    ' set lai cong thuc
    'set lai cong thuc cua cac cell NNC va PNC
    Dim lCol_temp As Long
    Dim lRow_temp As Long
    Dim temp As Long
    
    Dim strFormula As String
    Dim vSoTien As Variant
    
    Dim xmlNodeCell_temp As MSXML.IXMLDOMNode
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11)
            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
            fpSpread1.sheet = fpSpread1.SheetCount - 1
            fpSpread1.Col = lCol_temp
            fpSpread1.Row = lRow_temp
            
            fpSpread1.Formula = "BD5"
'            fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
'                            (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11), "Value")
                            
            
            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10)
            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
            fpSpread1.sheet = fpSpread1.SheetCount - 1
            fpSpread1.Col = lCol_temp
            fpSpread1.Row = lRow_temp
            temp = lRow_temp - 18
            ' sua ct tinh
            fpSpread1.GetText fpSpread1.ColLetterToNumber("BH"), 15 + temp, vSoTien
            strFormula = getFormulaTienPNC(temp, CDbl(vSoTien), "BH" & 15 + temp)
            
            'fpSpread1.Formula = "IF((BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100)>0,ROUND(BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100,0),0)"
            fpSpread1.Formula = strFormula
            ' end
'            fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
'                            (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10), "Value")
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "02" Then
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "72" Then
            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
            fpSpread1.sheet = fpSpread1.SheetCount - 1
            fpSpread1.Col = lCol_temp
            fpSpread1.Row = lRow_temp
            fpSpread1.Formula = "BD5"
            
            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
            fpSpread1.sheet = fpSpread1.SheetCount - 1
            fpSpread1.Col = lCol_temp
            fpSpread1.Row = lRow_temp
            temp = lRow_temp - 18
            fpSpread1.Formula = ""
            fpSpread1.value = "0"
        Else
            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
            fpSpread1.sheet = fpSpread1.SheetCount - 1
            fpSpread1.Col = lCol_temp
            fpSpread1.Row = lRow_temp
            fpSpread1.Formula = "BD5"
'            fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
'                            (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7), "Value")
            
            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
            fpSpread1.sheet = fpSpread1.SheetCount - 1
            fpSpread1.Col = lCol_temp
            fpSpread1.Row = lRow_temp
            temp = lRow_temp - 18
            
            ' sua ct tinh
            fpSpread1.GetText fpSpread1.ColLetterToNumber("BH"), 15 + temp, vSoTien
            strFormula = getFormulaTienPNC(temp, CDbl(vSoTien), "BH" & 15 + temp)
            
            'fpSpread1.Formula = "IF((BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100)>0,ROUND(BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100,0),0)"
            fpSpread1.Formula = strFormula
            ' end
'            fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
'                            (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6), "Value")
    End If

    ' End set
    
    
    If Trim(GetAttribute(TAX_Utilities_New.NodeValidity, "Class")) <> vbNullString Then
        'Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_New.NodeValidity, "Class"))
        ' Neu chua co object moi tao lai
        If objTaxBusiness Is Nothing Then
            Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_New.NodeValidity, "Class"))
        End If
        
        Set objTaxBusiness.fps = fpSpread1
        strOldValue = objTaxBusiness.getValueTK(strDataFileBS)
        strTemp = objTaxBusiness.getDieuChinhGiam(strOldValue)
        
        'Lay ve gia tri tong
        strTongOld = objTaxBusiness.getValueCTDC(strDataFileBS)
        strTongCurr = objTaxBusiness.getChiTieuTongDC(CStr(strTongOld))
        'end
        
        strDieuChinhTangGiam = Split(strTemp, "###")
        If strDieuChinhTangGiam(0) <> "" Then
            arrDieuChinhGiam = Split(strDieuChinhTangGiam(0), "~")
            numRowII = UBound(arrDieuChinhGiam)
            flagGiam = True
        End If
        If strDieuChinhTangGiam(1) <> "" Then
            arrDieuChinhTang = Split(strDieuChinhTangGiam(1), "~")
            numRowI = UBound(arrDieuChinhTang)
            flagTang = True
        End If
        If GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01GTGT" Then
                If strDieuChinhTangGiam(2) <> "" Then
                    arrDieuChinh4043 = Split(strDieuChinhTangGiam(2), "~")
                    numRowIII = UBound(arrDieuChinh4043)
                    flag4043 = True
                End If
                ' A. Dieu chinh so thue CT 40 43
                fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
                fpSpread1.EventEnabled(EventAllEvents) = False
                tempCurrSheet = mCurrentSheet
                mCurrentSheet = fpSpread1.SheetCount - 1
                fpSpread1.sheet = mCurrentSheet
                ' them so dong dieu chinh thay doi vao
                ' set cac gia tri cua cot
                If flag4043 = True Then
                    For j = 0 To numRowIII
                        
                        arrValue = Split(arrDieuChinh4043(j), "_")
                        If arrValue(4) <> 0 Then
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BJ"), 5 + j, Round(Val(arrValue(2)), 0)
                            UpdateCell fpSpread1.ColLetterToNumber("BJ"), 5 + j, Round(Val(arrValue(2)), 0)
                            'UpdateCell fpSpread1.ColLetterToNumber("BF"), 15 + j, arrValue(2)
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BK"), 5 + j, Round(Val(arrValue(3)), 0)
                            UpdateCell fpSpread1.ColLetterToNumber("BK"), 5 + j, Round(Val(arrValue(3)), 0)
                            'UpdateCell fpSpread1.ColLetterToNumber("BG"), 15 + j, arrValue(3)
                        Else
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BJ"), 5 + j, "0"
                            UpdateCell fpSpread1.ColLetterToNumber("BJ"), 5 + j, "0"
                            'UpdateCell fpSpread1.ColLetterToNumber("BF"), 15 + j, arrValue(2)
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BK"), 5 + j, "0"
                            UpdateCell fpSpread1.ColLetterToNumber("BK"), 5 + j, "0"
                            'UpdateCell fpSpread1.ColLetterToNumber("BG"), 15 + j, arrValue(3)
                            
                        End If
                        fpSpread1.SetText fpSpread1.ColLetterToNumber("BL"), 5 + j, Round(Val(arrValue(4)), 0)
                        UpdateCell fpSpread1.ColLetterToNumber("BL"), 5 + j, Round(Val(arrValue(4)), 0)
                        'UpdateCell fpSpread1.ColLetterToNumber("BH"), 15 + j, arrValue(4)
                    Next j
                End If
        End If
        ' set gia tri tong 32 cho to khai 02_GTGT
        If GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02GTGT" Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            tempCurrSheet = mCurrentSheet
            mCurrentSheet = fpSpread1.SheetCount - 1
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.SetText fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            UpdateCell fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            mCurrentSheet = tempCurrSheet
            fpSpread1.EventEnabled(EventAllEvents) = True
        End If
        ' Set gia tri tong 34 cho to khai 03_GTGT
        If GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03GTGT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01ATNDN" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01BTNDN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01TTDB" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01TAIN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02TAIN" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03TNDN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_05GTGT" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02BVMT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01PHXD" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02TNDN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01BVMT" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02NTNN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03NTNN" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_04NTNN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01TD_GTGT" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03_TD_TAIN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_04GTGT" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01NTNN" Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            tempCurrSheet = mCurrentSheet
            mCurrentSheet = fpSpread1.SheetCount - 1
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.SetText fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            UpdateCell fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            mCurrentSheet = tempCurrSheet
            fpSpread1.EventEnabled(EventAllEvents) = True
        End If
        
        ' I. Dieu chinh tang so thue
        fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
        fpSpread1.EventEnabled(EventAllEvents) = False
        tempCurrSheet = mCurrentSheet
        mCurrentSheet = fpSpread1.SheetCount - 1
        ' xoa dong cu truoc khi them dong
        fpSpread1.Row = 9
        fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
'        fpSpread1.EventEnabled(EventAllEvents) = False
        fpSpread1.sheet = mCurrentSheet
        Do
            countDel = countDel + 1
            fpSpread1.Row = fpSpread1.Row + 1
        Loop Until UCase(fpSpread1.Text) = "AA"
        
        fpSpread1.EventEnabled(EventAllEvents) = False
        For j = 0 To countDel - 1
            DeleteNode mCurrentSheet, fpSpread1.ColLetterToNumber("BD"), 9, False
        Next j
        ' them so dong dieu chinh thay doi vao
        For j = 0 To numRowI - 1
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            InsertNode fpSpread1.ColLetterToNumber("BD"), 9
        Next j
        ' set cac gia tri cua cot
        If flagTang = True Then
            For j = 0 To numRowI
                
                arrValue = Split(arrDieuChinhTang(j), "_")
                fpSpread1.SetText fpSpread1.ColLetterToNumber("B"), 9 + j, j + 1
                
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BE"), 9 + j, arrValue(0)
                UpdateCell fpSpread1.ColLetterToNumber("BE"), 9 + j, arrValue(0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BD"), 9 + j, arrValue(1)
                UpdateCell fpSpread1.ColLetterToNumber("BD"), 9 + j, arrValue(1)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BF"), 9 + j, Round(Val(arrValue(2)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BF"), 9 + j, Round(Val(arrValue(2)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BG"), 9 + j, Round(Val(arrValue(3)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BG"), 9 + j, Round(Val(arrValue(3)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BH"), 9 + j, Round(Val(arrValue(4)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BH"), 9 + j, Round(Val(arrValue(4)), 0)
            Next j
        End If
        
        ' II. Dieu chinh giam so thue
        fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
        fpSpread1.EventEnabled(EventAllEvents) = False
        tempCurrSheet = mCurrentSheet
        mCurrentSheet = fpSpread1.SheetCount - 1
        ' xoa dong cu truoc khi them dong
        fpSpread1.Row = 13 + numRowI
        fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
'        fpSpread1.EventEnabled(EventAllEvents) = False
        fpSpread1.sheet = mCurrentSheet
        Do
            countDel = countDel + 1
            fpSpread1.Row = fpSpread1.Row + 1
        Loop Until UCase(fpSpread1.Text) = "BB"
        
        fpSpread1.EventEnabled(EventAllEvents) = False
        For j = 0 To countDel - 1
            DeleteNode mCurrentSheet, fpSpread1.ColLetterToNumber("BD"), 13 + numRowI, False
        Next j
        ' them so dong dieu chinh thay doi vao
        For j = 0 To numRowII - 1
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            InsertNode fpSpread1.ColLetterToNumber("BD"), 13 + numRowI
        Next j
        ' set cac gia tri cua cot
        If flagGiam = True Then
            For j = 0 To numRowII
                arrValue = Split(arrDieuChinhGiam(j), "_")
                fpSpread1.SetText fpSpread1.ColLetterToNumber("B"), 13 + numRowI + j, j + 1
                
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BE"), 13 + numRowI + j, arrValue(0)
                UpdateCell fpSpread1.ColLetterToNumber("BE"), 13 + numRowI + j, arrValue(0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BD"), 13 + numRowI + j, arrValue(1)
                UpdateCell fpSpread1.ColLetterToNumber("BD"), 13 + numRowI + j, arrValue(1)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BF"), 13 + numRowI + j, Round(Val(arrValue(2)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BF"), 13 + numRowI + j, Round(Val(arrValue(2)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BG"), 13 + numRowI + j, Round(Val(arrValue(3)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BG"), 13 + numRowI + j, Round(Val(arrValue(3)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BH"), 13 + numRowI + j, Round(Val(arrValue(4)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BH"), 13 + numRowI + j, Round(Val(arrValue(4)), 0)
            Next j
        End If

        ' bo set cac cong thuc tinh phat nop cham
'    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
'            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11)
'            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
'            fpSpread1.sheet = fpSpread1.SheetCount - 1
'            fpSpread1.Col = lCol_temp
'            fpSpread1.Row = lRow_temp
'            fpSpread1.Formula = ""
'
'            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10)
'            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
'            fpSpread1.Col = lCol_temp
'            fpSpread1.Row = lRow_temp
'            fpSpread1.Formula = ""
'        Else
'            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
'            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
'            fpSpread1.sheet = fpSpread1.SheetCount - 1
'            fpSpread1.Col = lCol_temp
'            fpSpread1.Row = lRow_temp
'            fpSpread1.Formula = ""
'
'
'            Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
'            ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
'            fpSpread1.Col = lCol_temp
'            fpSpread1.Row = lRow_temp
'            fpSpread1.Formula = ""
'    End If
    ' End set




        mCurrentSheet = tempCurrSheet
        UpdateDataKHBS_TT28 fpSpread1
        'fpSpread1.ActiveSheet = fpSpread1.SheetCount - 1
        DisplayMessage "0222", msOKOnly, miInformation
    End If
End Sub

''' Form_Activate description
''' Resize grid and move form to center screen
''' No parameter
Private Sub Form_Activate()
    On Error GoTo ErrorHandle
        
    ResizeGrid
'    Me.Top = (frmSystem.ScaleHeight - Me.Height) \ 2 + 100
'    Me.Left = (frmSystem.Width - Me.Width) \ 2 - 100
'    If Me.Top < 0 Then Me.Top = 0
'    If Me.Left < 0 Then Me.Left = 0
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Activate", Err.Number, Err.Description
End Sub

''' Form_KeyDown description
''' Form keydown event:
''' When user press F1 -> process help
''' Parameter 1 KeyCode: vbKeyCode
''' Parameter 2 Shift: Ctrl or Alt or Shift key
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandle
    Dim strHelpContexID As String
    Dim i As Integer
    Dim lCol As Long, lRow As Long

    If KeyCode = vbKeyF1 Then
        fpSpread1.sheet = mCurrentSheet
        lCol = fpSpread1.ActiveCol
        lRow = fpSpread1.ActiveRow
        GetCellSpan fpSpread1, lCol, lRow
        strHelpContexID = GetAttribute(xmlDocumentInit(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow)), "HelpContextID") 'Split(GetAttribute(xmlDocumentInit(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow)), "HelpContexID"), "_")
        
'        Sua gan helpcontext=0
        If strHelpContexID <> vbNullString Then
            fpSpread1.HelpContextID = CLng(strHelpContexID) 'Val(strHelpContexID(0) & strHelpContexID(1) & CStr(fpSpread1.ColLetterToNumber(strHelpContexID(2))) & strHelpContexID(3))
        Else
            fpSpread1.HelpContextID = 0
        End If
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "Form_KeyDown", Err.Number, Err.Description
End Sub

''' Form_KeyUp description
''' Form keyup event:
''' When user press Alt + F4 -> process Exit
''' Parameter 1 KeyCode: vbKeyCode
''' Parameter 2 Shift: Ctrl or Alt or Shift key
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 4 Then
        cmdExit_Click
    End If
End Sub

''' Form_Load description
''' Init form: Load interface template, load data, setup grid
''' No parameter
Private Sub Form_Load()
    On Error GoTo ErrorHandle
       
    Dim i As Integer
    Dim lFileNum As Long
    Dim fso As New FileSystemObject
    ' Phuc vu BC26
    Dim numRowI As Integer
    Dim arrResult() As String
    Dim varMenuId As String
    Dim j As Integer
    ' end BC26
            
            
    'hien thi combobox tim kiem
    'dhdang
     Cb_seach.ListIndex = 0
    
    If Dir(TAX_Utilities_New.GetAbsolutePath("..\InterfaceTemplates\Template.xls")) <> "" Then
'        If fpSpread1.IsExcelFile("..\InterfaceTemplates\Template.xls") Then
'            fpSpread1.EventEnabled(EventSheetChanged) = False
'            fpSpread1.ImportExcelBook GetAbsolutePath("..\InterfaceTemplates\Template.xls"), vbNullString
'            fpSpread1.EventEnabled(EventSheetChanged) = True
'        End If
        fpSpread1.EventEnabled(EventAllEvents) = False
        fpSpread1.LoadFromFile "..\InterfaceTemplates\Template.xls"
        fpSpread1.EventEnabled(EventAllEvents) = True
        fpSpread1.Refresh
    End If
    
    i = getFormIndex(TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    arrActiveForm(i).showed = True
    
    If strKHBS = "frmKHBS_BS" Then
        LoadKHBS
        Exit Sub
    End If
        
    mOnLoad = True
    fpSpread1.EventEnabled(EventAllEvents) = False

    If GetAttribute(TAX_Utilities_New.NodeMenu, "Year") = "0" Then
        SetControlCaption Me, IIf(GetAttribute(TAX_Utilities_New.NodeMenu, "FormName") <> "", GetAttribute(TAX_Utilities_New.NodeMenu, "FormName"), "frmCommonInterfaces")
    Else
        SetControlCaption Me, "frmInterfaces"
    End If
    
    LoadTemplate fpSpread1
    SetupSpread
    FormatGrid
    
    Dim idMenu As Variant
    ' set ngay dau quy
    Dim dNgayDauKy As Date
    ' end
    
    idMenu = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")

    If Trim(GetAttribute(TAX_Utilities_New.NodeValidity, "Class")) <> vbNullString Then
        Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_New.NodeValidity, "Class"))
        Set objTaxBusiness.fps = fpSpread1
        ' to khai GTGT se co to khai thang / quy
        If idMenu = "01" Or idMenu = "02" Or idMenu = "04" Or idMenu = "95" Or idMenu = "71" Or idMenu = "36" Or idMenu = "68" Then
             objTaxBusiness.strTkThangQuy = strQuy
        End If
        ' set ngay dau quy
        If idMenu = "01" Or idMenu = "02" Then
            If strQuy = "TK_QUY" Then
                dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_New.ThreeMonths), CInt(TAX_Utilities_New.Year), iNgayTaiChinh, iThangTaiChinh)
                objTaxBusiness.dNgayDauQuy = dNgayDauKy
            End If
        End If
        ' end
        objTaxBusiness.Prepare1
    End If
    Dim idToKhai As Variant
    idToKhai = TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    If idToKhai = "01" Or idToKhai = "11" Or idToKhai = "12" Or idToKhai = "05" Or idToKhai = "03" Or idToKhai = "73" Then
        objTaxBusiness.strLoaiNNKD = strLoaiNNKD
    End If
    LoadStatusFile
    LoadInitFiles
    
    TAX_Utilities_New.AdjustDataReDim fpSpread1.SheetCount - 2
    
    Set objTaxBusiness.fps = Nothing
    fpSpread1.EventEnabled(EventChange) = True
    mOnSetupData = True
    SetupData fpSpread1
    mOnSetupData = False
    fpSpread1.EventEnabled(EventChange) = False
    
    Set objTaxBusiness.fps = fpSpread1
    '***************
    
    ' 10062011
    ' To khai 01_TTDB va NTNN se co to khai phat sinh hoac thang
    If idMenu = "70" Or idMenu = "05" Or idMenu = "81" Or idMenu = "73" Then
        objTaxBusiness.StrTKThang_PS = strLoaiTKThang_PS
    End If
    ' end
    ' To khai 08/TNCN se co to khai theo quy hoac tu thang den thang
    If idMenu = "74" Or idMenu = "75" Then
        objTaxBusiness.strLoaiTKQT = strLoaiTKQT
        objTaxBusiness.strQuy = strQuy
    End If
    
    
    ' Neu la to khai thang/quy TNCN thi nguyen tac van phai ghi nhu cu, phai ghi nhan tung lan bo sung 1
    Dim Parentid As String
    Parentid = TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue
    
    If (Parentid = "101_11") Then
        If strKHBS = "TKCT" Then
            objTaxBusiness.strloaitk = "TKCT"
        ElseIf strKHBS = "TKBS" Then
            objTaxBusiness.strloaitk = "TKBS"
            objTaxBusiness.StrSolanBosung = strSolanBS
            
        End If
    ' Neu la to khai thang/quy TNCN thi nguyen tac van phai ghi nhu hien tai
    ElseIf (Parentid = "101_10") Then
        If strKHBS = "TKCT" And strSolanBS = "" Then
            objTaxBusiness.strloaitk = "TKCT"
            ' Set lai gia tri cua so lan bo sung ve null
            strSolanBS = ""
        ElseIf strKHBS = "TKCT" And Val(strSolanBS) > 0 Then
            objTaxBusiness.strloaitk = "TKCT"
            objTaxBusiness.StrSolanBosung = strSolanBS
            ' Set lai gia tri cua so lan bo sung ve null
            strSolanBS = ""
        ElseIf strKHBS = "TKBS" Then
            objTaxBusiness.strloaitk = "TKBS"
            objTaxBusiness.StrSolanBosung = strSolanBS
            ' Set lai gia tri cua so lan bo sung ve null
            strSolanBS = ""
        End If
    ' Cac to khai khac
    Else
        If strKHBS = "TKCT" Then
            objTaxBusiness.strloaitk = "TKCT"
        ElseIf strKHBS = "TKBS" Then
            objTaxBusiness.strloaitk = "TKBS"
            objTaxBusiness.StrSolanBosung = strSolanBS
        End If
    End If

    ' Ho tro load so ton dau ky BC26
    varMenuId = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
    TAX_Utilities_New.CheckNewDataBC26 = isNewdata
    If Val(varMenuId) = 68 And isNewdata = True Then
        If Not objTaxBusiness Is Nothing Then
             arrResult = objTaxBusiness.loadTonCuoiKy
            numRowI = objTaxBusiness.numRowInsert
            If numRowI >= 0 Then
                fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
                fpSpread1.EventEnabled(EventAllEvents) = False
                For j = 0 To numRowI
                    fpSpread1.EventEnabled(EventAllEvents) = False
                    mCurrentSheet = 1
                    fpSpread1.sheet = mCurrentSheet
                    InsertNode 4, 22
                Next j
            End If
        End If
        fpSpread1.EventEnabled(EventAllEvents) = True
        fpSpread1.Refresh
        ' tinh lai STT
        'dhdang sua loi
        'ngay 21/01/2011
            fpSpread1.sheet = 1
            j = 1
            fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
            fpSpread1.Row = 22
            Do
                 fpSpread1.Text = str(j)
                 fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                 fpSpread1.Row = j + 22
                 j = j + 1
            Loop Until fpSpread1.Text = "aa"
    End If
    'arrResult
    ' end BC26
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepare2
    End If
    
    SetSheetVisible fpSpread1
    
    ' tesst
    If strKHBS = "TKBS" And (varMenuId = "02" Or varMenuId = "01" Or varMenuId = "04" Or varMenuId = "11" Or varMenuId = "12" Or varMenuId = "05" Or varMenuId = "06" _
    Or varMenuId = "86" Or varMenuId = "87" Or varMenuId = "89" Or varMenuId = "71" Or varMenuId = "72" Or varMenuId = "77" Or varMenuId = "03" Or varMenuId = "73" _
    Or varMenuId = "80" Or varMenuId = "81" Or varMenuId = "70" Or varMenuId = "82" Or varMenuId = "83" Or varMenuId = "85") Then
        fpSpread1.sheet = fpSpread1.SheetCount - 1
        fpSpread1.SheetVisible = True
        LoadKHBS_TT28
        If varMenuId = "01" Then
            TonghopKHBS
        End If
        fpSpread1.sheet = 1
        fpSpread1.SheetName = GetAttribute(GetMessageCellById("0120"), "Msg")
    End If
    
    'Set status for first time.
    fpSpread1.ActiveSheet = 1
    fpSpread1.sheet = 1
    mCurrentSheet = 1
    SetStatus fpSpread1.ActiveCol, fpSpread1.ActiveRow


    fpSpread1.EventEnabled(EventAllEvents) = True
    
    '**********************************
    'Update data when import from file
    If strHiddenFormName = "ImportTaxReport" Then
        UpdateData False
        strHiddenFormName = ""
    End If
    '**********************************
    ' Init data version and printing version
    If Not LoadSessionValueFromFile(TAX_Utilities_New.DataFolder & "Session.dat") Then
        Unload Me
        Exit Sub
    End If
    '**********************************
    mOnLoad = False
    hasActiveForm = True
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
End Sub

'Load status msgs from file
Sub LoadStatusFile()
On Error GoTo ErrorHandle
    Set xmlDocumentStatus = New MSXML.DOMDocument
    xmlDocumentStatus.Load App.path & "\Status.xml"
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "LoadStatusFile", Err.Number, Err.Description
End Sub

''' Form_Resize description
''' After form resize -> move button follow form
''' No parameter
Private Sub Form_Resize()
    'fpSpread1.Visible = False
    ResizeButton
    SetFormCaption Me, imgCaption, lblCaption
    'fpSpread1.Visible = True
    'fpSpread1.SetFocus
End Sub

''' Form_Unload description
''' Release memory
''' Parameter1 Cancel    : don't use in this form
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    Dim lSheet As Long
    Dim i As Integer
    
    i = getFormIndex(TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    arrActiveForm(i).showed = False
    
    For lSheet = 0 To TAX_Utilities_New.xmlDataCount
        TAX_Utilities_New.Data(lSheet) = Nothing
    Next

    Set objTaxBusiness = Nothing
    TAX_Utilities_New.NodeValidity = Nothing
    hasActiveForm = False
    
    If strHiddenFormName = "frmTraCuu" Then
        frmTraCuu.Show
    Else
        frmTreeviewMenu.Show
    End If
    strHiddenFormName = ""
    strKHBS = ""
    ResetAdjustData
    Set frmInterfaces = Nothing
    TAX_Utilities_New.NodeValidity = Nothing
    Set arrErrCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Unload", Err.Number, Err.Description
End Sub

''' SetupSpread description
''' Set default properties of grid
''' No parameter
Private Sub SetupSpread()
    On Error GoTo ErrorHandle
    
    Dim lSheet As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .sheet = lSheet
            .AllowCellOverflow = False
            .AllowEditOverflow = True
            .Appearance = AppearanceFlat
            .ArrowsExitEditMode = True
            '.GrayAreaBackColor = RGB(238, 238, 238)
            .GrayAreaBackColor = vbButtonFace
            
            .MaxCols = .DataColCnt - 1
            .MaxRows = .DataRowCnt - 1
            .GridShowHoriz = False
            .GridShowVert = False
                        
            .EditModePermanent = True
            .EditModeReplace = True
            .ColHeadersShow = False
            .RowHeadersShow = False
            .BorderStyle = BorderStyleNone
            .EditEnterAction = EditEnterActionNext
            .ProcessTab = True
            .ScrollBarExtMode = True
            .ScrollBarTrack = ScrollBarTrackOff
            .ScrollBars = ScrollBarsBoth ' ScrollBarsVertical
            .SetActionKey ActionKeyClear, False, False, 0
            .TabStripPolicy = TabStripPolicyAsNeeded
            .TabStripFont.Name = "Tahoma"
            .TextTip = TextTipFloating
        
            If UCase(.SheetName) <> UCase("Header") Then
                .SheetName = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(.sheet - 1), "Caption")
            Else
                mHeaderSheet = .sheet
            End If
            
            
            .SetTextTipAppearance "Tahoma", 8, False, False, RGB(255, 255, 235), &H0
            .Protect = True
        Next
        .ActiveSheet = 1
        .sheet = 1
        mCurrentSheet = 1
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "SetupSpread", Err.Number, Err.Description
End Sub

''' FormatGrid description
''' Set default properties of cell in grid
''' No parameter
Private Sub FormatGrid()
    On Error GoTo ErrorHandle
    
    Dim lSheet As Long, i As Long, j As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .sheet = lSheet
            If .SheetVisible Or .sheet = 1 Or (strKHBS = "TKBS" And .sheet = .SheetCount - 1) Then
                For i = 1 To .MaxRows
                    .Row = i
                    If .RowHeight(i) > 10 And .RowHeight(i) < 15 Then .RowHeight(i) = 14
                    For j = 1 To .MaxCols
                        .Col = j
                        
                        If .BackColor = 12632256 Then
                            'Form backcolor
                            '.BackColor = RGB(238, 238, 238)
                            .BackColor = vbButtonFace
                            Me.BackColor = .BackColor
                        End If
                        
                        If .BackColor = 9868950 Then
                            'Grid header backcolor
                            .BackColor = RGB(215, 215, 215)
                        End If
                        
                        If .BackColor = 16777164 Then
                            'Grid hight light 1 backcolor
                            .BackColor = RGB(233, 245, 254)
                        End If
                        
                        If .BackColor = 13434879 Then
                            'Grid hight light 2 backcolor
                            .BackColor = RGB(255, 255, 235)
                        End If
                        
                        If .CellType = CellTypeNumber Then
                            .TypeNumberDecimal = ","
                            .TypeNumberSeparator = "."
                            .TypeNumberNegStyle = TypeNumberNegStyle1
                            .TypeNumberMax = 99999999999999#
                            .TypeNumberMin = -99999999999999#
                        End If
                        
                        If .CellType = CellTypeDate Then
                            .TypeDateCentury = True
                            .TypeDateFormat = TypeDateFormatDDMMYY
                            .TypeDateSeparator = Asc("/")
                        End If
                        
                        Select Case Trim(.Text)
                            Case "chk"
                                .CellType = CellTypeCheckBox
                                .TypeCheckCenter = True
                            Case "cbo"
                                .CellType = CellTypeComboBox
                                .Text = ""
                            Case "cmd"
                                .CellType = CellTypeButton
                            Case "picture"
                                .CellType = CellTypePicture
                        End Select
                    Next
                Next
            End If
            
            If .sheet = .SheetCount Then
                For i = 1 To .MaxRows
                    .Row = i
                    If .RowHeight(i) > 10 And .RowHeight(i) < 15 Then .RowHeight(i) = 14
                    For j = 1 To .MaxCols
                        .Col = j
                        
                        If .BackColor = 12632256 Then
                            'Form backcolor
                            '.BackColor = RGB(238, 238, 238)
                            .BackColor = vbButtonFace
                            Me.BackColor = .BackColor
                        End If
                        
                        If .BackColor = 9868950 Then
                            'Grid header backcolor
                            .BackColor = RGB(215, 215, 215)
                        End If
                        
                        If .BackColor = 16777164 Then
                            'Grid hight light 1 backcolor
                            .BackColor = RGB(233, 245, 254)
                        End If
                        
                        If .BackColor = 13434879 Then
                            'Grid hight light 2 backcolor
                            .BackColor = RGB(255, 255, 235)
                        End If
                    Next
                Next
            End If
        Next
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "FormatGrid", Err.Number, Err.Description
End Sub

''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Private Function UpdateCell(ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String) As Boolean
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fpSpread1, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow))
    
    If xmlNodeCell Is Nothing Then
        Exit Function
    End If
    
    If GetAttribute(xmlNodeCell, "Value") <> pValue Then
        SetAttribute xmlNodeCell, "Value", pValue
        UpdateCell = True
    End If
    
    Set xmlNodeCell = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog Me.Name, "UpdateCell", Err.Number, Err.Description
End Function

Private Sub fpSpread1_BeforeEditMode(ByVal Col As Long, ByVal Row As Long, ByVal UserAction As FPUSpreadADO.BeforeEditModeActionConstants, CursorPos As Variant, Cancel As Variant)
    If UserAction = BeforeEditModeMouse Then
        'Action executed by Mouse click
        fpSpread1.SetActiveCell Col, Row
        mCurrentSheet = fpSpread1.ActiveSheet
    End If
End Sub

''' fpSpread1_ButtonClicked description
''' Update value for cell (checkbox cell)
''' Parameter1 pCol         : active column
''' Parameter2 pRow         : active row
''' Parameter3 ButtonDown   : left, right or center mouse button
Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    On Error GoTo ErrorHandle
    Dim frmDD As frmDuongDan
    Dim frmOp_Pr As frm_Opcheck
    Dim strFileName As String
    Dim options As Integer
    Dim star As String
    Dim endd As String
    
    If mOnLoad Then Exit Sub
    
    Set frmDD = New frmDuongDan
    Set frmOp_Pr = New frm_Opcheck
       
    With fpSpread1
        .sheet = mCurrentSheet
        GetCellSpan fpSpread1, Col, Row
        .Col = Col
        .Row = Row
        If .CellType = CellTypeCheckBox Then
            UpdateCell Col, Row, IIf(ButtonDown = 1, "x", vbNullString)
        End If
        If Row < 10 Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                strFileName = frmDD.getFileName
                If Trim(strFileName) = vbNullString Or Trim(strFileName) = "" Then
                    Exit Sub
                Else
                    If ImportExcel(strFileName) = True Then
                    'Debug.Print Time
                        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "59" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "70" Then
                            'moveData5A
                            moveDataNKH
                            'dhdang edit
                            'date 08-06-2010
                            'Turning Load BK xong them moi dong(F5)
                            'CallFinish
'                        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" And .ActiveSheet = 2 Then
'                            moveData01_2
                        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "05" Then
                            moveData01TTDB
                        Else
                            moveData
                        End If
                    'Debug.Print Time
                    End If
                End If
            End If
        End If
        'dhdang edit dieu khien cell C_19 to 05_09
        If Row = 19 And Col = .ColLetterToNumber("C") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "45" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 26
                            .Row = i + 22
                            .Col = .ColLetterToNumber("C")
                            .Text = "1"
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 26
                                    .Row = i + 22
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("C")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 25) Then
                                        For i = 0 To .MaxRows - 26
                                        .Row = i + 22
                                        ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                        .Col = .ColLetterToNumber("C")
                                        .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                
                                                    .Col = .ColLetterToNumber("C")
                                                    .Text = "1"
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
            End If
            End If
        End If
        'dhdang
        'xu ly nut check chon tren to 05A(cell N20)
        
        If Row = 20 And Col = .ColLetterToNumber("G") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                'Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 25
                            .Row = i + 22
                            .Col = .ColLetterToNumber("B")
                            If UCase(.Text) = "AA" Then Exit For
                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                            ' begin nvhai edit
                            ' Lay gia tri cua cot MST de kiem tra, vi neu ko co MST thi ko duoc quyet toan tai CQCT
                            .Col = .ColLetterToNumber("E")
                            ' Kiem tra xem co null hay ko? Neu la null thi ko duoc check
                            If Trim(.Text) = "" Or Trim(.Text) = vbNullString Then
                                .Col = .ColLetterToNumber("G")
                                .Text = "0"
                            Else ' Neu ko null thi moi check vao
                                .Col = .ColLetterToNumber("G")
                                .Text = "1"
                            End If
                            ' end nvhai edit
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 26
                                    .Row = i + 22
                                    .Col = .ColLetterToNumber("B")
                                    If UCase(.Text) = "AA" Then Exit For
    
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("G")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 24) Then
                                    
                                        For i = 0 To .MaxRows - 26
                                            .Row = i + 22
                                            .Col = .ColLetterToNumber("B")
                                            If UCase(.Text) = "AA" Then Exit For
                                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                            .Col = .ColLetterToNumber("G")
                                            .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                ' begin nvhai edit
                                                ' Lay gia tri cua cot MST de kiem tra, vi neu ko co MST thi ko duoc quyet toan tai CQCT
                                                .Col = .ColLetterToNumber("E")
                                                ' Kiem tra xem co null hay ko? Neu la null thi ko duoc check
                                                If Trim(.Text) = "" Or Trim(.Text) = vbNullString Then
                                                    .Col = .ColLetterToNumber("G")
                                                    .Text = "0"
                                                Else ' Neu ko null thi moi check vao
                                                    .Col = .ColLetterToNumber("G")
                                                    .Text = "1"
                                                End If
                                                ' end nvhai edit
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
                End If
            End If
        End If
        'xu ly nut chech chon 05A_C20
        If Row = 20 And Col = .ColLetterToNumber("C") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                'Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 25
                            .Row = i + 22
                            .Col = .ColLetterToNumber("B")
                            If UCase(.Text) = "AA" Then Exit For
                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                            .Col = .ColLetterToNumber("C")
                            .Text = "1"
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 26
                                    .Row = i + 22
                                    .Col = .ColLetterToNumber("B")
                                    If UCase(.Text) = "AA" Then Exit For
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("C")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 24) Then
                                    
                                        For i = 0 To .MaxRows - 26
                                            .Row = i + 22
                                            .Col = .ColLetterToNumber("B")
                                            If UCase(.Text) = "AA" Then Exit For
                                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                            .Col = .ColLetterToNumber("C")
                                            .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                                .Col = .ColLetterToNumber("C")
                                                .Text = "1"
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
            End If
            End If
        End If
        
        'xu ly nut chech chon 05B_W19
        If Row = 20 And Col = .ColLetterToNumber("Y") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                'Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 24
                            .Row = i + 22
                             .Col = .ColLetterToNumber("B")
                            If UCase(.Text) = "AA" Then Exit For
                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                            .Col = .ColLetterToNumber("Y")
                            .Text = "1"
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 25
                                    .Row = i + 22
                                     .Col = .ColLetterToNumber("B")
                                    If UCase(.Text) = "AA" Then Exit For
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("Y")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 23) Then
                                    
                                        For i = 0 To .MaxRows - 25
                                        .Row = i + 22
                                        .Col = .ColLetterToNumber("B")
                                        If UCase(.Text) = "AA" Then Exit For
                                        ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                        .Col = .ColLetterToNumber("Y")
                                        .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                .Col = .ColLetterToNumber("B")
                                                If UCase(.Text) = "AA" Then Exit For
                                                ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                                .Col = .ColLetterToNumber("Y")
                                                .Text = "1"
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
            End If
            End If
        End If
        
        ' xu ly nut check chon tren to khai 06KK-TNCN
        'xu ly nut chech chon 06B_C20
        If Row = 20 And Col = .ColLetterToNumber("C") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "59" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                'Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 25
                            .Row = i + 22
                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                            .Col = .ColLetterToNumber("C")
                            .Text = "1"
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 26
                                    .Row = i + 22
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("C")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 24) Then
                                    
                                        For i = 0 To .MaxRows - 26
                                        .Row = i + 22
                                        ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                        .Col = .ColLetterToNumber("C")
                                        .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                                .Col = .ColLetterToNumber("C")
                                                .Text = "1"
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
            End If
            End If
        End If
        
        
        'xu ly nut chech chon 02BH_C20
        If Row = 20 And Col = .ColLetterToNumber("C") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "42" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                'Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 32
                            .Row = i + 22
                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                            .Col = .ColLetterToNumber("C")
                            .Text = "1"
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 32
                                    .Row = i + 22
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("C")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 32) Then
                                    
                                        For i = 0 To .MaxRows - 32
                                        .Row = i + 22
                                        ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                        .Col = .ColLetterToNumber("C")
                                        .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                                .Col = .ColLetterToNumber("C")
                                                .Text = "1"
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
            End If
            End If
        End If
        'xu ly nut chech chon 02XS_C20
        If Row = 20 And Col = .ColLetterToNumber("C") And GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "43" Then
            If .CellType = CellTypeButton Then
                'Dim strFileName As String
                options = frmOp_Pr.getOptions
                'Dim i As Integer
                If options = 1 Then
                    For i = 0 To .MaxRows - 27
                            .Row = i + 22
                            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                            .Col = .ColLetterToNumber("C")
                            .Text = "1"
                    Next
                ElseIf options = 2 Then
                            For i = 0 To .MaxRows - 27
                                    .Row = i + 22
                                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                    .Col = .ColLetterToNumber("C")
                                    .Text = "0"
                            Next
                            
                            ElseIf options = 3 Then
                            star = frmOp_Pr.getStar
                            endd = frmOp_Pr.getEndd
                            If star = "" Or endd = "" Then
                              DisplayMessage "0169", msOKOnly, miCriticalError
                            Else
                                    If star > 0 And endd < (.MaxRows - 27) Then
                                    
                                        For i = 0 To .MaxRows - 27
                                        .Row = i + 22
                                        ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                        .Col = .ColLetterToNumber("C")
                                        .Text = "0"
                                        Next
                                        
                                        For i = star To endd
                                                .Row = i + 21
                                                ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                                                .Col = .ColLetterToNumber("C")
                                                .Text = "1"
                                        Next
                                     Else
                                     DisplayMessage "0168", msOKOnly, miCriticalError
                                     End If
                              End If
                Else
                  'MsgBox "Loi.", vbInformation
            End If
            End If
        End If
        
    End With
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "fpSpread1_ButtonClicked", Err.Number, Err.Description
End Sub

Public Sub taiDuLieu(ByVal strFileName As String)
    
End Sub

''' fpSpread1_Change description
''' Event change of grid
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
Private Sub fpSpread1_Change(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo ErrorHandle
    Dim lValue As String
    Dim IsUpdate As Boolean
    
    If mOnLoad = True Then
        'This action occur only on Setttingup Data
        If mOnSetupData Then
            With fpSpread1
                .Col = Col
                .Row = Row
                
                If Not .Lock Then Exit Sub
                If fpSpread1.sheet = fpSpread1.SheetCount Then Exit Sub
                If TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, Col, Row)) _
                    Is Nothing Then Exit Sub
                    
                .EventEnabled(EventAllEvents) = False
                .sheet = mCurrentSheet
                'This event raise to formula cell
                If .Formula <> "" Then
                        .ReCalcCell Col, Row
                        If .CellType = CellTypeNumber Then
                            lValue = .value
                        Else
                            lValue = .Text
                        End If
                Else
                    Exit Sub
                End If
                '*****************************************
                Select Case .CellType
                    Case 10
                    ' Checkbox -> See Business object
                    Case Else
                        IsUpdate = UpdateCell(Col, Row, lValue)
                End Select
                If Not mOnLoad Then TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_New.AdjustData(mCurrentSheet - 1))
                'End If
                .EventEnabled(EventAllEvents) = True
            End With
        End If
        Exit Sub
    End If
    
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
        .sheet = mCurrentSheet
        
        .Col = Col
        .Row = Row
        If arrErrCells.Exists(mCurrentSheet & "_" & GetCellID(fpSpread1, Col, Row)) Then
            .CellNote = ""
            .BackColor = arrErrCells.Item(mCurrentSheet & "_" & GetCellID(fpSpread1, Col, Row))
            arrErrCells.Remove mCurrentSheet & "_" & GetCellID(fpSpread1, Col, Row)
        End If
        'If .Lock = False Then
        ' When user change value of cell, call UpdateCell function
        
        If .CellType = CellTypeNumber Then
            lValue = .value
        Else
            lValue = .Text
        End If
        Select Case .CellType
            Case 10
            ' Checkbox -> See Business object
            Case Else
                IsUpdate = UpdateCell(Col, Row, lValue)
        End Select
        TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_New.AdjustData(mCurrentSheet - 1))
        'End If
        If .SheetName = "PL 01-1/TTDB" Then
            fpSpread1_LeaveCell Col, Row, Col, Row, True
        End If
        .EventEnabled(EventAllEvents) = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "fpSpread1_Change", Err.Number, Err.Description
End Sub

''' IncreaseRowInDOM description
''' Mapping CellID in DOM with cells on grid
''' call by InsertNode function
''' Parameter1 pRow     : the row inserted
'Private Sub IncreaseRowInDOM(ByVal pRow As Long)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim lCol As Long, lRow As Long, i As Long
'
'    If TAX_Utilities_New.Data(mCurrentSheet - 1) Is Nothing Then Exit Sub
'    Set xmlNodeListCell = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
'
'    For i = xmlNodeListCell.length - 1 To 0 Step -1
'        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
'        If lRow >= pRow Then
'            ' Increase value of row attribute + 1 (CellID)
'            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + 1)
'
'            ' Increase value of row attribute + 1 (CellID2)
'            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
'            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + 1)
'        End If
'    Next
'
'    Set xmlNodeListCell = Nothing
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "IncreaseRowInDOM", Err.Number, Err.Description
'End Sub
'Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim lCol As Long, lRow As Long, i As Long
'
'    If xmlDomData Is Nothing Then Exit Sub
'    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
'
'    For i = xmlNodeListCell.length - 1 To 0 Step -1
'        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
'        If lRow >= pRow Then
'            ' Increase value of row attribute + 1 (CellID)
'            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
'
'            ' Increase value of row attribute + 1 (CellID2)
'            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
'            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
'        End If
'    Next
'
'    Set xmlNodeListCell = Nothing
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "IncreaseRowInDOM", Err.Number, Err.Description
'End Sub

''' DecreaseRowInDOM description
''' Mapping CellID in DOM with cells on grid
''' call by DeleteNode function
''' Parameter1 pRow     : the row deleted
'Private Sub DecreaseRowInDOM(ByVal pRow As Long)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim lCol As Long, lRow As Long, i As Long
'
'    If TAX_Utilities_New.Data(mCurrentSheet - 1) Is Nothing Then Exit Sub
'    Set xmlNodeListCell = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
'
'    For i = 0 To xmlNodeListCell.length - 1
'        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
'        If lRow >= pRow Then
'            ' Decrease value of row attribute - 1 "CellID"
'            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow - 1)
'
'            ' Decrease value of row attribute - 1 "CellID2"
'            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
'            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow - 1)
'        End If
'    Next
'
'    Set xmlNodeListCell = Nothing
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "DecreaseRowInDOM", Err.Number, Err.Description
'End Sub
Private Sub DecreaseRowInDOM(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
    
    If TAX_Utilities_New.Data(intSheet - 1) Is Nothing Then Exit Sub
    Set xmlNodeListCell = TAX_Utilities_New.Data(intSheet - 1).getElementsByTagName("Cell")
    
    For i = 0 To xmlNodeListCell.length - 1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Decrease value of row attribute - 1 "CellID"
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow - lRows)
            
            ' Decrease value of row attribute - 1 "CellID2"
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow - lRow2s)
        End If
    Next
    
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "DecreaseRowInDOM", Err.Number, Err.Description
End Sub


''' InsertRow description
''' Insert new row on grid
''' call by InsertNode function
''' Parameter1 pRow     : the row inserted

'Private Sub InsertRow(ByVal pRow As Long)
'    On Error GoTo ErrorHandle
'
'    Dim i As Long, lBgColor As Long
'
'    With fpSpread1
'        .Sheet = mCurrentSheet
'        .MaxRows = .MaxRows + 1
'        .InsertRows pRow, 1
'        .CopyRowRange pRow - 1, pRow - 1, pRow
'
'        For i = 1 To fpSpread1.MaxCols
'            .col = i
'            .Row = pRow - 1
'            lBgColor = .BackColor
'            .Row = pRow
'            'Set BgColor to inserted cell
'            If lBgColor <> &HC0C0FF Then 'vbRed
'                .BackColor = lBgColor
'            Else
'                .BackColor = vbWhite
'            End If
'            '***************************
''            .Row = pRow
''            .col = i
'            ' Reset empty value for new row on grid
'            If .Lock = False Then
'                If .CellType = CellTypeNumber Then
'                    .SetText i, pRow, 0
'                Else
'                    .SetText i, pRow, vbNullString
'                End If
'                .CellNote = vbNullString
'            End If
'
'        Next
'    End With
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "InsertRow", Err.Number, Err.Description
'End Sub

''' DeleteRow description
''' Delete current row on grid
''' call by DeleteNode function
''' Parameter1 pRow     : the row deleted
'Private Sub DeleteRow(ByVal pRow As Long)
'    On Error GoTo ErrorHandle
'
'    With fpSpread1
'        .Sheet = mCurrentSheet
'        .DeleteRows pRow, 1
'        .MaxRows = .MaxRows - 1
'    End With
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "DeleteRow", Err.Number, Err.Description
'End Sub
Private Sub DeleteRow(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long)
    On Error GoTo ErrorHandle
    
    With fpSpread1
        .EventEnabled(EventChange) = False
        .ReDraw = False
        '.Visible = False
        .sheet = intSheet
        .DeleteRows pRow, lRows
        .MaxRows = .MaxRows - lRows
        '.Visible = True
        .ReDraw = True
        .EventEnabled(EventChange) = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "DeleteRow", Err.Number, Err.Description
End Sub


''' InsertNode description
''' Insert 1 row on grid, insert 1 node in DOM, mapping CellID betweed DOM and grid
''' call when user press F5 ("Dynamic" property is True on data file)
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
'Private Sub InsertNode(ByVal pCol As Long, ByVal pRow As Long)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeCells As MSXML.IXMLDOMNode
'    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
'    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
'    Dim lCol As Long, lRow As Long
'
'    ' Get cellspan for merge cell on interface templates
'    GetCellSpan fpSpread1, pCol, pRow
'
'    Set xmlNodeCells = TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow)).parentNode
'
'    'If Not xmlNodeCells.nextSibling Is Nothing Then GoTo EXIT_SUB
'    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
'    If Val(GetAttribute(xmlNodeCells.parentNode, "MaxRows")) = xmlNodeCells.parentNode.childNodes.length Then GoTo EXIT_SUB
'
'    ' insert new row on grid
'    InsertRow pRow + 1
'
'    ' increase value of row in xmlDocument
'    IncreaseRowInDOM pRow + 1
'
'    Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
'    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
'        ' Set new ID for node (CellID)
'        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
'        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fpSpread1, lCol, lRow + 1)
'
'        ' Set first cell = 1
'        SetAttribute xmlNodeNewCell, "FirstCell", "1"
'
'        ' Reset empty value for new node
'        fpSpread1.col = lCol
'        fpSpread1.Row = lRow
'        If fpSpread1.CellType = CellTypeNumber Then
'            SetAttribute xmlNodeNewCell, "Value", "0"
'        Else
'            SetAttribute xmlNodeNewCell, "Value", vbNullString
'        End If
'
'        ' Set new ID for node (CellID2)
'        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
'        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fpSpread1, lCol, lRow + 1)
'    Next
'
'    ' Insert new node to DOM object
'    If Not xmlNodeCells.nextSibling Is Nothing Then
'        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, xmlNodeCells.nextSibling
'    Else
'        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
'    End If
'
'EXIT_SUB:
'    Set xmlNodeNewCell = Nothing
'    Set xmlNodeNewCells = Nothing
'    Set xmlNodeCells = Nothing
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "InsertNode", Err.Number, Err.Description
'End Sub
Public Sub InsertNode(ByVal pCol As Long, ByVal pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    
    ' Get cellspan for merge cell on interface templates
    GetCellSpan fpSpread1, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fpSpread1, xmlNodeCells, lRow2s, lLRowBound, lURowBound)
    
    'If Not xmlNodeCells.nextSibling Is Nothing Then GoTo EXIT_SUB
    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    If Val(GetAttribute(xmlNodeCells.parentNode, "MaxRows")) = xmlNodeCells.parentNode.childNodes.length Then GoTo EXIT_SUB
    
    ' insert new row on grid
    InsertRow fpSpread1, lURowBound + 1, lRows
    'fpSpread1.SetFocus
    
    ' increase value of row in xmlDocument
    IncreaseRowInDOM fpSpread1, TAX_Utilities_New.Data(mCurrentSheet - 1), lURowBound + 1, lRows, lRow2s
    'IncreaseRowInDOM lURowBound + 1, lRows, lRow2s

    Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
                
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
        
        ' Reset empty value for new node
        fpSpread1.Col = lCol
        fpSpread1.Row = lRow
        Select Case fpSpread1.CellType
            Case CellTypeNumber
                SetAttribute xmlNodeNewCell, "Value", "0"
            Case Else
                SetAttribute xmlNodeNewCell, "Value", vbNullString
        End Select
        
        ' Set new ID for node (CellID2)
        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
    Next
    
    ' Insert new node to DOM object
    If Not xmlNodeCells.nextSibling Is Nothing Then
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, xmlNodeCells.nextSibling
    Else
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
    End If
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "InsertNode", Err.Number, Err.Description
End Sub


''' DeleteNode description
''' Delete 1 row on grid, delete 1 node in DOM, mapping CellID betweed DOM and grid
''' call when user press F6 ("Dynamic" property is True on data file)
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
'Private Sub DeleteNode(ByVal pCol As Long, ByVal pRow As Long)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeCells As MSXML.IXMLDOMNode, xmlNodeTemp As MSXML.IXMLDOMNode
'    Dim xmlNodeTemp2 As MSXML.IXMLDOMNode
'    Dim lCol As Long, lRow As Long
'
'    GetCellSpan fpSpread1, pCol, pRow
'
'    Set xmlNodeCells = TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow)).parentNode
'
'    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
'
'    If xmlNodeCells.parentNode.childNodes.length <= 1 Then
'        ClearRow pRow
'        GoTo EXIT_SUB
'    Else
'        If GetAttribute(xmlNodeCells.firstChild, "FirstCell") = "0" Then
'            'Set FirstCell attr to "0" for next Cells node
'            Set xmlNodeCells = xmlNodeCells.nextSibling
'            SetAttribute xmlNodeCells.firstChild, "FirstCell", "0"
'            Set xmlNodeCells = xmlNodeCells.previousSibling
'        End If
'    End If
'
'    'If xmlNodeCells.xml = xmlNodeCells.parentNode.childNodes(0).xml Then GoTo EXIT_SUB
'
'    ' Delete curent row from grid
'    DeleteRow pRow
'
'    xmlNodeCells.parentNode.removeChild xmlNodeCells
'
'    ' Decrease value of row in xmlDocument
'    DecreaseRowInDOM pRow + 1
'
'    fpSpread1.col = fpSpread1.ActiveCol
'    fpSpread1.Row = fpSpread1.ActiveRow
'    If fpSpread1.Lock = True Then
'        fpSpread1.SetActiveCell fpSpread1.ActiveCol, fpSpread1.ActiveRow - 1
'    End If
'EXIT_SUB:
'    Set xmlNodeCells = Nothing
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "DeleteNode", Err.Number, Err.Description
'End Sub

Private Sub DeleteNode(ByVal intSheet As Integer, ByVal pCol As Long, ByVal pRow As Long, Optional ByVal blnForce As Boolean = True)
    On Error GoTo ErrorHandle
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    Dim xmlNodeCells As MSXML.IXMLDOMNode, xmlNodeTemp As MSXML.IXMLDOMNode
    Dim xmlNodeTemp2 As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    
    GetCellSpan fpSpread1, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_New.Data(intSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fpSpread1, xmlNodeCells, lRow2s, lLRowBound, lURowBound)

    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    
'*********************************************************
' added
'Date: 01/03/06
    'Check whether user want to delete
    If lRows > 1 And blnForce And Not IsEmptyValue(xmlNodeCells) Then
        If DisplayMessage("0075", msYesNo, miQuestion, , mrYes) = mrNo Then
            Exit Sub
        End If
    End If
'*********************************************************
        
    If xmlNodeCells.parentNode.childNodes.length <= 1 Then
        ClearRows xmlNodeCells
        TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
        GoTo EXIT_SUB
    Else
        If GetAttribute(xmlNodeCells.firstChild, "FirstCell") = "0" Then
            'Set FirstCell attr to "0" for next Cells node
            Set xmlNodeCells = xmlNodeCells.nextSibling
            SetAttribute xmlNodeCells.firstChild, "FirstCell", "0"
            Set xmlNodeCells = xmlNodeCells.previousSibling
        End If
    End If
    
    'If xmlNodeCells.xml = xmlNodeCells.parentNode.childNodes(0).xml Then GoTo EXIT_SUB
    
    'Jump active cell to prevous section
    'fpSpread1.SetActiveCell fpSpread1.ActiveCol, fpSpread1.ActiveRow - lRows
    
    ' Delete curent row on Form
    DeleteRow intSheet, lLRowBound, lRows
    'fpSpread1.SetFocus
    
    xmlNodeCells.parentNode.removeChild xmlNodeCells
    
    ' Decrease value of row in xmlDocument
    DecreaseRowInDOM intSheet, lLRowBound + 1, lRows, lRow2s
    
    fpSpread1.Col = fpSpread1.ActiveCol
    fpSpread1.Row = fpSpread1.ActiveRow
    If fpSpread1.Lock = True Then
        Do
            fpSpread1.Row = fpSpread1.Row - 1
        Loop Until (Not fpSpread1.Lock) Or (fpSpread1.Row = 1)
        fpSpread1.SetActiveCell fpSpread1.Col, fpSpread1.Row
    End If
'   ************************************
'    added
'   Date: 12/04/06
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
'   ************************************
EXIT_SUB:
    Set xmlNodeCells = Nothing
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "DeleteNode", Err.Number, Err.Description
End Sub


''' fpSpread1_Click description
''' Event fpSpread1_Click
''' allow user edit on cell
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
Private Sub fpSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lCurSheet As Long
    Dim lCurCol As Long, lCurRow As Long
    'If BusinessObjOnProcess(objTaxBusiness) = True Then Exit Sub
    With fpSpread1
        .ArrowsExitEditMode = False
        'Backup sheet, col, row values
        lCurSheet = .sheet
        lCurCol = .Col
        lCurRow = .Row
        
        .sheet = .ActiveSheet
        .Col = Col
        .Row = Row
        If Not (.CellType = CellTypeCheckBox Or .CellType = CellTypeButton) Then
        '*********************************
        'Sua loi xoc xech form
            .Refresh
        '*********************************
            GetCellSpan fpSpread1, Col, Row
            .SetActiveCell Col, Row
            
            '*********************************
            If GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "05_TNCN" Then
                If .sheet = 2 And .Col = .ColLetterToNumber("D") And .Row = 5 Then
                    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bang_Ke_05AK.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
                End If
                If .sheet = 3 And .Col = .ColLetterToNumber("C") And .Row = 4 Then
                    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bang_Ke_05BK.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
                End If
                If .sheet = 4 And .Col = .ColLetterToNumber("C") And .Row = 4 Then
                    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "PhuLuc_01.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
                End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "06_TNCN10" Then
                If .sheet = 2 And .Col = .ColLetterToNumber("D") And .Row = 2 Then
                    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bangke_06BTNCN.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
                End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "09_TNCN" Then
            ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_BH" Then
                If .sheet = 2 And .Col = .ColLetterToNumber("D") And .Row = 3 Then
                    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bang_Ke_02ABK_BH.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
                End If
            ElseIf GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") = "02_TNCN_XS" Then
                If .sheet = 2 And .Col = .ColLetterToNumber("D") And .Row = 3 Then
                    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bang_Ke_02ABK_XS.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
                End If
            End If
        End If
        
        'Restore sheet, col, row value
        .sheet = lCurSheet
        .Col = lCurCol
        .Row = lCurRow
    End With
End Sub

Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Mode = 1 Then
        With fpSpread1
            .sheet = mCurrentSheet
            .Col = Col
            .Row = Row
            If .CellType = CellTypeNumber Then
                .TypeNumberNegStyle = TypeNumberNegStyle2
            End If
        End With
    End If
End Sub

''' fpSpread1_KeyDown description
''' Event fpSpread1_KeyDown
''' allow user edit on cell
''' Parameter1 KeyCode   : return vbKeyCode when user press keys
''' Parameter2 Shift     : return Shift when user press keys
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        fpSpread1.ArrowsExitEditMode = False
        Exit Sub
    End If
End Sub

''' fpSpread1_KeyUp description
''' Event fpSpread1_KeyUp
''' allow user insert new row on grid
''' allow user delete active row on grid
''' Parameter1 KeyCode   : return vbKeyCode when user press keys
''' Parameter2 Shift     : return Shift when user press keys
Private Sub fpSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandle
    
    Dim lCol As Long, lRow As Long
    Dim i As Long
    
    ' Neu la cac mau in tong hop tu to quyet toan 05TNCN->09TNCN va cac chung tu cua TNCN thi cung bo qua
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "45" Then Exit Sub
    ' Neu la to khai 04/GTGT
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "71" Then
        If Not objTaxBusiness.isAddRow(fpSpread1.ActiveCol, fpSpread1.ActiveRow) Then
            Exit Sub
        End If
    End If
    
    If Not ((KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Or (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyEscape)) Then Exit Sub
    
    fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
    fpSpread1.EventEnabled(EventAllEvents) = False
    If KeyCode = vbKeyF5 Then
     ' xu ly cho to khai 04GTGT
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "71" Then
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
            fpSpread1.Row = 45
            i = 1
            Do
                If fpSpread1.ActiveRow = fpSpread1.Row Then
                    objTaxBusiness.lViTri = i
                    objTaxBusiness.strMaNhomAdd = fpSpread1.Text
                    Exit Do
                End If
                i = i + 1
                fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                fpSpread1.Row = fpSpread1.Row + 1
            Loop Until fpSpread1.Text = "aa"
        End If
    
        If objTaxBusiness.InsertEnable(KeyCode, Shift) Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            InsertNode fpSpread1.ActiveCol, fpSpread1.ActiveRow
        End If
    End If
    If KeyCode = vbKeyF6 Then
        ' xu ly cho to khai 04GTGT
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "71" Then
             fpSpread1.sheet = mCurrentSheet
            fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
            fpSpread1.Row = 45
            i = 1
            Do
                If fpSpread1.ActiveRow = fpSpread1.Row Then
                    objTaxBusiness.lViTri = i
                    objTaxBusiness.strMaNhomAdd = fpSpread1.Text
                    Exit Do
                End If
                i = i + 1
                fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                fpSpread1.Row = fpSpread1.Row + 1
            Loop Until fpSpread1.Text = "aa"
        End If
        
        If objTaxBusiness.DeleteEnable(KeyCode, Shift) Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            DeleteNode mCurrentSheet, fpSpread1.ActiveCol, fpSpread1.ActiveRow
        End If
    End If
    If KeyCode = vbKeyDelete Then
        If objTaxBusiness.DeleteEnable(KeyCode, Shift) Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.Col = fpSpread1.ActiveCol
            fpSpread1.Row = fpSpread1.ActiveRow
            If fpSpread1.CellType = CellTypeComboBox Then
                fpSpread1.Text = vbNullString
                UpdateCell fpSpread1.ActiveCol, fpSpread1.ActiveRow, vbNullString
                If (Not objTaxBusiness Is Nothing) Then objTaxBusiness.CellChange fpSpread1.Col, fpSpread1.Row
                TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
            End If
        End If
    End If
    fpSpread1.EventEnabled(EventAllEvents) = True
    fpSpread1.Refresh
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "fpSpread1_KeyUp", Err.Number, Err.Description
    fpSpread1.Refresh
End Sub

Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim count2, count3 As Long
Dim str(20) As Variant
Dim sum1(20), sum2(20) As Variant
Dim i, j, k, l, exist, exist1, exist1_num, inserted As Long
    With fpSpread1
        .sheet = mCurrentSheet
        .Col = Col
        .Row = Row
        If .CellType = CellTypeNumber Then
            .TypeNumberNegStyle = TypeNumberNegStyle1
        End If
        If .SheetName = "PL 01-1/TTDB" Then
'tinh so dong
            .Col = .ColLetterToNumber("B")
            .Row = 37
            Do While .Text <> "aa"
                 count2 = count2 + 1
                 .Row = count2 + 37
            Loop
'tinh so dong tong
            .Col = .ColLetterToNumber("B")
            .Row = 40 + count2
            Do While .Text <> "aa"
                 count3 = count3 + 1
                 .Row = count3 + 40 + count2
            Loop
'chuyen? vao phan tong? cong
            .Row = 37
            .Col = .ColLetterToNumber("L")
       'tinh exist and move data vao str(),sum1(),sum2()
            Do
                If exist <> 0 Then
                    exist1 = 0
                    i = 0
                    Do
                        If .Text <> "" And .Text = str(i) Then
                            exist1 = 1
                            exist1_num = i
                        Else
                            
                        End If
                        i = i + 1
                    Loop Until i = exist
                    
                    If exist1 = 0 And .Text <> "" Then
                        str(exist) = .Text
                        .Col = .ColLetterToNumber("N")
                        sum1(exist) = sum1(exist) + .value
                        .Col = .ColLetterToNumber("P")
                        sum2(exist) = .value
                        .Col = .ColLetterToNumber("L")
                        exist = exist + 1
                    ElseIf exist1 = 1 And .Text <> "" Then
                        .Col = .ColLetterToNumber("N")
                        sum1(exist1_num) = sum1(exist1_num) + Conversion.CDbl(.value)
                        .Col = .ColLetterToNumber("P")
                        sum2(exist1_num) = sum2(exist1_num) + Conversion.CDbl(.value)
                        .Col = .ColLetterToNumber("L")
                    End If
                Else
                    If .Text <> "" Then
                        str(0) = .Text
                        .Col = .ColLetterToNumber("N")
                        sum1(0) = sum1(exist) + .value
                        .Col = .ColLetterToNumber("P")
                        sum2(0) = .value
                        .Col = .ColLetterToNumber("L")
                        exist = exist + 1
                    End If
                End If
                .Row = .Row + 1
            Loop Until .Row = 37 + count2
            
            If exist <> 0 Then
                .Row = 37 + count2
                k = count3
'them bot dong tong
                If exist > k Then
                    Do
                        InsertNode .ActiveCol, 40 + count2
                        k = k + 1
                    Loop Until k = exist
                 End If
                k = count3
                If exist < k Then
                    Do
                        DeleteNode .sheet, .ActiveCol + 2, 40 + count2
                        k = k - 1
                        .Refresh
                    Loop Until k = exist
                End If
'chuyen du lieu vao dong tong
        
                'dntai them bien tam de luu gia tri row
                Dim rowTemp As Integer
                .Row = 40 + count2

                If exist >= 1 Then
                    j = 0
                    Do
                        rowTemp = .Row
                        .Col = .ColLetterToNumber("L")
                        .value = str(j)
                        UpdateCell .Col, .Row, .Text
                        .Col = .ColLetterToNumber("N")
                        .value = sum1(j)
                        UpdateCell .Col, .Row, .value
                        .Col = .ColLetterToNumber("P")
                        .value = sum2(j)
                        
                        'luu lai bien row
                        .Row = rowTemp
                        UpdateCell .Col, .Row, .value
                        .Col = .ColLetterToNumber("L")
                        .Row = .Row + 1
                        
                        j = j + 1
                    Loop Until j = exist
                End If
            Else
                If count3 > 1 Then
                    Do
                        DeleteNode .sheet, .ActiveCol, 39 + count2 + count3
                        count3 = count3 - 1
                    Loop Until count3 = 1
                    .Row = 41
                    .Col = .ColLetterToNumber("L")
                    .value = ""
                    UpdateCell .Col, .Row, .value
                Else
                        .Row = 41
                        .Col = .ColLetterToNumber("L")
                        .value = ""
                        UpdateCell .Col, .Row, .Text
                        .Col = .ColLetterToNumber("N")
                        .value = 0
                        UpdateCell .Col, .Row, .value
                        .Col = .ColLetterToNumber("P")
                        .value = 0
                        UpdateCell .Col, .Row, .value
                        .Col = .ColLetterToNumber("L")
                End If
            End If
        Else
        .Refresh
        End If
    End With
        fpSpread1.ArrowsExitEditMode = True
        SetStatus NewCol, NewRow
End Sub

''' fpSpread1_SheetChanged description
''' Event fpSpread1_SheetChanged
''' mapping value betweed active sheet and mCurrentSheet variable
''' Parameter1 OldSheet   : sheet before change
''' Parameter2 NewSheet   : sheet after change
Private Sub fpSpread1_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    On Error GoTo ErrorHandle
    
    fpSpread1.sheet = NewSheet
    mCurrentSheet = NewSheet
    If fpSpread1.SheetVisible = True Then
        SetStatus fpSpread1.ActiveCol, fpSpread1.ActiveRow
    End If
    'Begin dhDang edit
    ' Trong truong hop la to khai quyet toan TNCN thi frame3 moi hien thi
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "17" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "42" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "43" Or GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "59" Then
        If ((GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(1), "Active") <> "0") And NewSheet = 2) Or ((GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(2), "Active") <> "0") And NewSheet = 3) Then

            Frame3.Visible = True
            
            Frame3.Left = 10
            Frame3.Width = Frame1.Width + 150
            'Frame3.Height = 300
        
        Else
            Frame3.Visible = False
        End If
    End If
    'End dhDang edit
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "fpSpread1_SheetChanged", Err.Number, Err.Description
End Sub

''' SetCellNote description
''' Set CellNote for error cell
''' Parser pCellString (containt sheetname and cellID)
''' Parameter1 pCellString  : containt sheetname and cellID
''' Parameter2 pNoteText    : the string input into cellnote
Private Function SetCellNote(ByVal pCellString As String, ByVal lNoErrColor As Long, ByVal pNoteText As String, Optional blnWarning As Boolean = False) As Boolean
    On Error GoTo ErrorHandle
    
    Dim lAnchor As Long
    Dim lSheetName As String, lCellString As String, lStringTemp As String
    Dim lCol As Long, lRow As Long, i As Long
    Dim mResult As Integer
    
    SetCellNote = True
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For i = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, 1)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, 1)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For i = 1 To .SheetCount - 1
            .sheet = i
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Or UCase(.SheetName) = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
'                If blTestVisibleSheet = True And .SheetVisible = False Then
'                    'if sheet of PL is invisible, ask user
'                    If DisplayMessage("0042", msYesNo, miQuestion) = mrYes Then
'                        TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue = "1"
'                        .SheetVisible = True
'                    Else
'                        Exit For
'                    End If
'                End If
                If Not GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(.sheet - 1), "Active") <> "0" Then
                    .Col = lCol
                    .Row = lRow
                    .CellNote = ""
                    .BackColor = lNoErrColor
                    SetCellNote = False
                    Exit Function
                End If
                .Col = lCol
                .Row = lRow
                
                If Trim(pNoteText) = "" Then
                    .CellNote = ""
                ElseIf Trim(.CellNote) = "" Then
                    .CellNote = pNoteText
                Else
                    .CellNote = .CellNote & vbCrLf & pNoteText
                End If
                'If .Lock = False Then
                    If Trim(.CellNote) <> "" Then
                        If Not blnWarning Then
                            .BackColor = &HC0C0FF   'VB 'vbRed
                        Else
                            .BackColor = 12713215 'Vb Yellow '16777088   'VB 'vbgreen
                        End If
                    Else
                        .BackColor = lNoErrColor
                    End If
                'End If
                Exit For
            End If
        Next
    End With
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog Me.Name, "SetCellNote", Err.Number, Err.Description
End Function

''' get Sheet, Col, Row from Cell Formula
'''Parameter: Cell Formula string
'''Parameter: sheet integer
'''parameter: Col integer
'''parameter: Row integer
Private Sub getCellPosition(pCellString As String, lSheet As Long, lCol As Long, lRow As Long)
    On Error GoTo ErrorHandle
    
    Dim lAnchor As Long
    Dim lSheetName As String, lCellString As String, lStringTemp As String
    Dim i As Long
    
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For i = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, 1)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, 1)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For i = 1 To .SheetCount
            .sheet = i
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Or UCase(.SheetName) = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
                lSheet = i
                Exit For
            End If
        Next
    End With
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "getCellPosition", Err.Number, Err.Description
End Sub

''' CheckValidData description
''' Check all formula in last sheet, if error put the notetext into cellnode
''' No parameter
''' Return True if no error checking
''' Return False if one or more error occur
Private Function CheckValidData() As Boolean
    On Error GoTo ErrorHandle
    
    Dim i As Long, lNoErrColor As Long
    Dim strCellString As String
    
    
    Dim vFunction As Variant, vCell As Variant
    Dim vMsg As Variant, vWarning As Variant
    Dim vOrder As Variant, vFormulaFunc As Variant
    Dim cOrder As New Collection
    Dim sheet1 As Long
    
    Dim vGroupTK As String
    
    CheckValidData = True
    If checkCauTrucData = False Then
        CheckValidData = False
    End If
    
    ' Doi voi truong hop la to khai bo sung thi ko checkValidData, nhung truong hop bo sung cua TNCN thi van phai check
    vGroupTK = TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue
    If strKHBS = "TKBS" And (vGroupTK <> "101_11" And vGroupTK <> "101_1" And vGroupTK <> "101_2" And vGroupTK <> "101_3" And vGroupTK <> "101_4" And vGroupTK <> "101_8") Then
        Exit Function
    End If
    
    '*****************************
    ' added
'    Dim strArrActive() As String
'
'    'Backup node validity
'    For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'        ReDim Preserve strArrActive(i)
'        strArrActive(i) = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(i), "Active")
'    Next i
'
'    If Not objTaxBusiness Is Nothing Then
'        For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'            Call objTaxBusiness.SetActiveSheet(TAX_Utilities_New.NodeValidity.childNodes(i))
'        Next i
'    End If
    '*****************************
    With fpSpread1
'        Do While sheet1 <> .SheetCount
'            .sheet = sheet1 + 1
'            If .SheetVisible Then
'                delNullRow sheet1
'            End If
'            sheet1 = sheet1 + 1
'        Loop
    '**************************
    ' remove
    'Reason: Move these commands into ResetErrorCells procedure
    
        .ReDraw = False
        If .SheetCount = 1 Then Exit Function
'        .Sheet = mHeaderSheet
'
'        For i = 12 To .MaxRows
'            .Sheet = mHeaderSheet
'            .col = 2
'            .Row = i
'            If .Formula <> vbNullString Then
'                .col = .col + 1
'                strCellString = .Formula
'                If strCellString <> vbNullString Then _
'                    SetCellNote strCellString, .BackColor, ""
'            End If
'        Next
'    **************************
'    '**************************
'    ' added
'        If .SheetCount = 1 Then Exit Function
'        ResetErrorCells
'    '**************************
    
        Dim isSet As Boolean
        'set error note for cell
        .sheet = mHeaderSheet
        For i = 12 To .MaxRows
            .sheet = mHeaderSheet
            .Col = 2
            .Row = i
            vFormulaFunc = .Formula
            
            If Trim(.Text) <> "" Then
                .GetText .ColLetterToNumber("B"), i, vFunction
                .GetText .ColLetterToNumber("E"), i, vMsg
                .GetText .ColLetterToNumber("S"), i, vWarning
                .GetText .ColLetterToNumber("T"), i, vOrder
                .Col = .Col + 1
                vCell = .Formula
                lNoErrColor = .BackColor
                If vFormulaFunc <> vbNullString Then
                    If Val(vFunction) <> 1 Then
                        If UCase(Trim(vWarning)) = "Y" Then
                            isSet = SetCellNote(vCell, lNoErrColor, " " & vMsg, False)
                            If Trim(vCell) <> "" And isSet = True Then
                                cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                                CheckValidData = False
                            End If
                        ElseIf UCase(Trim(vWarning)) = "N" Then
                            isSet = SetCellNote(vCell, lNoErrColor, " " & vMsg, True)
'htphuong edit
                            If Trim(vCell) <> "" And isSet = True Then
                                cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                            End If
                        End If
                    End If
                Else 'Dynamic
                    If Val(vFunction) <> 1 Then
                        If Trim(vCell) <> "" And UCase(Trim(vWarning)) = "Y" Then cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                        If UCase(Trim(vWarning)) = "Y" Then CheckValidData = False
                    End If
                End If
            End If
        Next
        
        'focus on the first error cell
        Dim min As Long, X As Long, strCell As String
        Dim lSheet As Long, lCol As Long, lRow As Long
                
        If cOrder.count > 0 Then
            min = Val(Left(cOrder(1), InStr(cOrder(1), "[]")))
            strCell = Right(cOrder(1), Len(cOrder(1)) - InStr(cOrder(1), "[]") - 1)
            For i = 2 To cOrder.count
                X = Val(Left(cOrder(i), InStr(cOrder(i), "[]")))
                If min >= X Then
                    min = X
                    strCell = Right(cOrder(i), Len(cOrder(i)) - InStr(cOrder(i), "[]") - 1)
                End If
            Next
            getCellPosition strCell, lSheet, lCol, lRow
            .sheet = lSheet
            .ActiveSheet = lSheet
            .SetActiveCell lCol, lRow
            .EventEnabled(EventAllEvents) = False
            .SetFocus
            .EventEnabled(EventAllEvents) = True
        End If
        .ReDraw = True
    End With
    
    'Restore active properties of node validity
'    For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(i), "Active", strArrActive(i)
'    Next i
    
    Exit Function
    
ErrorHandle:
    'Restore active properties of node validity
'    For i = 0 To TAX_Utilities_New.NodeValidity.childNodes.length - 1
'        SetAttribute TAX_Utilities_New.NodeValidity.childNodes(i), "Active", strArrActive(i)
'    Next i
    SaveErrorLog Me.Name, "CheckValidData", Err.Number, Err.Description
End Function

''' ResizeGrid description
''' Resize grid after load data
''' No parameter
Private Sub ResizeGrid()
    On Error GoTo ErrorHandle
    
    Dim lSheet As Integer
    Dim i As Long, lColWidth As Long, lGridWidth As Long, lMaxGridWidth As Long
    Dim lRowHeight As Long, lGridHeight As Long, lMaxGridHeight As Long
    
    
'    With fpSpread1
'        For lSheet = 1 To .SheetCount
'            .sheet = lSheet
'            lGridWidth = 0
'            lGridHeight = 0
'            If UCase(.SheetName) <> UCase("Header") Then
'                ' Calculated grid width
'                For i = 1 To .MaxCols
'                    .ColWidthToTwips .ColWidth(i), lColWidth
'                    lGridWidth = lGridWidth + lColWidth
'                Next
'                If lMaxGridWidth < lGridWidth Then lMaxGridWidth = lGridWidth
'
'                ' Calculated grid height
'                For i = 1 To .MaxRows
'                    .RowHeightToTwips i, .RowHeight(i), lRowHeight
'                    lGridHeight = lGridHeight + lRowHeight
'                Next
'                If lMaxGridHeight < lGridHeight Then lMaxGridHeight = lGridHeight
'            End If
'        Next
'
'        If .Width > lMaxGridWidth + 200 Then
'            .Width = lMaxGridWidth + 200 + IIf(lMaxGridWidth < 9000, 9000 - lMaxGridWidth, 0)
'            Me.Width = lMaxGridWidth + 330 + IIf(lMaxGridWidth < 9000, 9000 - lMaxGridWidth, 0)
'        End If
'
'        If .Height > lMaxGridHeight + 400 Then
'            .Height = lMaxGridHeight + 400
'            Me.Height = .Height + 1200
'        End If
'
'    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "ResizeGrid", Err.Number, Err.Description
End Sub

''' ResizeButton description
''' Resize button after resize form
''' No parameter
Private Sub ResizeButton()
    Dim menuID As String
    On Error GoTo ErrorHandle
    
    ' Resize width
    Frame1.Left = 50
    Frame1.Width = Me.Width - 300
    Frame1.Height = Me.Height - 1500
        
    fpSpread1.Left = 100
    fpSpread1.Width = Frame1.Width - 200
    fpSpread1.Height = Frame1.Height - 300
        
    Frame2.Left = 50
    Frame2.Top = Frame1.Height + 250
    Frame2.Width = Frame1.Width
    Frame2.Height = 800
    
    ' Begin dhDang edit
    ' Truong hop dac biet man hinh ho tro in quyet toan cho ca nhan thi hien thi luon phan tim kiem
    menuID = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
    If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "45" Then
        Frame3.Visible = True
        Frame3.Left = 10
        Frame3.Width = Frame1.Width + 150
    End If
    ' End dhDang edit
    
    ' Remove button
    If GetAttribute(TAX_Utilities_New.NodeMenu, "Year") <> "0" Then
    'If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") <> "100_1" Then
        ' Business function
        cmdLoadToKhai.Left = Frame1.Width - 8235  ' 9425
        cmdInsert.Left = Frame1.Width - 8235
        cmdClear.Left = Frame1.Width - 7045
        cmdSave.Left = Frame1.Width - 5855
        cmdPrint.Left = Frame1.Width - 4710
        cmdDelete.Left = Frame1.Width - 3565
        cmdKiemTra.Left = Frame1.Width - 3565
        cmdExport.Left = Frame1.Width - 2420
        cmdExit.Left = Frame1.Width - 1240
        If (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Then
            cmdDelete.Left = Frame1.Width - 9370
            cmdKiemTra.Visible = True
        End If
    Else
        ' Header informations
        cmdLoadToKhai.Visible = False
        cmdInsert.Visible = False
        cmdClear.Visible = False
        cmdPrint.Visible = False
        cmdDelete.Visible = False
        cmdExport.Visible = False
        cmdSave.Left = Frame1.Width - 2755 '2985
        cmdExit.Left = Frame1.Width - 1405 '1635
    End If
'****************************************************
    If strKHBS = "frmKHBS_BS" Then
        cmdLoadToKhai.Visible = False
        cmdDelete.Left = Frame1.Width - 8005 '4335
        cmdClear.Left = Frame1.Width - 6685 '8385
        cmdSave.Left = Frame1.Width - 5365 '7035
        cmdPrint.Left = Frame1.Width - 4045 '5685
        
        cmdExport.Left = Frame1.Width - 2725 '2985
        cmdExit.Left = Frame1.Width - 1405 '1635
        
        cmdInsert.Visible = False
        'cmdClear.Visible = False
        'cmdPrint.Visible = False
        'cmdDelete.Visible = False
        'cmdExport.Visible = False
        'cmdSave.Left = Frame1.Width - 2755 '2985
        'cmdExit.Left = Frame1.Width - 1405 '1635
    
    End If
    
    ' doi voi cac to khai bo sung se khng hien thi nut them phu luc
    If strKHBS = "TKBS" Then
        cmdInsert.Visible = False
    End If
    ' set cac to khai bo sung theo TT28 moi hien thi nut tong hop to khai
    If strKHBS = "TKBS" And (menuID = "01" Or menuID = "02" Or menuID = "04" Or menuID = "71" Or menuID = "72" _
    Or menuID = "11" Or menuID = "12" Or menuID = "06" Or menuID = "05" Or menuID = "86" Or menuID = "87" Or menuID = "89" Or menuID = "77" Or menuID = "03" Or menuID = "73" _
    Or menuID = "80" Or menuID = "81" Or menuID = "70" Or menuID = "82" Or menuID = "83" Or menuID = "85") Then
        Command1.Visible = True
        Command1.Left = Frame1.Width - 8460
    Else
        Command1.Visible = False
    End If
    

    If fpSpread1.SheetCount <= 3 Then
        cmdInsert.Visible = False
    End If
'****************************************************
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "ResizeButton", Err.Number, Err.Description
End Sub

''' LoadInitFiles description
''' Set max len for string type cell
''' Set min value for numeric type cell
''' Set max value for numeric type cell
''' Call after load template
''' No parameter
Private Sub LoadInitFiles()
    On Error GoTo ErrorHandle
    Dim i As Long, lCol As Long, lRow As Long
    Dim xmlNodeListIni As MSXML.IXMLDOMNodeList
    Dim xmlNodeIni As MSXML.IXMLDOMNode
    Dim fso As New FileSystemObject
    
    Dim temp As Boolean
    
    For i = 0 To fpSpread1.SheetCount - 2
        ReDim Preserve xmlDocumentInit(i)
        Set xmlDocumentInit(i) = New MSXML.DOMDocument
        If fso.FileExists(GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(i), "InterfaceIni"))) Then
            xmlDocumentInit(i).Load GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(i), "InterfaceIni"))
            Set xmlNodeListIni = xmlDocumentInit(i).getElementsByTagName("Cell")
            For Each xmlNodeIni In xmlNodeListIni
                fpSpread1.sheet = i + 1
                ParserCellID fpSpread1, GetAttribute(xmlNodeIni, "CellID"), lCol, lRow
                fpSpread1.Col = lCol
                fpSpread1.Row = lRow
                If Val(GetAttribute(xmlNodeIni, "MaxLen")) <> 0 Then
                    fpSpread1.TypeMaxEditLen = Val(GetAttribute(xmlNodeIni, "MaxLen"))
                End If
                If fpSpread1.CellType = CellTypeNumber Then
                    If strKHBS = "frmKHBS_BS" Then
                        fpSpread1.TypeNumberMin = Val("-999999999999")
                        fpSpread1.TypeNumberMax = Val(GetAttribute(xmlNodeIni, "MaxValue"))
                    Else
                        fpSpread1.TypeNumberMin = Val(GetAttribute(xmlNodeIni, "MinValue"))
                        fpSpread1.TypeNumberMax = Val(GetAttribute(xmlNodeIni, "MaxValue"))
                    End If
                End If
                fpSpread1.CellTag = GetAttribute(xmlNodeIni, "HelpContextID") & fpSpread1.CellTag
            Next
        End If
    Next
    
    Set fso = Nothing
    Set xmlNodeIni = Nothing
    Set xmlNodeListIni = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "LoadInitFiles", Err.Number, Err.Description
End Sub

''' ResetData description
''' Reset data in active sheet
''' Number type cell -> set to zero
''' String type cell -> set to vbNullString
''' No parameter
Private Sub ResetData()
    On Error GoTo ErrorHandle
    
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim IsUpdate As Boolean
    Dim idtkhai As Variant
    
    Dim totalCell, countCell As Integer
    
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.ResetData
    End If
    fpSpread1.ReDraw = False
    
    ' Bien totalCell nay dung de tinh tong so cell phai clear data, trong truong hop la to khai TNCN thi ko clear mot so chi tieu
    totalCell = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell").length
    countCell = 1
    'c�c tk GTGT,TNDN,TAIN,TTDB,NTNN ko cler phan check tren TK
    'vttoan: them ID (86,87,88) cua cac to (01_BVMT,02BVMT,01_PHXD)
    'dntai : them ID 77 to 02_TAIN
    idtkhai = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
    If (idtkhai = "01" Or idtkhai = "02" Or idtkhai = "04" Or idtkhai = "11" Or idtkhai = "12" Or idtkhai = "06" Or idtkhai = "05" Or idtkhai = "70" Or idtkhai = "72" Or idtkhai = "77" Or idtkhai = "75" Or idtkhai = "74") Then
        For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fpSpread1.sheet = mCurrentSheet
            ParserCellID fpSpread1, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fpSpread1.Col = lCol
            fpSpread1.Row = lRow

            If ((idtkhai = "01" And (lRow < 22 Or lRow > 48)) Or (idtkhai = "02" And (lRow < 38 Or lRow > 54)) Or (idtkhai = "04" And (lRow < 34 Or lRow > 41)) Or (idtkhai = "11" And (lRow < 20 Or lRow > 35)) Or (idtkhai = "12" And (lRow < 34 Or lRow > 49)) Or (idtkhai = "06" And (lRow < 34 Or lRow > 48 + (TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length - 11) / 13)) Or (idtkhai = "05" And (lRow < 31 Or lRow > fpSpread1.MaxRows - 15)) Or (idtkhai = "70" And (lRow < 51 Or lRow > 58 + (TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length - 19) / 14)) Or (idtkhai = "77" And (lRow < 18 Or lRow > fpSpread1.MaxRows - 11)) _
            Or (idtkhai = "75" And (lRow < 38 Or lRow > fpSpread1.MaxRows - 5)) Or (idtkhai = "74" And (lRow < 19 Or lRow > 61)) Or (idtkhai = "72" And (lRow < 43 Or lRow > 48))) And mCurrentSheet = 1 Then

                GoTo nextClear1
            Else
                Select Case fpSpread1.CellType
                    Case CellTypeCheckBox
                        fpSpread1.Text = vbNullString
                        IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                    Case CellTypeComboBox
                        fpSpread1.Text = vbNullString
                        IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                    Case CellTypeNumber
                        fpSpread1.value = 0
                        IsUpdate = UpdateCell(lCol, lRow, "0")
                    Case Else
                        fpSpread1.value = vbNullString
                        IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                End Select
            End If
            'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
            TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_New.AdjustData(mCurrentSheet - 1))
nextClear1:
        Next
    ElseIf (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_10") Or (TAX_Utilities_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11") Then
        For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            ' Doi voi to khai 09/TNCN thi hai chi tieu tu thang den thang cung ko duoc clear
            'Cac TK TNCN ko cler cac chi tieu hearder v� footer
            fpSpread1.sheet = mCurrentSheet
            ParserCellID fpSpread1, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fpSpread1.Col = lCol
            fpSpread1.Row = lRow
            If ((Trim(GetAttribute(TAX_Utilities_New.NodeMenu, "ID")) = "41" And countCell <= 3) Or ((idtkhai = "46" Or idtkhai = "47" Or idtkhai = "48" Or idtkhai = "49") And (lRow < 36 Or lRow > 43)) Or ((idtkhai = "15" Or idtkhai = "16") And (lRow < 38 Or lRow > 57)) Or ((idtkhai = "50" Or idtkhai = "51") And (lRow < 36 Or lRow > 54)) Or ((idtkhai = "36") And (lRow < 36 Or lRow > 63)) Or ((idtkhai = "76") And (lRow < 36)) Or ((idtkhai = "59") And (lRow < 27 Or lRow > 60)) Or idtkhai = "42" Or idtkhai = "43") And mCurrentSheet = 1 Then GoTo nextClear
            Select Case fpSpread1.CellType
                Case CellTypeCheckBox
                    fpSpread1.Text = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                Case CellTypeComboBox
                    fpSpread1.Text = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                Case CellTypeNumber
                    fpSpread1.value = 0
                    IsUpdate = UpdateCell(lCol, lRow, "0")
                Case Else
                    fpSpread1.value = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
            End Select
            'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
            TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_New.AdjustData(mCurrentSheet - 1))
nextClear:
            ' Khi clear du lieu khi bat dau den ngay nhap, nguoi ky, to khai chinh thuc, bo sung, lan bo sung thi thoat khoi vong for luonthen
            'dntai xu ly rieng voi to 09TNCN de khong clear nguoi ky , ngay ky, ho ten nvdl, chung chi so , to khai chinh thuc, bo sung
            If Trim(GetAttribute(TAX_Utilities_New.NodeMenu, "ID")) = "41" Or Trim(GetAttribute(TAX_Utilities_New.NodeMenu, "ID")) = "76" Then
                If countCell = totalCell - 8 And mCurrentSheet = 1 Then
                    Exit For
                End If          'end
            ElseIf Trim(GetAttribute(TAX_Utilities_New.NodeMenu, "ID")) = "17" Then
                If countCell = totalCell - 7 And mCurrentSheet = 1 Then
                    Exit For
                End If          'end
            Else
                If countCell = totalCell - 5 And mCurrentSheet = 1 Then
                    Exit For
                End If
            End If
            countCell = countCell + 1
        Next
    ElseIf idtkhai = "03" Or idtkhai = "80" Or idtkhai = "81" Or idtkhai = "82" Or idtkhai = "73" Or idtkhai = "85" Or idtkhai = "71" Or idtkhai = "86" Or idtkhai = "87" Or idtkhai = "89" Then
    Else
        For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fpSpread1.sheet = mCurrentSheet
            ParserCellID fpSpread1, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fpSpread1.Col = lCol
            fpSpread1.Row = lRow
            Select Case fpSpread1.CellType
                Case CellTypeCheckBox
                    fpSpread1.Text = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                Case CellTypeComboBox
                    fpSpread1.Text = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
                Case CellTypeNumber
                    fpSpread1.value = 0
                    IsUpdate = UpdateCell(lCol, lRow, "0")
                Case Else
                    fpSpread1.value = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
            End Select
            'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
            TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_New.AdjustData(mCurrentSheet - 1))
        Next
    End If
    
    ' Xoa cac dong trong
    With fpSpread1
        .sheet = mCurrentSheet
        If .SheetVisible Then
            If idtkhai = "17" Then
                delNullRowOn05 mCurrentSheet - 1
            ElseIf idtkhai = "59" Then
                delNullRowOn06 mCurrentSheet - 1
            Else
                delNullRow mCurrentSheet - 1
            End If
        End If
    End With
    
    If Not objTaxBusiness Is Nothing Then
        'Set new data to grid
        objTaxBusiness.SetData
    End If
    
    fpSpread1.ReDraw = True
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ResetData", Err.Number, Err.Description
End Sub

'Description: Check if Data on interface has changed
Function IsAdjustData() As Boolean
    Dim i As Long
    IsAdjustData = False
'    For i = 0 To TAX_Utilities_New.AdjustDataCount - 1
'        If TAX_Utilities_New.AdjustData(i) = True Then
'            IsAdjustData = True
'            Exit Function
'        End If
'    Next
'*********************
    For i = 0 To TAX_Utilities_New.AdjustDataCount
        If TAX_Utilities_New.AdjustData(i) = True Then
            IsAdjustData = True
            Exit Function
        End If
    Next
'*********************
End Function

'reset value of all elements in array TAX_Utilities_New.AdjustData to false
'mean Data is not changed
Sub ResetAdjustData()
    Dim i As Long
'    For i = 0 To TAX_Utilities_New.AdjustDataCount - 1
'        TAX_Utilities_New.AdjustData(i) = False
'    Next
    For i = 0 To TAX_Utilities_New.AdjustDataCount
        TAX_Utilities_New.AdjustData(i) = False
    Next
'**********************
End Sub

Sub SetActiveFirstCell(Optional ByRef lSheet As Long, Optional ByRef lCol As Long, Optional ByRef lRow As Long)
Dim iCurrentSheet As Integer
Dim blFirstCell As Boolean
Dim i As Long, j As Long
'Dim lSheet As Long, i As Long, j As Long
'Dim lRow As Long, lCol As Long
With fpSpread1
    .SetFocus
    .sheet = .ActiveSheet
    blFirstCell = False
    If .SheetVisible = True Then
        lSheet = .sheet
        For i = 1 To .MaxRows
            For j = 1 To .MaxCols
                lRow = i
                lCol = j
                GetCellSpan fpSpread1, lCol, lRow
                .Row = lRow
                .Col = lCol
                If .Lock = False And .CellType <> CellTypeCheckBox Then
                    .SetActiveCell .Col, .Row
                    blFirstCell = True
                    Exit For
                End If
            Next
            If blFirstCell = True Then Exit For
        Next
    End If
End With

End Sub

'Set Status msg for active cell
Sub SetStatus(Optional lCol As Long, Optional lRow As Long)
On Error GoTo ErrorHandle
'    Dim lRow As Long, lCol As Long
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim strStatusId As String, strStatusMsg As String
    With fpSpread1
        fpSpread1.sheet = mCurrentSheet
        
        If lCol = -1 And lRow = -1 Then
            lCol = .ActiveCol
            lRow = .ActiveRow
        End If
        
        GetCellSpan fpSpread1, lCol, lRow
        Set xmlNodeCell = TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow))
        
        If xmlNodeCell Is Nothing Then
            Exit Sub
        End If
        
        strStatusId = GetAttribute(xmlNodeCell, "StatusID")
                
        If Trim(strStatusId) = "" Then
            lblStatus.caption = ""
        Else
            Set xmlNodeList = xmlDocumentStatus.getElementsByTagName("St")
            For Each xmlNode In xmlNodeList
                If Trim(GetAttribute(xmlNode, "ID")) = Trim(strStatusId) Then
                    strStatusMsg = GetAttribute(xmlNode, "Msg")
                    Exit For
                End If
            Next
            
            If Trim(strStatusMsg) <> "" Then
                lblStatus.caption = Replace(strStatusMsg, "\n", vbCrLf)
            Else
                lblStatus.caption = ""
            End If
        End If
        
    Set xmlNodeCell = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    End With
    
Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetStatus", Err.Number, Err.Description
End Sub

'********************************
'Description: SetHelpContextId procedure set the HelpContextId to
'   fpSpread1 by HelpContextId of specified cell (stored in CellTag
'   property).
'Input:
'    lCol: col position of cell
'    lRow: row position of cell
'Output:
'Return:
'********************************
Private Sub SetHelpContextId(lCol As Long, lRow As Long)
On Error GoTo ErrHandle
    Dim arrStr() As String, intSheetTemp As Integer
    Dim varCols As Variant, varRows As Variant, varCol As Variant, varRow As Variant
    'Backup sheet
    intSheetTemp = fpSpread1.sheet
    'Turn off Event handler
    fpSpread1.EventEnabled(EventAllEvents) = False
    fpSpread1.sheet = fpSpread1.ActiveSheet
    fpSpread1.GetCellSpan lCol, lRow, varCol, varRow, varCols, varRows
    If CLng(varCol) <> -1 Then
        fpSpread1.Col = CLng(varCol)
    Else
        fpSpread1.Col = lCol
    End If
    
    If CLng(varRow) <> -1 Then
        fpSpread1.Row = CLng(varRow)
    Else
        fpSpread1.Row = lRow
    End If
    
    If fpSpread1.CellTag <> vbNullString Then
        arrStr = Split(fpSpread1.CellTag, "~")
        If arrStr(0) <> vbNullString Then
            fpSpread1.HelpContextID = CLng(arrStr(0))
        Else
            fpSpread1.HelpContextID = 0
        End If
    End If
    'Restore sheet
    fpSpread1.sheet = intSheetTemp
    'Turn on Event handler
    fpSpread1.EventEnabled(EventAllEvents) = True
    Exit Sub
ErrHandle:
    fpSpread1.sheet = intSheetTemp
    fpSpread1.EventEnabled(EventAllEvents) = True
    fpSpread1.HelpContextID = 0
    SaveErrorLog Me.Name, "SetHelpContextId", Err.Number, Err.Description
End Sub

'*********************************
'Description: ClearRow procedure reset content of row.
'Input: Row's reset
'*********************************
'Private Sub ClearRow(lRow As Long)
'    Dim lCol As Long
'
'    With fpSpread1
'        '.EventEnabled(EventAllEvents) = False
'        .Sheet = mCurrentSheet
'        .Row = lRow
'
'        For lCol = 1 To .MaxCols
'            .col = lCol
'            If Not .Lock Then
'                .BackColor = vbWhite
'                .CellNote = ""
'
'                Select Case .CellType
'                    Case CellTypeNumber
'                        .Text = 0
'                    Case CellTypeEdit
'                        .Text = ""
'                    Case CellTypeComboBox
'                        .Text = ""
'                    Case CellTypeDate
'                        .Text = ""
'                    Case CellTypePic
'                        .Text = ""
'                    Case Else
'                        .Text = ""
'                End Select
'            End If
'
'            fpSpread1_Change lCol, lRow
'        Next lCol
'        '.EventEnabled(EventAllEvents) = True
'    End With
'End Sub

Private Sub ClearRows(xmlCellsNode As MSXML.IXMLDOMNode) '(ByVal lRow As Long, ByVal lRows As Long)
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode

    With fpSpread1
        For Each xmlCellNode In xmlCellsNode.childNodes
            ParserCellID fpSpread1, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
            .sheet = mCurrentSheet
            .Col = lCol
            .Row = lRow
            If .Lock = False Or (.Lock = True And .Formula = vbNullString) Then
                Select Case .CellType
                    Case CellTypeNumber
                        .Text = "0"
                        SetAttribute xmlCellNode, "Value", "0"
                    Case Else
                        .Text = ""
                        SetAttribute xmlCellNode, "Value", ""
                End Select
           End If
                
            ' set lai col
            .Col = lCol
            .Row = lRow
            If .CellNote <> vbNullString Then
                .CellNote = vbNullString
                .BackColor = vbWhite
            End If
            '*******************************
        Next
    End With
End Sub

Private Sub ResetErrorCells()
    Dim lCtrl As Long, lNoErrColor As Long
    Dim strCellString As String
    
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.ResetErrorCells
    End If
    
    With fpSpread1
        .ReDraw = False
        .sheet = mHeaderSheet
        
        For lCtrl = 12 To .MaxRows
            .sheet = mHeaderSheet
            .Col = 2
            .Row = lCtrl
            If .Formula <> vbNullString Then
                .Col = .Col + 1
                strCellString = .Formula
                lNoErrColor = .BackColor
                SetCellNote strCellString, lNoErrColor, ""
            End If
        Next
        .ReDraw = True
    End With
End Sub

Private Function ResetDataAndForm(intSheet As Integer)
    Dim xmlSecionNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    
    Set xmlSecionNode = TAX_Utilities_New.Data(intSheet - 1).getElementsByTagName("Section")(0)
    'fpSpread1.Visible = False
    While Not xmlSecionNode Is Nothing
        If GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            While xmlSecionNode.childNodes.length > 1
                Set xmlCellNode = xmlSecionNode.lastChild.selectSingleNode("Cell")
                ParserCellID fpSpread1, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                fpSpread1.sheet = intSheet
                DeleteNode intSheet, lCol, lRow, False
                DoEvents
            Wend
        End If
        Set xmlSecionNode = xmlSecionNode.nextSibling
    Wend
    'fpSpread1.Visible = True
    'TAX_Utilities_New.AdjustData(intSheet - 1) = True
End Function

Private Function IsEmptyValue(xmlCellsNode As MSXML.IXMLDOMNode) As Boolean
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim blnIsEmptyValue As Boolean
    
    blnIsEmptyValue = True
    
    For Each xmlCellNode In xmlCellsNode.childNodes
        ParserCellID fpSpread1, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
        fpSpread1.Col = lCol
        fpSpread1.Row = lRow
        Select Case fpSpread1.CellType
            Case CellTypeNumber, CellTypePercent
                If Not IsNullNumber(GetAttribute(xmlCellNode, "Value")) Then
                    blnIsEmptyValue = False
                    Exit For
                End If
            Case CellTypePic
                If Not IsNullPic(GetAttribute(xmlCellNode, "Value")) Then
                    blnIsEmptyValue = False
                    Exit For
                End If
'            Case CellTypeDate
'                If GetAttribute(xmlCellNode, "Value") <> "" Then
'                    blnIsEmptyValue = False
'                    Exit For
'                End If
            Case Else
                If GetAttribute(xmlCellNode, "Value") <> "" Then
                    blnIsEmptyValue = False
                    Exit For
                End If
        End Select
    Next
    
    IsEmptyValue = blnIsEmptyValue
End Function

Private Function IsNullPic(ByVal strValue As String) As Boolean
    strValue = Replace$(strValue, "/", "")
    strValue = Replace$(strValue, "\", "")
    strValue = Replace$(strValue, ".", "")
    If Trim(strValue) = "" Then IsNullPic = True
End Function

Private Sub LoadKHBS()

    mOnLoad = True
    fpSpread1.EventEnabled(EventAllEvents) = False
    SetControlCaption Me, "frmInterfaces"
    LoadTemplate fpSpread1
    SetupSpread
    
    FormatGrid

    If Trim(GetAttribute(TAX_Utilities_New.NodeValidity, "Class")) <> vbNullString Then
        Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_New.NodeValidity, "Class"))
        Set objTaxBusiness.fps = fpSpread1
        objTaxBusiness.Prepare1
    End If

    LoadStatusFile
    LoadInitFiles
    
    TAX_Utilities_New.AdjustDataReDim fpSpread1.SheetCount - 2
    
    Set objTaxBusiness.fps = Nothing
    fpSpread1.EventEnabled(EventChange) = True
    mOnSetupData = True
    mOnSetupData = False
    
    'SetSheetVisible fpSpread1
    Dim xmlSheetNode As MSXML.IXMLDOMNode
    Dim intCtrl As Integer
    
     With fpSpread1
        For intCtrl = 1 To .SheetCount
            .sheet = intCtrl
            For Each xmlSheetNode In TAX_Utilities_New.NodeValidity.childNodes
                If .SheetName = GetAttribute(xmlSheetNode, "Caption") Then
                    If GetAttribute(xmlSheetNode, "Active") = "0" Then
                        .SheetVisible = False
                    End If
                    Exit For
                End If
            Next
        Next intCtrl
    End With
    
    
    fpSpread1.EventEnabled(EventChange) = False
     
    fpSpread1.sheet = fpSpread1.SheetCount - 1
    fpSpread1.ActiveSheet = fpSpread1.sheet
    fpSpread1.SheetVisible = True
    mCurrentSheet = fpSpread1.ActiveSheet
    FormatGrid
    LoadInitFiles
    Set objTaxBusiness.fps = fpSpread1
    objTaxBusiness.Prepare_KHBS
     
    
    SetupDataKHBS fpSpread1
     
     Set objTaxBusiness.fps = fpSpread1
     '***************
     
     
     If Not objTaxBusiness Is Nothing Then
         If strKHBS = "frmKHBS_BS" Then
            objTaxBusiness.loaiKHBS = "frmKHBS_BS"
         End If
         objTaxBusiness.Prepare2
     End If
    

     With fpSpread1
        For intCtrl = 1 To .SheetCount
            .sheet = intCtrl
            For Each xmlSheetNode In TAX_Utilities_New.NodeValidity.childNodes
                If .SheetName = GetAttribute(xmlSheetNode, "Caption") Then
                    If GetAttribute(xmlSheetNode, "Active") = "0" Then
                        .SheetVisible = False
                    End If
                    Exit For
                End If
            Next
        Next intCtrl
    End With
    fpSpread1.sheet = 1
    fpSpread1.SheetName = GetAttribute(GetMessageCellById("0120"), "Msg")
    fpSpread1.ActiveSheet = fpSpread1.SheetCount - 1
    fpSpread1.EventEnabled(EventAllEvents) = True
    
    mOnLoad = False
    hasActiveForm = True
   
  Set xmlSheetNode = Nothing
   
End Sub
Private Sub LoadKHBS_TT28()
    fpSpread1.EventEnabled(EventAllEvents) = False
    SetupDataKHBS_TT28 fpSpread1
    fpSpread1.EventEnabled(EventAllEvents) = True
   
End Sub


Private Sub saveKHBS()
    Dim strDataFileName As String
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlListNodeCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell1s As MSXML.IXMLDOMNode
    Dim fso As New FileSystemObject
    Dim blnSaveSession As Boolean
    
    On Error GoTo ErrorHandle
        
'       If CheckValidKHBSData = False Then
'        DisplayMessage "0016", msOKOnly, miInformation
'        Exit Sub
'       End If
       If saveLastKHBS = False Then Exit Sub
        
       blnSaveSession = True
       
       'xmlNodeCell1s.Attributes.getNamedItem("DateKHBS").nodeValue = TAX_Utilities_New.DateKHBS
       
       
        If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Day") <> "1" Then
            If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "95" _
            Or GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "71" Then
                If strQuy = "TK_THANG" Then
                    strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                ElseIf strQuy = "TK_QUY" Then
                    strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
                End If
            Else
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
            End If
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
             strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") <> "1" Then
                 'Data file contain Day from and to.
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
                 'Data file contain Day.
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        Else
                 'Data file not contain Day from and to.
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
             '*********************************
        End If
       
        If TAX_Utilities_New.DataChanged And blnSaveSession Then
            If intDataSession >= 999 Then
                intDataSession = 0
            Else
                intDataSession = intDataSession + 1
            End If
            If intPrintingSession >= 999 Then
                intPrintingSession = 0
            Else
                intPrintingSession = intPrintingSession + 1
            End If
            If SaveSessionValueToFile(TAX_Utilities_New.DataFolder & "Session.dat") Then
                TAX_Utilities_New.DataChanged = False
            Else
                Exit Sub
            End If
        End If
    
        
        Set xmlNodeCell1s = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Section")(2)
        Set xmlListNodeCell = TAX_Utilities_New.Data(0).getElementsByTagName("Section")
        Dim xmlNodeNewCells As MSXML.IXMLDOMNode
        For Each xmlNodeCells In xmlListNodeCell
         Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
            If Not xmlNodeCell1s.nextSibling Is Nothing Then
                 xmlNodeCell1s.parentNode.insertBefore xmlNodeNewCells, xmlNodeNewCells.nextSibling
            Else
                xmlNodeCell1s.parentNode.insertBefore xmlNodeNewCells, Null
            End If
        Next
    
'        If Not xmlNodeCell1s.nextSibling Is Nothing Then
'            xmlNodeCell1s.parentNode.insertBefore xmlNodeCells, xmlNodeCell1s.nextSibling
'            'xmlNodeCell1s.removeChild
'        Else
'            xmlNodeCell1s.parentNode.insertBefore xmlNodeCells, Null
'        End If
        TAX_Utilities_New.Data(CLng(TAX_Utilities_New.xmlDataCount)).save strDataFileName
        
        DisplayMessage "0002", msOKOnly, miInformation
        Dim i As Integer
        Set xmlNodeCell1s = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Section")(0)
        For i = 3 To TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Section").length - 1
            xmlNodeCell1s.parentNode.removeChild TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Section")(3)
        Next
        ResetAdjustData
         
   '  End If
         Exit Sub
   
ErrorHandle:
    SaveErrorLog Me.Name, "SaveKHBS", Err.Number, Err.Description
End Sub



Private Function CheckValidKHBSData() As Boolean
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeC As MSXML.IXMLDOMNode
    Dim xmlNodeH As MSXML.IXMLDOMNode
    Dim xmlListNodeCell As MSXML.IXMLDOMNodeList
    Dim xmlListNodeCellKHBS As MSXML.IXMLDOMNodeList
    Dim strCellID() As String
    Dim strCellID1 As String
    Dim strValue As String
    Dim strValueCheck As String
    Dim lCol As Long, lRow As Long
    
    CheckValidKHBSData = True
    
    
    Set xmlListNodeCellKHBS = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")
    Set xmlListNodeCell = TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
    For Each xmlNodeCells In xmlListNodeCellKHBS
        strCellID = Split(GetAttribute(xmlNodeCells, "CellID"), "_")
        If strCellID(0) = "BC" Then
                Set xmlNodeC = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).nodeFromID("BG_" & strCellID(1))
                Set xmlNodeH = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).nodeFromID("BH_" & strCellID(1))
                strValue = CDbl(GetAttribute(xmlNodeH, "Value"))
            strCellID1 = Trim(Mid(GetAttribute(xmlNodeCells, "Value"), 100, 20))
                For Each xmlNode In xmlListNodeCell
                    If GetAttribute(xmlNode, "CellID") = strCellID1 Then
                        fpSpread1.sheet = 1
                        ParserCellID fpSpread1, GetAttribute(xmlNode, "CellID"), lCol, lRow
                        fpSpread1.Col = lCol
                        fpSpread1.Row = lRow
                        strValueCheck = fpSpread1.value
                        If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Then
                            If GetAttribute(xmlNode, "CellID") = "L_17" Or GetAttribute(xmlNode, "CellID") = "L_7" Then
                                    strValueCheck = -strValueCheck
                            End If
                        End If
                        If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Then
                            If GetAttribute(xmlNode, "CellID") = "L_11" Or GetAttribute(xmlNode, "CellID") = "L_12" Then
                                    strValueCheck = -strValueCheck
                            End If
                        End If
                        
                        If strValue <> strValueCheck Then
                            CheckValidKHBSData = False
                            If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "01" Then
                                If GetAttribute(xmlNode, "CellID") = "L_17" Or GetAttribute(xmlNode, "CellID") = "L_7" Then
                                        fpSpread1.value = -fpSpread1.value
                                End If
                            End If
                            
                            If GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID") = "02" Then
                                If GetAttribute(xmlNode, "CellID") = "L_18" Or GetAttribute(xmlNode, "CellID") = "L_19" Then
                                        strValueCheck = -strValueCheck
                                End If
                            End If
                            With fpSpread1
                                .sheet = 1
                                ParserCellID fpSpread1, GetAttribute(xmlNode, "CellID"), lCol, lRow
                                .Col = lCol
                                .Row = lRow
                                .BackColor = &HC0C0FF
                                .CellNote = GetAttribute(GetMessageCellById("0108"), "Msg")
                                
                                .sheet = .SheetCount - 1
                                ParserCellID fpSpread1, GetAttribute(xmlNodeC, "CellID"), lCol, lRow
                                .Col = lCol
                                .Row = lRow
                                .BackColor = &HC0C0FF
                                .ActiveSheet = .SheetCount - 1
                                .CellNote = GetAttribute(GetMessageCellById("0109"), "Msg") & "[" & Trim(Right(GetAttribute(xmlNodeCells, "Value"), 10)) & "]"
                                 
                             End With
                        Else
                             With fpSpread1
                                .sheet = 1
                                ParserCellID fpSpread1, GetAttribute(xmlNode, "CellID"), lCol, lRow
                                .Col = lCol
                                .Row = lRow
                                .BackColor = vbWhite
                                .CellNote = ""
                              
                                .sheet = .SheetCount - 1
                                ParserCellID fpSpread1, GetAttribute(xmlNodeC, "CellID"), lCol, lRow
                                .Col = lCol
                                .Row = lRow
                                .BackColor = vbWhite
                                .CellNote = ""
                             End With
                        End If
                        Exit For
                    End If
                Next
        End If
               
    Next
    
    Set xmlNode = Nothing
    Set xmlNodeC = Nothing
    Set xmlNodeCells = Nothing
    Set xmlListNodeCell = Nothing
    Set xmlListNodeCellKHBS = Nothing
    
End Function



Public Function delNullRow(sheet As Long)
    On Error GoTo ErrorHandle
    Dim xmlNodeListSec As MSXML.IXMLDOMNodeList
    Dim xmlNodeListRow As MSXML.IXMLDOMNodeList
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeSec As MSXML.IXMLDOMNode
    Dim xmlNodeRow As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim numSec, Row, row1, celllg, hasVl As Long
    Dim sumRowDel, countDel As Long
    
    Dim cellid, value As Variant
    Dim OldSheet As Long
    
    Dim maxRow As Long
    
    sumRowDel = TAX_Utilities_New.Data(sheet).getElementsByTagName("Cell").length
    
    OldSheet = fpSpread1.ActiveSheet
    ' Xem lai vi sao lai countDel <> 19
    ' 09112011
    fpSpread1.sheet = sheet + 1
    maxRow = fpSpread1.MaxRows
    'Do While countDel <> 19
    Do While countDel <> maxRow
        countDel = countDel + 1
        Set xmlNodeListSec = TAX_Utilities_New.Data(sheet).getElementsByTagName("Section")
'sec
        numSec = 0
        For Each xmlNodeSec In xmlNodeListSec
            If GetAttribute(xmlNodeSec, "Dynamic") = "1" Then
                Set xmlNodeListRow = xmlNodeListSec.Item(numSec).childNodes
        'row
                Row = 0
                For Each xmlNodeRow In xmlNodeListRow
                    hasVl = 0
                    Set xmlNodeListCell = xmlNodeListRow.Item(Row).childNodes
               'cell
                    For Each xmlNodeCell In xmlNodeListCell
                        value = GetAttribute(xmlNodeCell, "Value")
                        'If GetAttribute(xmlNodeCell, "FirstCell") = "" And value <> "" And value <> "0" And value <> "cbo" And value <> "0%" And value <> "5%" And value <> "10%" Then
                        If (GetAttribute(xmlNodeCell, "FirstCell") <> "" And value <> "") Or (GetAttribute(xmlNodeCell, "FirstCell") = "" And value <> "" And value <> "0" And value <> "cbo" And value <> "0%" And value <> "5%" And value <> "10%") Then
                            hasVl = hasVl + 1
                        End If
                        cellid = GetAttribute(xmlNodeCell, "CellID")
                    Next
                    If hasVl = 0 Then
                        If Mid(cellid, 2, 1) = "_" Then
                            fpSpread1.ActiveSheet = sheet + 1
                            DeleteNode sheet + 1, fpSpread1.ColLetterToNumber(Left(cellid, 1)), CLng(Right(cellid, Len(cellid) - 2)), True
                             Exit For
                        ElseIf Mid(cellid, 3, 1) = "_" Then
                            fpSpread1.ActiveSheet = sheet + 1
                            DeleteNode sheet + 1, fpSpread1.ColLetterToNumber(Left(cellid, 2)), CLng(Right(cellid, Len(cellid) - 3)), True
                            Exit For
                        Else
                            
                        End If
                    End If
                    Row = Row + 1
                Next
            End If
            numSec = numSec + 1
        Next
    Loop
    fpSpread1.ActiveSheet = OldSheet
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "delNullRow", Err.Number, Err.Description
End Function


Private Function saveLastKHBS() As Boolean
    Dim strDataFileName As String
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlListNodeCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell1s As MSXML.IXMLDOMNode
    Dim fso As New FileSystemObject

     saveLastKHBS = False
     If GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Day") <> "1" Then
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS1_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ThreeMonth") = "1" Then
             strDataFileName = TAX_Utilities_New.DataFolder & "KHBS1_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_New.ThreeMonths & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") <> "1" Then
                 'Data file contain Day from and to.
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS1_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_New.Year & "_" & Replace(TAX_Utilities_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_New.LastDay, "/", "") & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_New.NodeMenu, "Month") = "1" Then
                 'Data file contain Day.
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS1_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_New.Day & TAX_Utilities_New.month & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
        Else
                 'Data file not contain Day from and to.
                 strDataFileName = TAX_Utilities_New.DataFolder & "KHBS1_" & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_New.Year & "_" & TAX_Utilities_New.DateKHBS & ".xml"
             '*********************************
        End If
        
        TAX_Utilities_New.DataKHBS.save strDataFileName
        saveLastKHBS = True

End Function

' Ham check validate cau truc cua to khai
Public Function checkCauTrucData() As Boolean
    Dim result As Boolean
        ' Phuc vu check cau truc to khai
    Dim strCauTruc() As String
    Dim strChiTieu() As String
    Dim strTkhaiId As String
    Dim idx As Integer, i As Integer, j As Integer, currRow As Double, contDynamicRow As Integer
    Dim strSection As String
    Dim soCTTemp As Integer
    Dim soCTData As Integer
    Dim soSectionTemp As Integer
    Dim soSectionData As Integer
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim xmlNodeCellID As MSXML.IXMLDOMNode
    Dim xmlListNodeSection As MSXML.IXMLDOMNodeList
    Dim strKyHieuCT As String, strKyHieuCTTemp As String
    Dim strFirstRow As String
    Dim stepNumRow As Integer
    ' end
        ' Check so chi tieu cua to khai
        ' tam thoi coment de test
    Set xmlListNodeSection = TAX_Utilities_New.Data(0).getElementsByTagName("Section")
    strTkhaiId = GetAttribute(TAX_Utilities_New.NodeMenu, "ID")
    strCauTruc = getTemplateTk(strTkhaiId)
    ' Kiem tra neu tra ve null la to khai chua duoc dinh nghia cau truc thu tu cac chi tieu
    If strCauTruc(0) = "null" Then
        checkCauTrucData = True
        Exit Function
    End If
    soSectionTemp = UBound(strCauTruc)
    soSectionData = xmlListNodeSection.length
    ' Kiem tra neu khac so section thi bao loi
    If soSectionData <> soSectionTemp Then
        checkCauTrucData = False
        checkSoCT = 0 ' Khac so section
        Exit Function
    End If
    ' Kiem tra so chi tieu trong moi section va thu tu co chuan theo cau truc khong?
    contDynamicRow = 0
    For idx = 0 To UBound(strCauTruc) - 1
        strSection = strCauTruc(idx)
        Set xmlNodeCells = xmlListNodeSection.Item(idx)
        ' Kiem tra so luong chi tieu tren 1 section
        strChiTieu = Split(strSection, "~")
        soCTTemp = UBound(strChiTieu)
        ' Dynamic =0
        If Right$(strChiTieu(soCTTemp), 1) = "0" Then
            ' khong kiem tra section cuoi cua tk TNDN
            If strTkhaiId = "11" Or strTkhaiId = "12" Or strTkhaiId = "03" Then
                If idx = UBound(strCauTruc) - 1 Then
                    checkCauTrucData = True
                    Exit Function
                End If
            End If
        
            ' to khai 01/GTGT
            If strTkhaiId = "01" Then
                ' neu session cuoi in ra tu ban 3.1.3 se nhieu hon in tu ban 3.1.2 2 chi tieu
                soCTData = GetElementsNoData(xmlNodeCells.childNodes(0))
                If idx = 2 Then
                    If soCTTemp > soCTData And soCTTemp - soCTData <> 3 And soCTTemp - soCTData <> 1 Then
                        checkCauTrucData = False
                        checkSoCT = 1 ' Thieu chi tieu
                        Exit Function
                    End If
                    If soCTTemp < soCTData Then
                        checkCauTrucData = False
                        checkSoCT = 2 ' Thua chi tieu
                        Exit Function
                    End If
                    ' Kiem tra sai vi tri cac chi tieu tren interface template
                    For i = 0 To soCTTemp - 1
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)
                        Set xmlNodeCellID = xmlNodeCell.childNodes(i)
                        ' chi tieu kiem tra gia han thue se khong kiem tra
                        If i = 8 Then
                            Exit For
                        End If
                        strKyHieuCT = GetAttribute(xmlNodeCellID, "CellID")
                        
                        strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                                
                        If strKyHieuCTTemp <> strKyHieuCT Then
                            checkCauTrucData = False
                            checkSoCT = 4 ' Sai vi tri chi tieu
                            Exit Function
                        End If
                    Next i
                    
                Else
                    If soCTTemp > soCTData Then
                        checkCauTrucData = False
                        checkSoCT = 1 ' Thieu chi tieu
                        Exit Function
                    End If
                    If soCTTemp < soCTData Then
                        checkCauTrucData = False
                        checkSoCT = 2 ' Thua chi tieu
                        Exit Function
                    End If
                    ' Kiem tra sai vi tri cac chi tieu tren interface template
                    For i = 0 To soCTTemp - 1
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)
                        Set xmlNodeCellID = xmlNodeCell.childNodes(i)
                        strKyHieuCT = GetAttribute(xmlNodeCellID, "CellID")
                        If strTkhaiId <> "03" And strTkhaiId <> "70" And strTkhaiId <> "81" And strTkhaiId <> "71" And strTkhaiId <> "77" And strTkhaiId <> "87" And strTkhaiId <> "76" And strTkhaiId <> "06" And strTkhaiId <> "05" Then
                            strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                        Else
                            ' To khai 03/TNDN
                            If strTkhaiId = "03" Then
                                If idx = 5 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            ' To khai 08B/TNDN , 01_TAIN
                            ElseIf strTkhaiId = "76" Or strTkhaiId = "06" Then
                                If idx = 3 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            ' To khai 01/TTDB
                            ElseIf strTkhaiId = "05" Then
                                If idx = 10 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            Else
                                ' Du lieu cua section tinh trong cung tk voi du lieu dong
                                strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                            End If
                        End If
                        
                        If strKyHieuCTTemp <> strKyHieuCT Then
                            checkCauTrucData = False
                            checkSoCT = 4 ' Sai vi tri chi tieu
                            Exit Function
                        End If
                    Next i
                End If
            ElseIf strTkhaiId = "02" Or strTkhaiId = "04" Or strTkhaiId = "71" Then
            ' To khai 02,03,04/GTGT 02/TNDN
                soCTData = GetElementsNoData(xmlNodeCells.childNodes(0))
                If idx = UBound(strCauTruc) - 1 Then
                     If soCTTemp > soCTData And soCTTemp - soCTData <> 1 Then
                        checkCauTrucData = False
                        checkSoCT = 1 ' Thieu chi tieu
                        Exit Function
                    End If
                    If soCTTemp < soCTData Then
                        checkCauTrucData = False
                        checkSoCT = 2 ' Thua chi tieu
                        Exit Function
                    End If
                    ' Kiem tra sai vi tri cac chi tieu tren interface template
                    For i = 0 To soCTTemp - 1
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)
                        Set xmlNodeCellID = xmlNodeCell.childNodes(i)
                        ' chi tieu kiem tra gia han thue se khong kiem tra
                        If i = soCTTemp - 1 Then
                            Exit For
                        End If
                        strKyHieuCT = GetAttribute(xmlNodeCellID, "CellID")
                        
                        strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                                
                        If strKyHieuCTTemp <> strKyHieuCT Then
                            checkCauTrucData = False
                            checkSoCT = 4 ' Sai vi tri chi tieu
                            Exit Function
                        End If
                    Next i
                Else
                    If soCTTemp > soCTData Then
                        checkCauTrucData = False
                        checkSoCT = 1 ' Thieu chi tieu
                        Exit Function
                    End If
                    If soCTTemp < soCTData Then
                        checkCauTrucData = False
                        checkSoCT = 2 ' Thua chi tieu
                        Exit Function
                    End If
                    ' Kiem tra sai vi tri cac chi tieu tren interface template
                    For i = 0 To soCTTemp - 1
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)
                        Set xmlNodeCellID = xmlNodeCell.childNodes(i)
                        strKyHieuCT = GetAttribute(xmlNodeCellID, "CellID")
                        If strTkhaiId <> "03" And strTkhaiId <> "70" And strTkhaiId <> "81" And strTkhaiId <> "71" And strTkhaiId <> "77" And strTkhaiId <> "87" And strTkhaiId <> "76" And strTkhaiId <> "06" And strTkhaiId <> "05" Then
                            strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                        Else
                            ' To khai 03/TNDN
                            If strTkhaiId = "03" Then
                                If idx = 5 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            ' To khai 08B/TNDN , 01_TAIN
                            ElseIf strTkhaiId = "76" Or strTkhaiId = "06" Then
                                If idx = 3 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            ' To khai 01/TTDB
                            ElseIf strTkhaiId = "05" Then
                                If idx = 10 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            Else
                                ' Du lieu cua section tinh trong cung tk voi du lieu dong
                                strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                            End If
                        End If
                        
                        If strKyHieuCTTemp <> strKyHieuCT Then
                            checkCauTrucData = False
                            checkSoCT = 4 ' Sai vi tri chi tieu
                            Exit Function
                        End If
                    Next i
                End If
            Else
                soCTData = GetElementsNoData(xmlNodeCells.childNodes(0))
                If soCTTemp > soCTData Then
                    checkCauTrucData = False
                    checkSoCT = 1 ' Thieu chi tieu
                    Exit Function
                End If
                If soCTTemp < soCTData Then
                    checkCauTrucData = False
                    checkSoCT = 2 ' Thua chi tieu
                    Exit Function
                End If
                    ' Kiem tra sai vi tri cac chi tieu tren interface template
                    For i = 0 To soCTTemp - 1
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)
                        Set xmlNodeCellID = xmlNodeCell.childNodes(i)
                        strKyHieuCT = GetAttribute(xmlNodeCellID, "CellID")
                        If strTkhaiId <> "03" And strTkhaiId <> "70" And strTkhaiId <> "81" And strTkhaiId <> "71" And strTkhaiId <> "77" And strTkhaiId <> "87" And strTkhaiId <> "76" And strTkhaiId <> "06" And strTkhaiId <> "05" Then
                            strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                        Else
                            ' To khai 03/TNDN
                            If strTkhaiId = "03" Then
                                If idx = 5 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            ' To khai 08B/TNDN , 01_TAIN
                            ElseIf strTkhaiId = "76" Or strTkhaiId = "06" Then
                                If idx = 3 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            ' To khai 01/TTDB
                            ElseIf strTkhaiId = "05" Then
                                If idx = 10 Then
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                                Else
                                    strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & contDynamicRow + Val(Split(strChiTieu(i), "_")(1))
                                End If
                            Else
                                ' Du lieu cua section tinh trong cung tk voi du lieu dong
                                strKyHieuCTTemp = Split(strChiTieu(i), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)  ' Lay ky hieu cua temp + row cua du lieu
                            End If
                        End If
                        
                        If strKyHieuCTTemp <> strKyHieuCT Then
                            checkCauTrucData = False
                            checkSoCT = 4 ' Sai vi tri chi tieu
                            Exit Function
                        End If
                    Next i
            End If
        ' Dynamic =1
        Else
            soCTData = 0
            For i = 0 To xmlNodeCells.childNodes.length - 1
                soCTData = soCTData + GetElementsNoData(xmlNodeCells.childNodes(i))
            Next i

            If soCTData Mod soCTTemp <> 0 Then
                checkCauTrucData = False
                checkSoCT = 3 ' khac chi tieu
                Exit Function
            End If
            ' Kiem tra vi tri cua cac chi tieu
            currRow = Val(Split(strChiTieu(0), "_")(1))
            For i = 0 To xmlNodeCells.childNodes.length - 1
               Set xmlNodeCell = xmlNodeCells.childNodes(i)
               ' Kiem tra xem co phai la dong dau tien cua section dynamic khong
               Set xmlNodeCellID = xmlNodeCell.childNodes(0)
               strFirstRow = GetAttribute(xmlNodeCellID, "FirstCell")
               If strFirstRow = "0" Then
                    stepNumRow = contDynamicRow
               Else
                    stepNumRow = 0
               End If
               ' end kiem tra
               For j = 0 To soCTTemp - 1
                    Set xmlNodeCellID = xmlNodeCell.childNodes(j)
                    strKyHieuCT = GetAttribute(xmlNodeCellID, "CellID")
                    'strKyHieuCTTemp = Split(strChiTieu(j), "_")(0) & "_" & currRow + stepNumRow
                    strKyHieuCTTemp = Split(strChiTieu(j), "_")(0) & "_" & Split(strKyHieuCT, "_")(1)   ' Lay cell trong tep mau ghep voi row du lieu -> temp tuong ung
                    If strKyHieuCTTemp <> strKyHieuCT Then
                        checkCauTrucData = False
                        checkSoCT = 4 ' Sai vi tri chi tieu
                        Exit Function
                    End If
               Next j
               currRow = currRow + 1
               contDynamicRow = contDynamicRow + 1
            Next i
            contDynamicRow = contDynamicRow - 1
        End If



    Next idx
    ' End check
    checkCauTrucData = True
End Function
Public Sub UpdateDataKHBS_TT28(pGrid As fpSpread)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lRow As Long
    Dim blnNewData As Boolean, blnHasSetActiveCell As Boolean
    Dim i As Variant
    Dim strKHBSDataFileName As String
    Dim strDataFileName As String
    Dim strOriginDataFileName As String
    Dim varTemp As Variant
                
                'SetAttribute TAX_Utilities_New.NodeValidity.childNodes(lSheet), "Active", "1"
                
                With pGrid
                .sheet = .SheetCount - 1
                i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "bb"
                '------------------------------------------
                
                    
                    .Col = .ColLetterToNumber("B")
                    .Row = 24 + i - 7
                     UpdateCell .ColLetterToNumber("B"), .Row, .Text
                    .Col = .ColLetterToNumber("BE")
                    .Row = 17 + i - 7
                     UpdateCell .ColLetterToNumber("BE"), .Row, .value
                     .Row = 18 + i - 7
                     UpdateCell .ColLetterToNumber("BE"), .Row, IIf(Trim(.value) = "", 0, .value)
                     .Col = .ColLetterToNumber("BD")
                    .Row = 20 + i - 7
                     UpdateCell .ColLetterToNumber("BD"), .Row, .Text
                    .Col = .ColLetterToNumber("BG")
                    .Row = 22 + i - 7
                    UpdateCell .ColLetterToNumber("BG"), .Row, .Text
                    .Col = .ColLetterToNumber("BG")
                    .Row = 23 + i - 7
                    UpdateCell .ColLetterToNumber("BG"), .Row, .Text
                     .Col = .ColLetterToNumber("BF")
                     .Row = 15 + i - 7
                    UpdateCell .ColLetterToNumber("BF"), .Row, IIf(Trim(.value) = "", 0, .value)
                      .Col = .ColLetterToNumber("BG")
                     .Row = 15 + i - 7
                    UpdateCell .ColLetterToNumber("BG"), .Row, IIf(Trim(.value) = "", 0, .value)
                      .Col = .ColLetterToNumber("BH")
                     .Row = 15 + i - 7
                    UpdateCell .ColLetterToNumber("BH"), .Row, IIf(Trim(.value) = "", 0, .value)
                     .Col = .ColLetterToNumber("BF")
                     .Row = 16 + i - 7
                    UpdateCell .ColLetterToNumber("BF"), .Row, IIf(Trim(.value) = "", 0, .value)
                      .Col = .ColLetterToNumber("BG")
                     .Row = 16 + i - 7
                    UpdateCell .ColLetterToNumber("BG"), .Row, IIf(Trim(.value) = "", 0, .value)
                      .Col = .ColLetterToNumber("BH")
                     .Row = 16 + i - 7
                    UpdateCell .ColLetterToNumber("BH"), .Row, IIf(Trim(.value) = "", 0, .value)
                End With
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "UpdateDataKHBS", Err.Number, Err.Description
End Sub
Private Sub TonghopKHBS()
    Dim strTemp As String
    Dim strOldValue As String
    Dim strDieuChinhTangGiam() As String
    Dim arrDieuChinhGiam() As String
    Dim arrDieuChinhTang() As String
    Dim arrDieuChinh4043() As String
    Dim arrValue() As String ' Luu cac cell cua mot row
    Dim numRowI, numRowII, numRowIII, j As Integer
    Dim tempCurrSheet As Integer
    
    Dim flagTang, flagGiam, flag4043 As Boolean
    
    Dim strTongOld, strTongCurr As String ' Luu gia tri tong dieu chinh
    
    Dim countDel As Long
    numRowI = 0
    numRowII = 0
    numRowIII = 0
    
    'Set Fomura cell ngay NC va PNC = "" de ko tinh lai cac gia tri nay
    Dim lCol_temp As Long
    Dim lRow_temp As Long
    Dim temp As Long
    Dim xmlNodeCell_temp As MSXML.IXMLDOMNode
    
    Dim strFormula As String
    Dim vSoTien As Variant
    
    If isNewdataBS = False Then
        If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
                Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11)
                ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                fpSpread1.sheet = fpSpread1.SheetCount - 1
                fpSpread1.Col = lCol_temp
                fpSpread1.Row = lRow_temp
                fpSpread1.Formula = ""
                
                Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10)
                ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                fpSpread1.Col = lCol_temp
                fpSpread1.Row = lRow_temp
                fpSpread1.Formula = ""
            Else
                Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
                ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                fpSpread1.sheet = fpSpread1.SheetCount - 1
                fpSpread1.Col = lCol_temp
                fpSpread1.Row = lRow_temp
                fpSpread1.Formula = ""
                
                
                Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
                ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                fpSpread1.Col = lCol_temp
                fpSpread1.Row = lRow_temp
                fpSpread1.Formula = ""
        End If
    End If
'------------------------------------------------------
    
    If Trim(GetAttribute(TAX_Utilities_New.NodeValidity, "Class")) <> vbNullString Then
        ' Neu chua co object moi tao lai
        If objTaxBusiness Is Nothing Then
            Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_New.NodeValidity, "Class"))
        End If
        
        Set objTaxBusiness.fps = fpSpread1
        strOldValue = objTaxBusiness.getValueTK(strDataFileBS)
        strTemp = objTaxBusiness.getDieuChinhGiam(strOldValue)
        
        'Lay ve gia tri tong
        strTongOld = objTaxBusiness.getValueCTDC(strDataFileBS)
        strTongCurr = objTaxBusiness.getChiTieuTongDC(CStr(strTongOld))
        'end
        
        strDieuChinhTangGiam = Split(strTemp, "###")
        If strDieuChinhTangGiam(0) <> "" Then
            arrDieuChinhGiam = Split(strDieuChinhTangGiam(0), "~")
            numRowII = UBound(arrDieuChinhGiam)
            flagGiam = True
        End If
        If strDieuChinhTangGiam(1) <> "" Then
            arrDieuChinhTang = Split(strDieuChinhTangGiam(1), "~")
            numRowI = UBound(arrDieuChinhTang)
            flagTang = True
        End If
        If GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01GTGT" Then
                If strDieuChinhTangGiam(2) <> "" Then
                    arrDieuChinh4043 = Split(strDieuChinhTangGiam(2), "~")
                    numRowIII = UBound(arrDieuChinh4043)
                    flag4043 = True
                End If
                ' A. Dieu chinh so thue CT 40 43
                fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
                fpSpread1.EventEnabled(EventAllEvents) = False
                tempCurrSheet = mCurrentSheet
                mCurrentSheet = fpSpread1.SheetCount - 1
                fpSpread1.sheet = mCurrentSheet
                ' them so dong dieu chinh thay doi vao
                ' set cac gia tri cua cot
                If flag4043 = True Then
                    For j = 0 To numRowIII
                        
                        arrValue = Split(arrDieuChinh4043(j), "_")
                        If arrValue(4) <> 0 Then
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BJ"), 5 + j, Round(Val(arrValue(2)), 0)
                            UpdateCell fpSpread1.ColLetterToNumber("BJ"), 5 + j, Round(Val(arrValue(2)), 0)
                            'UpdateCell fpSpread1.ColLetterToNumber("BF"), 15 + j, arrValue(2)
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BK"), 5 + j, Round(Val(arrValue(3)), 0)
                            UpdateCell fpSpread1.ColLetterToNumber("BK"), 5 + j, Round(Val(arrValue(3)), 0)
                            'UpdateCell fpSpread1.ColLetterToNumber("BG"), 15 + j, arrValue(3)
                        Else
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BJ"), 5 + j, "0"
                            UpdateCell fpSpread1.ColLetterToNumber("BJ"), 5 + j, "0"
                            'UpdateCell fpSpread1.ColLetterToNumber("BF"), 15 + j, arrValue(2)
                            fpSpread1.SetText fpSpread1.ColLetterToNumber("BK"), 5 + j, "0"
                            UpdateCell fpSpread1.ColLetterToNumber("BK"), 5 + j, "0"
                            'UpdateCell fpSpread1.ColLetterToNumber("BG"), 15 + j, arrValue(3)
                            
                        End If
                        fpSpread1.SetText fpSpread1.ColLetterToNumber("BL"), 5 + j, Round(Val(arrValue(4)), 0)
                        UpdateCell fpSpread1.ColLetterToNumber("BL"), 5 + j, Round(Val(arrValue(4)), 0)
                        'UpdateCell fpSpread1.ColLetterToNumber("BH"), 15 + j, arrValue(4)
                    Next j
                End If
        End If
        ' set gia tri tong 32 cho to khai 02_GTGT
        If GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02GTGT" Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            tempCurrSheet = mCurrentSheet
            mCurrentSheet = fpSpread1.SheetCount - 1
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.SetText fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            UpdateCell fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            mCurrentSheet = tempCurrSheet
            fpSpread1.EventEnabled(EventAllEvents) = True
        End If
        ' Set gia tri tong 34 cho to khai 03_GTGT
        If GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03GTGT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01ATNDN" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01BTNDN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01TTDB" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01TAIN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02TAIN" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03TNDN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_05GTGT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01BVMT" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02BVMT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01PHXD" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02TNDN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_02NTNN" Or _
        GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_04NTNN" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03NTNN" _
        Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01TD_GTGT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_03_TD_TAIN" _
        Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_04GTGT" Or GetAttribute(TAX_Utilities_New.NodeValidity, "Class") = "TAX_Business.cls_01NTNN" Then
            fpSpread1.EventEnabled(EventAllEvents) = False
            tempCurrSheet = mCurrentSheet
            mCurrentSheet = fpSpread1.SheetCount - 1
            fpSpread1.sheet = mCurrentSheet
            fpSpread1.SetText fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            UpdateCell fpSpread1.ColLetterToNumber("BI"), 5, Round(Val(strTongCurr), 0)
            mCurrentSheet = tempCurrSheet
            fpSpread1.EventEnabled(EventAllEvents) = True
        End If
        
        ' I. Dieu chinh tang so thue
        fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
        fpSpread1.EventEnabled(EventAllEvents) = False
        tempCurrSheet = mCurrentSheet
        mCurrentSheet = fpSpread1.SheetCount - 1
        ' xoa dong cu truoc khi them dong
        fpSpread1.Row = 9
        fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
'        fpSpread1.EventEnabled(EventAllEvents) = False
        fpSpread1.sheet = mCurrentSheet
        Do
            countDel = countDel + 1
            fpSpread1.Row = fpSpread1.Row + 1
        Loop Until UCase(fpSpread1.Text) = "AA"
        
        fpSpread1.EventEnabled(EventAllEvents) = False
        For j = 0 To countDel - 1
            DeleteNode mCurrentSheet, fpSpread1.ColLetterToNumber("BD"), 9, False
        Next j
        ' them so dong dieu chinh thay doi vao
        For j = 0 To numRowI - 1
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            InsertNode fpSpread1.ColLetterToNumber("BD"), 9
        Next j
        ' set cac gia tri cua cot
        If flagTang = True Then
            For j = 0 To numRowI
                
                arrValue = Split(arrDieuChinhTang(j), "_")
                fpSpread1.SetText fpSpread1.ColLetterToNumber("B"), 9 + j, j + 1
                
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BE"), 9 + j, arrValue(0)
                UpdateCell fpSpread1.ColLetterToNumber("BE"), 9 + j, arrValue(0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BD"), 9 + j, arrValue(1)
                UpdateCell fpSpread1.ColLetterToNumber("BD"), 9 + j, arrValue(1)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BF"), 9 + j, Round(Val(arrValue(2)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BF"), 9 + j, Round(Val(arrValue(2)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BG"), 9 + j, Round(Val(arrValue(3)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BG"), 9 + j, Round(Val(arrValue(3)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BH"), 9 + j, Round(Val(arrValue(4)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BH"), 9 + j, Round(Val(arrValue(4)), 0)
            Next j
        End If
        
        ' II. Dieu chinh giam so thue
        fpSpread1_Change fpSpread1.ActiveCol, fpSpread1.ActiveRow
        fpSpread1.EventEnabled(EventAllEvents) = False
        tempCurrSheet = mCurrentSheet
        mCurrentSheet = fpSpread1.SheetCount - 1
        ' xoa dong cu truoc khi them dong
        fpSpread1.Row = 13 + numRowI
        fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
'        fpSpread1.EventEnabled(EventAllEvents) = False
        fpSpread1.sheet = mCurrentSheet
        Do
            countDel = countDel + 1
            fpSpread1.Row = fpSpread1.Row + 1
        Loop Until UCase(fpSpread1.Text) = "BB"
        
        fpSpread1.EventEnabled(EventAllEvents) = False
        For j = 0 To countDel - 1
            DeleteNode mCurrentSheet, fpSpread1.ColLetterToNumber("BD"), 13 + numRowI, False
        Next j
        ' them so dong dieu chinh thay doi vao
        For j = 0 To numRowII - 1
            fpSpread1.EventEnabled(EventAllEvents) = False
            fpSpread1.sheet = mCurrentSheet
            InsertNode fpSpread1.ColLetterToNumber("BD"), 13 + numRowI
        Next j
        ' set cac gia tri cua cot
        If flagGiam = True Then
            For j = 0 To numRowII
                arrValue = Split(arrDieuChinhGiam(j), "_")
                fpSpread1.SetText fpSpread1.ColLetterToNumber("B"), 13 + numRowI + j, j + 1
                
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BE"), 13 + numRowI + j, arrValue(0)
                UpdateCell fpSpread1.ColLetterToNumber("BE"), 13 + numRowI + j, arrValue(0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BD"), 13 + numRowI + j, arrValue(1)
                UpdateCell fpSpread1.ColLetterToNumber("BD"), 13 + numRowI + j, arrValue(1)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BF"), 13 + numRowI + j, Round(Val(arrValue(2)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BF"), 13 + numRowI + j, Round(Val(arrValue(2)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BG"), 13 + numRowI + j, Round(Val(arrValue(3)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BG"), 13 + numRowI + j, Round(Val(arrValue(3)), 0)
                fpSpread1.SetText fpSpread1.ColLetterToNumber("BH"), 13 + numRowI + j, Round(Val(arrValue(4)), 0)
                UpdateCell fpSpread1.ColLetterToNumber("BH"), 13 + numRowI + j, Round(Val(arrValue(4)), 0)
            Next j
        End If

        mCurrentSheet = tempCurrSheet
        UpdateDataKHBS_TT28 fpSpread1
        'set lai cong thuc cua cac cell NNC va PNC
        If isNewdataBS = False Then
            If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
    
                    fpSpread1.Formula = "BD5"
                    fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11), "Value")
    
    
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    temp = lRow_temp - 18
                    ' kiem tra neu set lai cong thuc
                    ' sua ct tinh
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("BH"), 15 + temp, vSoTien
                    strFormula = getFormulaTienPNC(temp, CDbl(vSoTien), "BH" & 15 + temp)
                    
                    'fpSpread1.Formula = "IF((BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100)>0,ROUND(BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100,0),0)"
                    fpSpread1.Formula = strFormula
                    ' end
                    fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10), "Value")
                ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "02" Then
                ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "72" Then
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    fpSpread1.Formula = "BD5"
                    fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7), "Value")
    
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    temp = lRow_temp - 18
                    fpSpread1.Formula = ""
                    fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6), "Value")
                Else
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    fpSpread1.Formula = "BD5"
                    fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7), "Value")
    
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    temp = lRow_temp - 18
                    ' kiem tra set lai cong thuc
                    ' sua ct tinh
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("BH"), 15 + temp, vSoTien
                    strFormula = getFormulaTienPNC(temp, CDbl(vSoTien), "BH" & 15 + temp)
                    
                    'fpSpread1.Formula = "IF((BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100)>0,ROUND(BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100,0),0)"
                    fpSpread1.Formula = strFormula
                    ' end
                    fpSpread1.value = GetAttribute(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6), "Value")
                End If
        Else
                If GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "01" Then
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 11)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
    
                    fpSpread1.Formula = "BD5"
                    
    
    
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 10)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    temp = lRow_temp - 18
                    ' kiem tra neu set lai cong thuc
                    ' sua ct tinh
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("BH"), 15 + temp, vSoTien
                    strFormula = getFormulaTienPNC(temp, CDbl(vSoTien), "BH" & 15 + temp)
                    
                    'fpSpread1.Formula = "IF((BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100)>0,ROUND(BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100,0),0)"
                    fpSpread1.Formula = strFormula
                    ' end
                    
                ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "02" Then
                ElseIf GetAttribute(TAX_Utilities_New.NodeMenu, "ID") = "72" Then
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    fpSpread1.Formula = "BD5"
                    
    
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    temp = lRow_temp - 18
                    fpSpread1.Formula = ""
                    
                Else
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    fpSpread1.Formula = "BD5"
                    
                    Set xmlNodeCell_temp = TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_New.Data(TAX_Utilities_New.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
                    ParserCellID fpSpread1, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                    fpSpread1.sheet = fpSpread1.SheetCount - 1
                    fpSpread1.Col = lCol_temp
                    fpSpread1.Row = lRow_temp
                    temp = lRow_temp - 18
                    ' kiem tra set lai cong thuc
                    ' sua ct tinh
                    fpSpread1.GetText fpSpread1.ColLetterToNumber("BH"), 15 + temp, vSoTien
                    strFormula = getFormulaTienPNC(temp, CDbl(vSoTien), "BH" & 15 + temp)
                    
                    'fpSpread1.Formula = "IF((BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100)>0,ROUND(BH" & 15 + temp & "*BE" & 17 + temp & "*0.05/100,0),0)"
                    fpSpread1.Formula = strFormula
                    ' end
                End If
        End If
'-------------------------------------------------------------------
        fpSpread1.ActiveSheet = fpSpread1.SheetCount - 1
        'DisplayMessage "0222", msOKOnly, miInformation
    End If
End Sub

' ham de tai bang ke 01_2
Private Sub moveData01_2()
Dim value As String
Dim xmlDocument As New MSXML.DOMDocument
Dim xmlNode As MSXML.IXMLDOMNode

Dim i, count, count1, count2 As Long
Dim inc As Boolean
Dim colStart As Integer
Dim varMenuId As String

Dim lRow2s As Long
Dim incSession As Integer

On Error GoTo ErrHandle

incSession = 0

fpSpread1.EventEnabled(EventAllEvents) = False
    ' Truong hop them du lieu va xoa du lieu da ton tai
    If themXoaDuLieu Then
        ResetData
        ResetDataAndForm mCurrentSheet
    End If
    
' Lay ID cua Menu
varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")

fpSpread2.Visible = False
ProgressBar1.Visible = True
ProgressBar1.max = fpSpread2.MaxRows
ProgressBar1.value = 0
If Trim(varMenuId) = "01" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_1_GTGT.xml"))
    colStart = 3
End If

Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
   fpSpread1.EventEnabled(EventAllEvents) = False
   fpSpread1.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row1").Item(0).Text)
   fpSpread2.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row2").Item(0).Text)
   fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
   count1 = Conversion.CInt(xmlDocument.getElementsByTagName("count").Item(0).Text)
   
    
    ' Truong hop them tiep du lieu
    Dim xmlSecionNode As MSXML.IXMLDOMNode
    Dim currentRow As Long
    Dim varData1, varData2 As Variant
    If themDuLieu Then
        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(0)
        'fpSpread1.Visible = False
        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
            If (xmlSecionNode.childNodes.length = 1) Then
                fpSpread1.sheet = mCurrentSheet
                fpSpread1.GetText colStart, fpSpread1.Row, varData1
                fpSpread1.GetText colStart + 1, fpSpread1.Row, varData2
                If Trim(varData1) = vbNullString And Trim(varData2) = vbNullString Then
                    fpSpread1.Row = fpSpread1.Row
                Else
                    InsertNode colStart, currentRow - 1
                    fpSpread1.Row = currentRow
                End If
            Else
                InsertNode colStart, currentRow - 1
                fpSpread1.Row = currentRow
            End If
        End If
    End If
'    ' Ket thuc truong hop them tiep du lieu


Do While count < count1 And count2 < fpSpread2.MaxRows
DoEvents
Frame2.Enabled = False
ProgressBar1.value = fpSpread2.Row
'check next row
    fpSpread1.sheet = mCurrentSheet
    fpSpread2.Row = fpSpread2.Row + 1
    value = fpSpread2.value
    If ((Mid(value, 1, 1) = "T" Or Trim(value) = "" Or Trim(value) = vbNullString)) Then
        count = count + 1
        inc = True
        ProgressBar1.value = fpSpread2.MaxRows
    ElseIf count = count1 And value = "" Then
        count = count + 1
    Else
        InsertNode colStart, fpSpread1.Row
        inc = False
        count2 = count2 + 1
    End If
        fpSpread2.Row = fpSpread2.Row - 1
    'insert cell
        For Each xmlNode In xmlNodeListMap
            fpSpread2.Col = Conversion.CInt(GetAttribute(xmlNode, "c2"))
            value = fpSpread2.value
           If value <> "" Or value <> vbNullString Then
            fpSpread1.Col = Conversion.CInt(GetAttribute(xmlNode, "c1"))
    'check type of cell
            If Conversion.CInt(GetAttribute(xmlNode, "type")) = 13 Then
                If fpSpread1.CellType = CellTypeNumber Then
                    fpSpread1.TypeNumberNegStyle = TypeNumberNegStyle1
                End If
                fpSpread1.value = value
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 12 Then
                fpSpread1.value = value
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 2 Then
                fpSpread1.Text = Left(fpSpread2.Text, IIf(InStr(1, fpSpread2.Text, ".") <> 0, InStr(1, fpSpread2.Text, ".") - 1, Len(fpSpread2.Text)))
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 1 Then
              If IsDate(fpSpread2.Text) Then
                Dim arrStr() As String
                Dim sDate As String
                If InStr(1, fpSpread2.Text, "-") <> 0 Then
                    arrStr = Split(fpSpread2.Text, "-")
                Else
                    arrStr = Split(fpSpread2.Text, "/")
                End If
                
                sDate = Right("00" & arrStr(0), 2) & "/" & Right("00" & arrStr(1), 2) & "/" & Right("20" & arrStr(2), 4)
                
                fpSpread1.Text = sDate
    
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
             End If
            Else
                Select Case strfileFont
                   Case "TCVN"
                      fpSpread1.Text = TAX_Utilities_New.Convert(value, TCVN, UNICODE)
                   Case "VNI"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VNI, UNICODE)
                   Case "VIQR"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VIQR, UNICODE)
                   Case "VISCII"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VISCII, UNICODE)
                   Case Else
                    fpSpread1.Text = value
                End Select
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
            End If
            
          End If
        Next
    'next row
        If inc = True Then
                   If themDuLieu Then
                               'have 2 hidden row
                        fpSpread1.Row = fpSpread1.Row + 5
                        fpSpread2.Row = fpSpread2.Row + 3
                        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(count)
                        'fpSpread1.Visible = False
                        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
                            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
                            If (xmlSecionNode.childNodes.length = 1) Then
                                fpSpread1.sheet = mCurrentSheet
                                fpSpread1.GetText colStart, fpSpread1.Row, varData1
                                fpSpread1.GetText colStart + 1, fpSpread1.Row, varData2
                                If Trim(varData1) = vbNullString And Trim(varData2) = vbNullString Then
                                    fpSpread1.Row = fpSpread1.Row
                                Else
                                    InsertNode colStart, currentRow - 1
                                    fpSpread1.Row = currentRow
                                End If
                            Else
                                InsertNode colStart, currentRow - 1
                                fpSpread1.Row = currentRow
                            End If
                        End If
                    End If
                    ' Ket thuc truong hop them tiep du lieu

        Else
            fpSpread1.Row = fpSpread1.Row + 1
            fpSpread2.Row = fpSpread2.Row + 1
        End If
            fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
            value = fpSpread2.value
    Loop
 ProgressBar1.Visible = False
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True
 If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
 
 Exit Sub
ErrHandle:
 DisplayMessage "0122", msOKOnly, miCriticalError
 ProgressBar1.Visible = False
 ResetData
 ResetDataAndForm mCurrentSheet
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True

End Sub

Private Sub moveDataToKhai08B()
    Dim value       As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode     As MSXML.IXMLDOMNode
    Dim i, count, count1, count2 As Long, lCol As Long, lRow As Long
    Dim inc          As Boolean
    Dim colStart     As Integer
    Dim varMenuId    As String
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim IsUpdate     As Boolean
    On Error GoTo ErrHandle

    'Delete data exit
    fpSpread1.EventEnabled(EventAllEvents) = False
    mCurrentSheet = 1

    For Each xmlNodeReset In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
        fpSpread1.sheet = mCurrentSheet
        ParserCellID fpSpread1, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
        fpSpread1.Col = lCol
        fpSpread1.Row = lRow

        If (lRow < 57 Or lRow > fpSpread1.MaxRows - 4) Then

            GoTo nextClear1
        Else

            Select Case fpSpread1.CellType

                Case CellTypeCheckBox
                    fpSpread1.Text = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)

                Case CellTypeComboBox
                    fpSpread1.Text = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)

                Case CellTypeNumber
                    fpSpread1.value = 0
                    IsUpdate = UpdateCell(lCol, lRow, "0")

                Case Else
                    fpSpread1.value = vbNullString
                    IsUpdate = UpdateCell(lCol, lRow, vbNullString)
            End Select

        End If

        'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
        TAX_Utilities_New.AdjustData(0) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_New.AdjustData(mCurrentSheet - 1))
nextClear1:
    Next

    ResetDataAndForm mCurrentSheet

    ' Lay ID cua Menu
    varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")

    fpSpread2.Visible = False
    ProgressBar1.Visible = True
    fpSpread2.sheet = mCurrentSheet
    ProgressBar1.max = fpSpread2.MaxRows
    ProgressBar1.value = 0
 
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\TK_BK_08B_TNCN.xml"))
    colStart = 3

    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
    fpSpread1.EventEnabled(EventAllEvents) = False
    fpSpread1.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row1").Item(0).Text)
    fpSpread2.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row2").Item(0).Text)
    fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
    count1 = Conversion.CInt(xmlDocument.getElementsByTagName("count").Item(0).Text)
   
    Do While count < count1 And count2 < fpSpread2.MaxRows
        DoEvents
        Frame2.Enabled = False
        ProgressBar1.value = fpSpread2.Row
        'check next row
        fpSpread1.sheet = mCurrentSheet
        fpSpread2.sheet = mCurrentSheet
        fpSpread2.Row = fpSpread2.Row + 1
        value = fpSpread2.value
    
        If Trim(value) = "" Or Trim(value) = vbNullString Or Trim(value) = "aa" Then
            count = count + 1
            inc = True
            ProgressBar1.value = fpSpread2.MaxRows
        ElseIf count = count1 And value = "" Then
            count = count + 1
        Else
            InsertNode colStart, fpSpread1.Row
            inc = False
            count2 = count2 + 1
        End If

        fpSpread2.Row = fpSpread2.Row - 1

        'insert cell
        For Each xmlNode In xmlNodeListMap
            fpSpread2.Col = Conversion.CInt(GetAttribute(xmlNode, "c2"))
            value = fpSpread2.value

            If value <> "" Or value <> vbNullString Then
                fpSpread1.Col = Conversion.CInt(GetAttribute(xmlNode, "c1"))

                'check type of cell
                If Conversion.CInt(GetAttribute(xmlNode, "type")) = 13 Then
                    If fpSpread1.CellType = CellTypeNumber Then
                        fpSpread1.TypeNumberNegStyle = TypeNumberNegStyle1
                    End If

                    fpSpread1.value = Round(value, 0)
                    UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
                ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 12 Then
                    fpSpread1.value = value
                    UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
                ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 2 Then
                    '            fpSpread2.CellType = CellTypeNumber
                    '            fpSpread2.TypeNumberDecPlaces = 0
                    fpSpread1.Text = Left(fpSpread2.Text, IIf(InStr(1, fpSpread2.Text, ".") <> 0, InStr(1, fpSpread2.Text, ".") - 1, Len(fpSpread2.Text)))
                    UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
                ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 1 Then

                    If IsDate(fpSpread2.Text) Then
                        Dim arrStr() As String
                        Dim sDate    As String

                        'fpSpread2.CellType = CellTypeDate
                        'fpSpread2.TypeDateFormat = TypeDateFormatDDMMYY
                        'Dim objCvt As New DateUtils
                        'fpSpread2.Text = CStr(objCvt.ToDate(fpSpread2.Text, "DD/MM/YYYY"))
                        If InStr(1, fpSpread2.Text, "-") <> 0 Then
                            arrStr = Split(fpSpread2.Text, "-")
                        Else
                            arrStr = Split(fpSpread2.Text, "/")
                        End If
            
                        sDate = Right("00" & arrStr(0), 2) & "/" & Right("00" & arrStr(1), 2) & "/" & Right("20" & arrStr(2), 4)
            
                        fpSpread1.Text = sDate

                        UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
                    End If

                Else

                    Select Case strfileFont

                        Case "TCVN"
                            fpSpread1.Text = TAX_Utilities_New.Convert(value, TCVN, UNICODE)

                        Case "VNI"
                            fpSpread1.Text = TAX_Utilities_New.Convert(value, VNI, UNICODE)

                        Case "VIQR"
                            fpSpread1.Text = TAX_Utilities_New.Convert(value, VIQR, UNICODE)

                        Case "VISCII"
                            fpSpread1.Text = TAX_Utilities_New.Convert(value, VISCII, UNICODE)

                        Case Else
                            fpSpread1.Text = value
                    End Select

                    UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
                End If
        
            End If

        Next

        'next row
        If inc = True Then
            'have 2 hidden row
            fpSpread1.Row = fpSpread1.Row + 5
            fpSpread2.Row = fpSpread2.Row + 3
        Else
            fpSpread1.Row = fpSpread1.Row + 1
            fpSpread2.Row = fpSpread2.Row + 1
        End If

        fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
        value = fpSpread2.value
    Loop

    ProgressBar1.Visible = False
    Frame2.Enabled = True
    fpSpread1.EventEnabled(EventAllEvents) = True

    If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
 
    Exit Sub
ErrHandle:
    DisplayMessage "0122", msOKOnly, miCriticalError
    ProgressBar1.Visible = False
    ResetData
    ResetDataAndForm mCurrentSheet
    Frame2.Enabled = True
    fpSpread1.EventEnabled(EventAllEvents) = True

End Sub



' Move data 01/TTDB
Private Sub moveData01TTDB()
Dim value As String
Dim xmlDocument As New MSXML.DOMDocument
Dim xmlNode As MSXML.IXMLDOMNode

Dim i, count, count1, count2 As Long
Dim inc As Boolean
Dim colStart As Integer
Dim varMenuId As String

Dim lRow2s As Long
Dim incSession As Integer

On Error GoTo ErrHandle

incSession = 0

fpSpread1.EventEnabled(EventAllEvents) = False
    ' Truong hop them du lieu va xoa du lieu da ton tai
    If themXoaDuLieu Then
        ResetData
        ResetDataAndForm mCurrentSheet
    End If
    
' Lay ID cua Menu
varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")

fpSpread2.Visible = False
ProgressBar1.Visible = True
ProgressBar1.max = fpSpread2.MaxRows
ProgressBar1.value = 0
If Trim(varMenuId) = "05" And fpSpread1.ActiveSheet = 2 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_1_TTDB.xml"))
    colStart = 3
ElseIf Trim(varMenuId) = "05" And fpSpread1.ActiveSheet = 3 Then
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_2_TTDB.xml"))
    colStart = 3
End If

Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell")
   fpSpread1.EventEnabled(EventAllEvents) = False
   fpSpread1.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row1").Item(0).Text)
   fpSpread2.Row = Conversion.CInt(xmlDocument.getElementsByTagName("Row2").Item(0).Text)
   fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
   count1 = Conversion.CInt(xmlDocument.getElementsByTagName("count").Item(0).Text)
   
    
    ' Truong hop them tiep du lieu
    Dim xmlSecionNode As MSXML.IXMLDOMNode
    Dim currentRow As Long
    Dim varData1, varData2 As Variant
    If themDuLieu Then
        Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(0)
        'fpSpread1.Visible = False
        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
            If (xmlSecionNode.childNodes.length = 1) Then
                fpSpread1.sheet = mCurrentSheet
                fpSpread1.GetText colStart, fpSpread1.Row, varData1
                fpSpread1.GetText colStart + 1, fpSpread1.Row, varData2
                If Trim(varData1) = vbNullString And Trim(varData2) = vbNullString Then
                    fpSpread1.Row = fpSpread1.Row
                Else
                    InsertNode colStart, currentRow - 1
                    fpSpread1.Row = currentRow
                End If
            Else
                InsertNode colStart, currentRow - 1
                fpSpread1.Row = currentRow
            End If
        End If
    End If
    ' Ket thuc truong hop them tiep du lieu
' Dat lai vi tri row cho phu luc 01-2 cua to 02 GTGT

Do While count < count1 And count2 < fpSpread2.MaxRows
DoEvents
Frame2.Enabled = False
ProgressBar1.value = fpSpread2.Row
'check next row
    fpSpread1.sheet = mCurrentSheet
    fpSpread2.Row = fpSpread2.Row + 1
    value = fpSpread2.value
    If ((Mid(value, 1, 1) = "T" Or Trim(value) = "" Or Trim(value) = vbNullString) And (Trim(varMenuId) = "01" Or Trim(varMenuId) = "02" Or Trim(varMenuId) = "14" Or Trim(varMenuId) = "05" Or Trim(varMenuId) = "59")) Or ((Trim(value) = "" Or Trim(value) = vbNullString)) Then
        count = count + 1
        inc = True
        ProgressBar1.value = fpSpread2.MaxRows
    ElseIf count = count1 And value = "" Then
        count = count + 1
    Else
        InsertNode colStart, fpSpread1.Row
        inc = False
        count2 = count2 + 1
    End If
        fpSpread2.Row = fpSpread2.Row - 1
    'insert cell
        For Each xmlNode In xmlNodeListMap
            fpSpread2.Col = Conversion.CInt(GetAttribute(xmlNode, "c2"))
            value = fpSpread2.value
           If value <> "" Or value <> vbNullString Then
            fpSpread1.Col = Conversion.CInt(GetAttribute(xmlNode, "c1"))
    'check type of cell
            If Conversion.CInt(GetAttribute(xmlNode, "type")) = 13 Then
                If fpSpread1.CellType = CellTypeNumber Then
                    fpSpread1.TypeNumberNegStyle = TypeNumberNegStyle1
                End If
                fpSpread1.value = value
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 12 Then
                fpSpread1.value = value
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 2 Then
                fpSpread1.Text = Left(fpSpread2.Text, IIf(InStr(1, fpSpread2.Text, ".") <> 0, InStr(1, fpSpread2.Text, ".") - 1, Len(fpSpread2.Text)))
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.value
            ElseIf Conversion.CInt(GetAttribute(xmlNode, "type")) = 1 Then
              If IsDate(fpSpread2.Text) Then
                Dim arrStr() As String
                Dim sDate As String
                If InStr(1, fpSpread2.Text, "-") <> 0 Then
                    arrStr = Split(fpSpread2.Text, "-")
                Else
                    arrStr = Split(fpSpread2.Text, "/")
                End If
                
                sDate = Right("00" & arrStr(0), 2) & "/" & Right("00" & arrStr(1), 2) & "/" & Right("20" & arrStr(2), 4)
                
                fpSpread1.Text = sDate
    
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
             End If
            Else
                Select Case strfileFont
                   Case "TCVN"
                      fpSpread1.Text = TAX_Utilities_New.Convert(value, TCVN, UNICODE)
                   Case "VNI"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VNI, UNICODE)
                   Case "VIQR"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VIQR, UNICODE)
                   Case "VISCII"
                    fpSpread1.Text = TAX_Utilities_New.Convert(value, VISCII, UNICODE)
                   Case Else
                    fpSpread1.Text = value
                End Select
                UpdateCell fpSpread1.Col, fpSpread1.Row, fpSpread1.Text
            End If
            
          End If
        Next
    'next row
        If inc = True Then
                Set xmlNodeListMap = xmlDocument.getElementsByTagName("cell1")
                
                Dim temp As Variant
                Dim temp1 As Double
                fpSpread1.Row = fpSpread1.Row + 9
                fpSpread2.Row = fpSpread2.Row + 3
'            End If
            'test
              If themDuLieu Then
                Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(count)
                'fpSpread1.Visible = False
                If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
                    currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
                    If (xmlSecionNode.childNodes.length = 1) Then
                        fpSpread1.sheet = mCurrentSheet
                        fpSpread1.GetText colStart, fpSpread1.Row, varData1
                        fpSpread1.GetText colStart + 1, fpSpread1.Row, varData2
                        If Trim(varData1) = vbNullString And Trim(varData2) = vbNullString Then
                            fpSpread1.Row = fpSpread1.Row
                        Else
                            InsertNode colStart, currentRow - 1
                            fpSpread1.Row = currentRow
                        End If
                    Else
                        InsertNode colStart, currentRow - 1
                        fpSpread1.Row = currentRow
                    End If
                End If
            End If
            
            ' end test
        Else
            fpSpread1.Row = fpSpread1.Row + 1
            fpSpread2.Row = fpSpread2.Row + 1
        End If
            fpSpread2.Col = Conversion.CInt(xmlDocument.getElementsByTagName("Col").Item(0).Text)
            value = fpSpread2.value
    Loop
 ProgressBar1.Visible = False
 Frame2.Enabled = True
 'fpSpread1.EventEnabled(EventAllEvents) = True
 If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
 ' Group cac mat hang thanh nhom
 If fpSpread1.ActiveSheet = 2 Then
    fpSpread1.Col = fpSpread1.ColLetterToNumber("L")
    fpSpread1.Row = 37
    fpSpread1.SetActiveCell fpSpread1.Col, fpSpread1.Row
    fpSpread1.SetFocus
    fpSpread1_LeaveCell fpSpread1.ColLetterToNumber("L"), 37, fpSpread1.ColLetterToNumber("N"), 37, True
 End If
 fpSpread1.EventEnabled(EventAllEvents) = True
 
 Exit Sub
ErrHandle:
 DisplayMessage "0122", msOKOnly, miCriticalError
 ProgressBar1.Visible = False
 ResetData
 ResetDataAndForm mCurrentSheet
 Frame2.Enabled = True
 fpSpread1.EventEnabled(EventAllEvents) = True

End Sub


Public Function delNullRowOn05(sheet As Long)
    On Error GoTo ErrorHandle
    Dim xmlNodeListSec As MSXML.IXMLDOMNodeList
    Dim xmlNodeListRow As MSXML.IXMLDOMNodeList
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeSec As MSXML.IXMLDOMNode
    Dim xmlNodeRow As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim numSec, Row, row1, celllg, hasVl As Long
    Dim sumRowDel, countDel As Long
    Dim strCol As String
    Dim colArr() As String
    Dim cellid, value As Variant
    Dim OldSheet As Long
    
    'dntai para templ
    Dim i As Long, j As Integer, varTemp As Variant, rowStart As Long
    
    Dim maxRow As Long
    'set sheet current
    OldSheet = fpSpread1.ActiveSheet
    'get section and check dynamic
    Set xmlNodeListSec = TAX_Utilities_New.Data(sheet).getElementsByTagName("Section")
    
    If GetAttribute(xmlNodeListSec.Item(0), "Dynamic") = "1" Then
        fpSpread1.sheet = sheet + 1
        'get CellID cell dau tien dong dau tien
        cellid = GetAttribute(xmlNodeListSec.Item(0).childNodes(0).firstChild, "CellID")
        'set location cell to array
        If fpSpread1.sheet = 2 Then
            strCol = "D~E~F~G~H~I~J~K~L~M~N~O~Q~R~S"
            colArr = Split(strCol, "~")
        ElseIf fpSpread1.sheet = 3 Then
            strCol = "C~D~E~F~G~H~I~J"
            colArr = Split(strCol, "~")
        ElseIf fpSpread1.sheet = 4 Then
            strCol = "C~D~E~F~G~H~I~J"
            colArr = Split(strCol, "~")
        End If
        With fpSpread1
            .EventEnabled(EventAllEvents) = False
            .Col = .ColLetterToNumber("B")
            'set row to start loop
            i = CLng(Mid(cellid, InStr(1, cellid, "_") + 1, Len(cellid)))
            'set rowStart de dung so sanh
            rowStart = i + 1
            .Row = i + 1
            Do
                If .Text = "aa" Then
                    Exit Do
                End If
                
                hasVl = 0
                For j = 0 To UBound(colArr)
                    .Col = .ColLetterToNumber(colArr(j))
                    value = .Text
                    If (Trim(value) <> vbNullString And Trim(value) <> "0") Then
                        hasVl = hasVl + 1
                        Exit For
                    End If
                Next
                
                If hasVl = 0 Then
                        fpSpread1.ActiveSheet = sheet + 1
'                        DeleteNode sheet + 1, .ColLetterToNumber(colArr(0)), .Row, True
                        .GetText .ColLetterToNumber("B"), .Row + 1, varTemp
                        'kiem tra neu tren sheet neu chi co 1 dong thi khong duoc xoa
                        If Trim(varTemp) = "aa" And i = rowStart Then
                            Exit Do
                        End If
                        DeleteRow sheet + 1, .Row, 1
                Else
                        i = i + 1
                        .Row = i
                End If
                    .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
            
            i = CLng(Mid(cellid, InStr(1, cellid, "_") + 1, Len(cellid)))
            .Row = i
            
            For j = 0 To UBound(colArr)
                .Col = .ColLetterToNumber(colArr(j))
                .CellNote = ""
                .BackColor = vbWhite
            Next
            .EventEnabled(EventAllEvents) = True
        End With
    End If
    
    
    
''    sumRowDel = TAX_Utilities_New.Data(sheet).getElementsByTagName("Cell").length
    

    ' Xem lai vi sao lai countDel <> 19
    ' 09112011

'    maxRow = fpSpread1.MaxRows
    'Do While countDel <> 19
'    Do While countDel <> maxRow
'        countDel = countDel + 1
'        Set xmlNodeListSec = TAX_Utilities_New.Data(sheet).getElementsByTagName("Section")
''sec
'        numSec = 0
'        For Each xmlNodeSec In xmlNodeListSec
'            If GetAttribute(xmlNodeSec, "Dynamic") = "1" Then
'                Set xmlNodeListRow = xmlNodeListSec.Item(numSec).childNodes
'        'row
'                Row = 0
'                For Each xmlNodeRow In xmlNodeListRow
'                    hasVl = 0
'                    Set xmlNodeListCell = xmlNodeListRow.Item(Row).childNodes
'               'cell
'                    For Each xmlNodeCell In xmlNodeListCell
'                        value = GetAttribute(xmlNodeCell, "Value")
'                        'If GetAttribute(xmlNodeCell, "FirstCell") = "" And value <> "" And value <> "0" And value <> "cbo" And value <> "0%" And value <> "5%" And value <> "10%" Then
'                        If (GetAttribute(xmlNodeCell, "FirstCell") <> "" And value <> "") Or (GetAttribute(xmlNodeCell, "FirstCell") = "" And value <> "" And value <> "0" And value <> "cbo" And value <> "0%" And value <> "5%" And value <> "10%") Then
'                            hasVl = hasVl + 1
'                        End If
'                        cellid = GetAttribute(xmlNodeCell, "CellID")
'                    Next
'                    If hasVl = 0 Then
'                        If Mid(cellid, 2, 1) = "_" Then
'                            fpSpread1.ActiveSheet = sheet + 1
'                            DeleteNode sheet + 1, fpSpread1.ColLetterToNumber(Left(cellid, 1)), CLng(Right(cellid, Len(cellid) - 2)), True
'                             Exit For
'                        ElseIf Mid(cellid, 3, 1) = "_" Then
'                            fpSpread1.ActiveSheet = sheet + 1
'                            DeleteNode sheet + 1, fpSpread1.ColLetterToNumber(Left(cellid, 2)), CLng(Right(cellid, Len(cellid) - 3)), True
'                            Exit For
'                        Else
'
'                        End If
'                    End If
'                    Row = Row + 1
'                Next
'            End If
'            numSec = numSec + 1
'        Next
'    Loop
    fpSpread1.ActiveSheet = OldSheet
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "delNullRowOn05", Err.Number, Err.Description
End Function


Public Function delNullRowOn01(sheet As Long)
    On Error GoTo ErrorHandle
    Dim Row, row1, celllg, hasVl As Long
    Dim sumRowDel, countDel As Long
    Dim strCol As String
    Dim colArr() As String
    Dim cellid, value As Variant
    Dim OldSheet As Long
    
    'dntai para templ
    Dim i As Long, j As Integer, varTemp As Variant, rowStart As Long, countSec As Integer
    
    Dim maxRow As Long
    'set sheet current
    OldSheet = fpSpread1.ActiveSheet
    'get section and check dynamic

    
    If sheet = 1 Or sheet = 2 Then
        fpSpread1.sheet = sheet + 1
        'set countSec de dem section
        countSec = 1
        'get CellID cell dau tien dong dau tien
'        cellid = GetAttribute(xmlNodeListSec.Item(0).childNodes(0).firstChild, "CellID")
        'set location cell to array
        If fpSpread1.sheet = 2 Then
            strCol = "C~D~E~F~G~H~I~K~L"
            colArr = Split(strCol, "~")
        ElseIf fpSpread1.sheet = 3 Then
            strCol = "C~D~E~F~G~H~I~J~K~L"
            colArr = Split(strCol, "~")
        End If
        With fpSpread1
            .EventEnabled(EventAllEvents) = False
            .Col = .ColLetterToNumber("B")
            'set row to start loop
            i = 8
            'set rowStart de dung so sanh
            rowStart = i
            .Row = i
            Do
                hasVl = 0
                For j = 0 To UBound(colArr)
                    .Col = .ColLetterToNumber(colArr(j))
                    value = .Text
                    If (Trim(value) <> vbNullString And Trim(value) <> "0") Then
                        hasVl = hasVl + 1
                        Exit For
                    End If
                Next
                
                If hasVl = 0 Then
                        fpSpread1.ActiveSheet = sheet + 1
'                        DeleteNode sheet + 1, .ColLetterToNumber(colArr(0)), .Row, True
                        .GetText .ColLetterToNumber("B"), .Row + 1, varTemp
                        'kiem tra neu tren sheet neu chi co 1 dong thi khong duoc xoa
                        If (Trim(varTemp) = "aa" Or Trim(varTemp) = "bb" Or Trim(varTemp) = "cc" Or Trim(varTemp) = "dd" Or Trim(varTemp) = "ee") And i = rowStart Then
                            If Trim(varTemp) = "ee" Then
                                i = i + 1
                            Else
                                i = i + 5
                            End If
                            rowStart = i
                            .Row = i
                        Else
                            DeleteNode sheet + 1, .ColLetterToNumber("C"), .Row, True
                            .Row = i
                        End If
                        
                Else
                        i = i + 1
                        .Row = i
                End If
                
                .Col = .ColLetterToNumber("B")
                varTemp = .Text
                If (Trim(varTemp) = "aa" Or Trim(varTemp) = "bb" Or Trim(varTemp) = "cc" Or Trim(varTemp) = "dd") Then
                        i = i + 4
                        .Row = i
                        rowStart = i
                End If
            Loop Until .Text = "ee"
            .EventEnabled(EventAllEvents) = True
        End With
    End If
    
 
    fpSpread1.ActiveSheet = OldSheet
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "delNullRowOn01", Err.Number, Err.Description
End Function



Public Function delNullRowOn06(sheet As Long)
    On Error GoTo ErrorHandle
    Dim xmlNodeListSec As MSXML.IXMLDOMNodeList
    Dim xmlNodeListRow As MSXML.IXMLDOMNodeList
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeSec As MSXML.IXMLDOMNode
    Dim xmlNodeRow As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim numSec, Row, row1, celllg, hasVl As Long
    Dim sumRowDel, countDel As Long
    Dim strCol As String
    Dim colArr() As String
    Dim cellid, value As Variant
    Dim OldSheet As Long
    
    'dntai para templ
    Dim i As Long, j As Integer, varTemp As Variant, rowStart As Long
    
    Dim maxRow As Long
    'set sheet current
    OldSheet = fpSpread1.ActiveSheet
    'get section and check dynamic
    Set xmlNodeListSec = TAX_Utilities_New.Data(sheet).getElementsByTagName("Section")
    
    If GetAttribute(xmlNodeListSec.Item(0), "Dynamic") = "1" Then
        fpSpread1.sheet = sheet + 1
        'get CellID cell dau tien dong dau tien
        cellid = GetAttribute(xmlNodeListSec.Item(0).childNodes(0).firstChild, "CellID")
        'set location cell to array
        If fpSpread1.sheet = 2 Then
            strCol = "D~E~F~G~H"
            colArr = Split(strCol, "~")
        End If
        With fpSpread1
            .EventEnabled(EventAllEvents) = False
            .Col = .ColLetterToNumber("B")
            'set row to start loop
            i = CLng(Mid(cellid, InStr(1, cellid, "_") + 1, Len(cellid)))
            'set rowStart de dung so sanh
            rowStart = i + 1
            .Row = i + 1
            Do
                If .Text = "aa" Then
                    Exit Do
                End If
                
                hasVl = 0
                For j = 0 To UBound(colArr)
                    .Col = .ColLetterToNumber(colArr(j))
                    value = .Text
                    If (Trim(value) <> vbNullString And Trim(value) <> "0") Then
                        hasVl = hasVl + 1
                        Exit For
                    End If
                Next
                
                If hasVl = 0 Then
                        fpSpread1.ActiveSheet = sheet + 1
'                        DeleteNode sheet + 1, .ColLetterToNumber(colArr(0)), .Row, True
                        .GetText .ColLetterToNumber("B"), .Row + 1, varTemp
                        'kiem tra neu tren sheet neu chi co 1 dong thi khong duoc xoa
                        If Trim(varTemp) = "aa" And i = rowStart Then
                            Exit Do
                        End If
                        DeleteRow sheet + 1, .Row, 1
                Else
                        i = i + 1
                        .Row = i
                End If
                    .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
            
            i = CLng(Mid(cellid, InStr(1, cellid, "_") + 1, Len(cellid)))
            .Row = i
            
            For j = 0 To UBound(colArr)
                .Col = .ColLetterToNumber(colArr(j))
                .CellNote = ""
                .BackColor = vbWhite
            Next
            .EventEnabled(EventAllEvents) = True
        End With
    End If
    
    fpSpread1.ActiveSheet = OldSheet
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "delNullRowOn06", Err.Number, Err.Description
End Function

Public Sub copyFormulasSheet2(numRow As Long, fps As fpSpread, rowStart As Long)
    Dim a As Long
    a = 0

    With fps

        .sheet = 2
            
        'truong hop so ban ghe len hon 10000
        If numRow >= 10000 Then

            Do While a * 2 <= 1024

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), rowStart + a, .ColLetterToNumber("BB"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), rowStart + a, .ColLetterToNumber("BC"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), rowStart + a, .ColLetterToNumber("BD"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), rowStart + a, .ColLetterToNumber("M"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), rowStart + a, .ColLetterToNumber("Q"), (rowStart + a + 1)

                    a = a + 2
                ElseIf a <> 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), rowStart + a - 1, .ColLetterToNumber("BB"), rowStart + a
                    .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), rowStart + a - 1, .ColLetterToNumber("BC"), rowStart + a
                    .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), rowStart + a - 1, .ColLetterToNumber("BD"), rowStart + a
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), rowStart + a - 1, .ColLetterToNumber("M"), rowStart + a
                    .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), rowStart + a - 1, .ColLetterToNumber("Q"), rowStart + a

                    a = a * 2
                End If

            Loop
                 
            a = 1
            Dim dem As Long
            Dim du  As Long
            dem = numRow \ 1024
            du = numRow Mod 1024

            If dem > 0 Then

                Do While a < dem
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), 1024 + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), 1024 + rowStart - 1, .ColLetterToNumber("BB"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), 1024 + rowStart - 1, .ColLetterToNumber("BC"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), 1024 + rowStart - 1, .ColLetterToNumber("BD"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), 1024 + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), 1024 + rowStart - 1, .ColLetterToNumber("M"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), 1024 + rowStart - 1, .ColLetterToNumber("Q"), rowStart + 1024 * a
                        
                    a = a + 1
                Loop

                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), du + rowStart - 1, .ColLetterToNumber("BB"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), du + rowStart - 1, .ColLetterToNumber("BC"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), du + rowStart - 1, .ColLetterToNumber("BD"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), du + rowStart - 1, .ColLetterToNumber("M"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), du + rowStart - 1, .ColLetterToNumber("Q"), rowStart + 1024 * a
                        
            Else
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), du + rowStart - 1, .ColLetterToNumber("BB"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), du + rowStart - 1, .ColLetterToNumber("BC"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), du + rowStart - 1, .ColLetterToNumber("BD"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), du + rowStart - 1, .ColLetterToNumber("M"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), du + rowStart - 1, .ColLetterToNumber("Q"), rowStart + 1024 * (a - 1)
                        
            End If

            ' truong hop nho hon 10000
        ElseIf numRow < 10000 Then
            a = 0

            Do While a * 2 < numRow

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), rowStart + a, .ColLetterToNumber("BB"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), rowStart + a, .ColLetterToNumber("BC"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), rowStart + a, .ColLetterToNumber("BD"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), rowStart + a, .ColLetterToNumber("M"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), rowStart + a, .ColLetterToNumber("Q"), (rowStart + a + 1)

                    a = a + 2
                ElseIf a <> 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), rowStart + a - 1, .ColLetterToNumber("BB"), rowStart + a
                    .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), rowStart + a - 1, .ColLetterToNumber("BC"), rowStart + a
                    .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), rowStart + a - 1, .ColLetterToNumber("BD"), rowStart + a
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), rowStart + a - 1, .ColLetterToNumber("M"), rowStart + a
                    .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), rowStart + a - 1, .ColLetterToNumber("Q"), rowStart + a

                    a = a * 2
                End If

            Loop
                
            .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + (numRow - a - 1), .ColLetterToNumber("A"), rowStart + a
            .CopyRange .ColLetterToNumber("BB"), rowStart, .ColLetterToNumber("BB"), rowStart + (numRow - a - 1), .ColLetterToNumber("BB"), rowStart + a
            .CopyRange .ColLetterToNumber("BC"), rowStart, .ColLetterToNumber("BC"), rowStart + (numRow - a - 1), .ColLetterToNumber("BC"), rowStart + a
            .CopyRange .ColLetterToNumber("BD"), rowStart, .ColLetterToNumber("BD"), rowStart + (numRow - a - 1), .ColLetterToNumber("BD"), rowStart + a
            .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + (numRow - a - 1), .ColLetterToNumber("B"), rowStart + a
            .CopyRange .ColLetterToNumber("M"), rowStart, .ColLetterToNumber("M"), rowStart + (numRow - a - 1), .ColLetterToNumber("M"), rowStart + a
            .CopyRange .ColLetterToNumber("Q"), rowStart, .ColLetterToNumber("BA"), rowStart + (numRow - a - 1), .ColLetterToNumber("Q"), rowStart + a
            
        End If
            
    End With

End Sub

Public Sub copyFormulasSheet3(numRow As Long, fps As fpSpread, rowStart As Long)
    Dim a As Long
    a = 0

    With fps

        .sheet = 3

        'truong hop so dong lon hon 10000
        If numRow >= 10000 Then

            Do While a * 2 <= 1024

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), rowStart + a, .ColLetterToNumber("AK"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), rowStart + a, .ColLetterToNumber("K"), (rowStart + a + 1)

                    a = a + 2
                ElseIf a <> 0 Then
                
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), rowStart + a - 1, .ColLetterToNumber("K"), rowStart + a
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), rowStart + a - 1, .ColLetterToNumber("K"), rowStart + a

                    a = a * 2
                End If

            Loop
                 
            a = 1
            Dim dem As Long
            Dim du  As Long
            dem = numRow \ 1024
            du = numRow Mod 1024

            If dem > 0 Then

                Do While a < dem
                
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), 1024 + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), 1024 + rowStart - 1, .ColLetterToNumber("AK"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), 1024 + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), 1024 + rowStart - 1, .ColLetterToNumber("K"), rowStart + 1024 * a
                        
                    a = a + 1
                Loop

                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), du + rowStart - 1, .ColLetterToNumber("AK"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), du + rowStart - 1, .ColLetterToNumber("K"), rowStart + 1024 * a
    
            Else
                
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), du + rowStart - 1, .ColLetterToNumber("AK"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), du + rowStart - 1, .ColLetterToNumber("K"), rowStart + 1024 * (a - 1)

            End If

            ' truong hop nho hon 1024000
        ElseIf numRow < 10000 Then
            a = 0

            Do While a * 2 < numRow

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), rowStart + a, .ColLetterToNumber("AK"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), rowStart + a, .ColLetterToNumber("K"), (rowStart + a + 1)
                    
                    a = a + 2
                ElseIf a <> 0 Then
                
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), rowStart + a - 1, .ColLetterToNumber("AK"), rowStart + a
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), rowStart + a - 1, .ColLetterToNumber("K"), rowStart + a
                
                    a = a * 2
                End If

            Loop
                
            .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + (numRow - a - 1), .ColLetterToNumber("A"), rowStart + a
            .CopyRange .ColLetterToNumber("AK"), rowStart, .ColLetterToNumber("AK"), rowStart + (numRow - a - 1), .ColLetterToNumber("AK"), rowStart + a
            .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + (numRow - a - 1), .ColLetterToNumber("B"), rowStart + a
            .CopyRange .ColLetterToNumber("K"), rowStart, .ColLetterToNumber("AJ"), rowStart + (numRow - a - 1), .ColLetterToNumber("K"), rowStart + a
            

        End If
            
    End With

End Sub

Public Sub copyFormulas06_TNCN(numRow As Long, fps As fpSpread, rowStart As Long)
    Dim a As Long
    a = 0

    With fps

        .sheet = 2

        'truong hop so dong lon hon 10000
        If numRow >= 10000 Then

            Do While a * 2 <= 1024

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), rowStart + a, .ColLetterToNumber("O"), (rowStart + a + 1)
                    

                    a = a + 2
                ElseIf a <> 0 Then
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), rowStart + a - 1, .ColLetterToNumber("O"), rowStart + a

                    a = a * 2
                End If

            Loop
                 
            
            Dim dem As Long
            Dim du  As Long
            dem = numRow \ 1024
            du = numRow Mod 1024
            a = 1
            If dem > 0 Then

                Do While a < dem
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), 1024 + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), 1024 + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), 1024 + rowStart - 1, .ColLetterToNumber("O"), rowStart + 1024 * a
                          
                    a = a + 1
                Loop
                .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), du + rowStart - 1, .ColLetterToNumber("O"), rowStart + 1024 * a
                
    
            Else
                .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), du + rowStart - 1, .ColLetterToNumber("O"), rowStart + 1024 * (a - 1)
                

            End If

            ' truong hop nho hon 1024000
        ElseIf numRow < 10000 Then
            a = 0

            Do While a * 2 < numRow

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), rowStart + a, .ColLetterToNumber("O"), (rowStart + a + 1)
                    
                    a = a + 2
                ElseIf a <> 0 Then
                    .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), rowStart + a - 1, .ColLetterToNumber("O"), rowStart + a
                    
                
                    a = a * 2
                End If

            Loop
            .CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + (numRow - a - 1), .ColLetterToNumber("B"), rowStart + a
            .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + (numRow - a - 1), .ColLetterToNumber("A"), rowStart + a
            .CopyRange .ColLetterToNumber("O"), rowStart, .ColLetterToNumber("O"), rowStart + (numRow - a - 1), .ColLetterToNumber("O"), rowStart + a
            

        End If
            
    End With

End Sub

Public Sub copyFormulas01_NTNN(numRow As Long, fps As fpSpread, rowStart As Long)
    Dim a As Long
    a = 0

    With fps

        .sheet = 1

        'truong hop so dong lon hon 10000
        If numRow >= 10000 Then

            Do While a * 2 <= 1024

                If a = 0 Then
                    
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
'                    .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), rowStart + a, .ColLetterToNumber("BW"), (rowStart + a + 1)
'                    .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), rowStart + a, .ColLetterToNumber("AU"), (rowStart + a + 1)
'                    .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), rowStart + a, .ColLetterToNumber("BW"), (rowStart + a + 1)
'
                    a = a + 2
                ElseIf a <> 0 Then
                    
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
'                    .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), rowStart + a - 1, .ColLetterToNumber("BW"), rowStart + a
'                    .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), rowStart + a - 1, .ColLetterToNumber("AU"), rowStart + a
'                    .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), rowStart + a - 1, .ColLetterToNumber("BW"), rowStart + a
'
                    a = a * 2
                End If

            Loop
                 
            a = 1
            Dim dem As Long
            Dim du  As Long
            dem = numRow \ 1024
            du = numRow Mod 1024

            If dem > 0 Then

                Do While a < dem
                    
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), 1024 + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
'                    .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), 1024 + rowStart - 1, .ColLetterToNumber("BW"), rowStart + 1024 * a
'                    .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), 1024 + rowStart - 1, .ColLetterToNumber("AU"), rowStart + 1024 * a
'                    .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), 1024 + rowStart - 1, .ColLetterToNumber("BM"), rowStart + 1024 * a
'
                    a = a + 1
                Loop
                
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
'                .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), du + rowStart - 1, .ColLetterToNumber("BW"), rowStart + 1024 * a
'                .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), du + rowStart - 1, .ColLetterToNumber("AU"), rowStart + 1024 * a
'                .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), du + rowStart - 1, .ColLetterToNumber("BM"), rowStart + 1024 * a
'
            Else
                
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * (a - 1)
'                .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), du + rowStart - 1, .ColLetterToNumber("BW"), rowStart + 1024 * (a - 1)
'                .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), du + rowStart - 1, .ColLetterToNumber("AU"), rowStart + 1024 * (a - 1)
'                .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), du + rowStart - 1, .ColLetterToNumber("BM"), rowStart + 1024 * (a - 1)
'
            End If

            ' truong hop nho hon 1024000
        ElseIf numRow < 10000 Then
            a = 0

            Do While a * 2 < numRow

                If a = 0 Then
                    
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
'                    .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), rowStart + a, .ColLetterToNumber("BW"), (rowStart + a + 1)
'                    .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), rowStart + a, .ColLetterToNumber("AU"), (rowStart + a + 1)
'                    .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), rowStart + a, .ColLetterToNumber("BM"), (rowStart + a + 1)
'
                    a = a + 2
                ElseIf a <> 0 Then
                    
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
'                    .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), rowStart + a - 1, .ColLetterToNumber("BW"), rowStart + a
'                    .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), rowStart + a - 1, .ColLetterToNumber("AU"), rowStart + a
'                    .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), rowStart + a - 1, .ColLetterToNumber("BM"), rowStart + a
'
                    a = a * 2
                End If

            Loop
            
            .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("BW"), rowStart + (numRow - a - 1), .ColLetterToNumber("A"), rowStart + a
'            .CopyRange .ColLetterToNumber("BW"), rowStart, .ColLetterToNumber("BW"), rowStart + (numRow - a - 1), .ColLetterToNumber("BW"), rowStart + a
'            .CopyRange .ColLetterToNumber("AU"), rowStart, .ColLetterToNumber("AU"), rowStart + (numRow - a - 1), .ColLetterToNumber("AU"), rowStart + a
'            .CopyRange .ColLetterToNumber("BM"), rowStart, .ColLetterToNumber("BW"), rowStart + (numRow - a - 1), .ColLetterToNumber("BM"), rowStart + a
'
        End If
            
    End With

End Sub

Public Sub moveDataNKH()
    Debug.Print "Total Time In: " & Time
    Dim xmlDocument     As New MSXML.DOMDocument
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim varMenuId       As String
    Dim rowStartSpread1 As Long
    Dim rowStartSpread2 As Long
    Dim i               As Long
    
    
    
    fpSpread1.EventEnabled(EventAllEvents) = False
          
    fpSpread1.Visible = False
    fpSpread2.Visible = True
    ProgressBar1.Visible = True
    ProgressBar1.value = 0

    DoEvents
    
    'fpSpread2.sheet = mCurrentSheet
    ' Lay ID cua Menu
    varMenuId = GetAttribute(TAX_Utilities_New.NodeValidity.parentNode, "ID")
    
    'Kiem tra neu to khai nha thau chi hien thi label status tai
    If Trim(varMenuId) = "70" Then
        Frame3.Visible = True
        txt_Seach.Visible = False
        Cb_seach.Visible = False
        Cmd_Seach.Visible = False
        Lbload.Visible = True
    End If

    If Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 2 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_05A_TNCN.xml"))
       
        rowStartSpread1 = 22
        rowStartSpread2 = 5
    ElseIf Trim(varMenuId) = "17" And fpSpread1.ActiveSheet = 3 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\BK_05B_TNCN.xml"))
 
        rowStartSpread1 = 22
        rowStartSpread2 = 4
    ElseIf Trim(varMenuId) = "59" And fpSpread1.ActiveSheet = 2 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\06_TNCN10.xml"))
 
        rowStartSpread1 = 22
        rowStartSpread2 = 3
        
    ElseIf Trim(varMenuId) = "70" And fpSpread1.ActiveSheet = 1 Then
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\PL_01_NTNN.xml"))
 
        rowStartSpread1 = 55
        rowStartSpread2 = 14
    End If
    
    fpSpread1.Row = rowStartSpread1
    
    Dim lrowCount As Long
    Dim varTemp, varTemp1, varTemp2  As Variant
    
        ' Xu ly truong hop nhap 1 dong ghi sau do tai du lieu
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim hasVl  As Integer
    Dim value As Variant
    Dim isFirstRown As Boolean
    

    ' Kiem tra tu dong maxrow len, neu gap bat ky mot dong nao bat dau co du lieu thi se lay do la maxrow luon
    For lrowCount = fpSpread2.MaxRows To 0 Step -1
        fpSpread2.GetText fpSpread2.ColLetterToNumber("B"), lrowCount, varTemp
        fpSpread2.GetText fpSpread2.ColLetterToNumber("F"), lrowCount, varTemp1
        ' Doi voi to khai 02/BH-TNCN va 02/XS-TNCN cot TN chiu thue la cot E
        fpSpread2.GetText fpSpread2.ColLetterToNumber("E"), lrowCount, varTemp2
        
        ' To khai 05/TNCN
        If Trim(varMenuId) = "17" Then
            If (Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "") And (Trim(varTemp1) <> vbNullString Or Trim(varTemp1) <> "") Then
                ' Tru tiep 4 dong header dau tien thi se duoc tong so dong can import vao
                If mCurrentSheet = 2 Then
                    lrowCount = lrowCount - 4
                ElseIf mCurrentSheet = 3 Then
                    lrowCount = lrowCount - 3
                End If
                Exit For
            End If
        End If
        
        ' To khai 02/BH-TNCN va to khai 02/XS-TNCN
        If Trim(varMenuId) = "59" Then
            If (Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "") And (Trim(varTemp2) <> vbNullString Or Trim(varTemp2) <> "") Then
                    If mCurrentSheet = 2 Then
                        lrowCount = lrowCount - 2
                    End If
                Exit For
            End If
        End If
        ' To khai 01/NTNN
        If ((Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "")) And Trim(varMenuId) = "70" Then
            
            If mCurrentSheet = 1 Then
                lrowCount = lrowCount - 13
                
            End If

            Exit For
        End If
    Next
    
    ' Truong hop them du lieu va xoa du lieu da ton tai
    If themXoaDuLieu Then
        ' dong dau tien luon la dong trang
        isFirstRown = True
        
        ResetData
        
'        ResetDataAndForm mCurrentSheet
    End If
    
    ' Truong hop them tiep du lieu
    Dim xmlSecionNode As MSXML.IXMLDOMNode
    Dim currentRow    As Long
    Dim varData1, varData2 As Variant
    

    ' To khai 01/NTNN them tu session 1
    If Trim(varMenuId) = "70" Then
         Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(1)
    Else
         Set xmlSecionNode = TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Section")(0)
    End If
    
   

    If themDuLieu Then

        If Not xmlSecionNode Is Nothing And GetAttribute(xmlSecionNode, "Dynamic") = "1" Then
            currentRow = xmlSecionNode.childNodes.length + fpSpread1.Row
        End If
    End If

    'Ca hai bang ke trong to quyet toan 5A bat dau tu dong 22, 5B bat dau tu dong 21
    If themDuLieu Then
        rowStartSpread1 = currentRow - 1
        ' Kiem tra du lieu dong dau tien neu du lieu khac rong thi insert tu dong tiep theo
        If xmlSecionNode.childNodes.length = 1 Then
            For Each xmlCellNode In xmlSecionNode.childNodes.Item(0).childNodes
                value = GetAttribute(xmlCellNode, "Value")
                If (GetAttribute(xmlCellNode, "FirstCell") <> "" And value <> "") Or (GetAttribute(xmlCellNode, "FirstCell") = "" And value <> "" And value <> "0") Then
                      hasVl = hasVl + 1
                End If
            Next
            ' truong hop dong dau tien trang
            If hasVl = 0 Then
                isFirstRown = True
            End If
        End If
    End If

    Debug.Print "COPY DATA IN : " & Time

    ' copy data vao Spread1

    ProgressBar1.max = lrowCount
    On Error GoTo ErrHandle

    If Trim(varMenuId) = "17" Then
        If fpSpread1.ActiveSheet = 2 Then

            gridData05A rowStartSpread1, rowStartSpread2, lrowCount, 2, isFirstRown

        ElseIf fpSpread1.ActiveSheet = 3 Then

            gridData05B rowStartSpread1, rowStartSpread2, lrowCount, 3, isFirstRown

        End If
    ElseIf Trim(varMenuId) = "59" Then
        gridData06TNCN rowStartSpread1, rowStartSpread2, lrowCount, 2
    ElseIf Trim(varMenuId) = "70" And fpSpread1.ActiveSheet = 1 Then
        gridData01NTNN rowStartSpread1, rowStartSpread2, lrowCount, 1
    End If

    Debug.Print "COPY DATA OUT: " & Time

        If strfileFont <> "UNICODE" Then
            If fpSpread1.ActiveSheet = 2 Then
                fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                fpSpread1.Row = rowStartSpread1

                Do
                    fpSpread1.Col = fpSpread1.ColLetterToNumber("D")

                    Select Case strfileFont

                        Case "TCVN"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, TCVN, UNICODE)

                        Case "VNI"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VNI, UNICODE)

                        Case "VIQR"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VIQR, UNICODE)

                        Case "VISCII"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VISCII, UNICODE)

                        Case Else
                            fpSpread1.Text = fpSpread1.Text
                    End Select

                    UpdateCell fpSpread1.ColLetterToNumber("D"), fpSpread1.Row, fpSpread1.Text

                    fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                    fpSpread1.Row = fpSpread1.Row + 1
                Loop While fpSpread1.Text = "aa"

            ElseIf fpSpread1.ActiveSheet = 3 Then
                fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                fpSpread1.Row = rowStartSpread1

                Do
                    fpSpread1.Col = fpSpread1.ColLetterToNumber("C")

                    Select Case strfileFont

                        Case "TCVN"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, TCVN, UNICODE)

                        Case "VNI"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VNI, UNICODE)

                        Case "VIQR"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VIQR, UNICODE)

                        Case "VISCII"
                            fpSpread1.Text = TAX_Utilities_New.Convert(fpSpread1.Text, VISCII, UNICODE)

                        Case Else
                            fpSpread1.Text = fpSpread1.Text
                    End Select

                    UpdateCell fpSpread1.ColLetterToNumber("C"), fpSpread1.Row, fpSpread1.Text
                    fpSpread1.Col = fpSpread1.ColLetterToNumber("B")
                    fpSpread1.Row = fpSpread1.Row + 1

                Loop While fpSpread1.Text = "aa"

            End If

        End If

        'Kiem tra neu to khai nha thau chi hien thi label status tai
    If Trim(varMenuId) = "70" Then
        Frame3.Visible = False
        txt_Seach.Visible = True
        Cb_seach.Visible = True
        Cmd_Seach.Visible = True
    End If

    
    fpSpread1.Visible = True
    ProgressBar1.Visible = False
    fpSpread1.EventEnabled(EventAllEvents) = True
    If Not objTaxBusiness Is Nothing Then objTaxBusiness.FinishImport
     Exit Sub
    Debug.Print "Total Time Out: " & Time
    

ErrHandle:
    SaveErrorLog Me.Name, "moveDataNKH", Err.Number, Err.Description
     Debug.Print "Erros: " & Err.Description
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       gridData05B
' Description:       1. Insert row them cac dong trong
'                    2. Set border cho grid
'                    3. copy du lieu Text tu spread2 sang spread1
'                    4. Set format cho Grid
'                    5. Copy Fomulas tu dong rowStartSpread1 cho cac dong con lai
' Created by :       nkhoan

' Parameters :       rowStartSpread1 : dong bat dau Spread1
'                    rowStartSpread2 : dong bat dau Spread2
'                    lrowCount : tong so dong can insert
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub gridData05B(rowStartSpread1 As Long, _
                        rowStartSpread2 As Long, _
                        lrowCount As Long, _
                        numSheet As Integer, isFirstRow As Boolean)
        
    ReDim fparray(lrowCount - 1, 6) As Variant
    fpSpread2.GetArray fpSpread2.ColLetterToNumber("B"), rowStartSpread2, fparray
    Dim a                As Long
    Dim rowStartSpread11 As Long
    a = 0
    rowStartSpread11 = rowStartSpread1
    On Error GoTo ErrHandle

    With fpSpread1
        .sheet = numSheet
        .EventEnabled(0) = False

        ' do hai phu luc A,B co dong bat dau 22, truong hop nay them du lieu vao grid da co du lieu
        If rowStartSpread1 > 22 Then
           
            .MaxRows = lrowCount + .MaxRows
            ' 1. Insert row them cac dong trong
            .InsertRows rowStartSpread1 + 1, lrowCount
           
            rowStartSpread11 = rowStartSpread1
        Else
            .MaxRows = lrowCount + .MaxRows - 1
            ' 1. Insert row them cac dong trong
            If isFirstRow = True Then
                
                .InsertRows rowStartSpread1 + 1, lrowCount - 1
            Else
                
                .InsertRows rowStartSpread1 + 1, lrowCount
            End If
        End If
        
        '2. Set border cho grid
        .SetCellBorder .ColLetterToNumber("B"), rowStartSpread1, .ColLetterToNumber("J"), (lrowCount + rowStartSpread1), 15, &O0, CellBorderStyleSolid
        ' set border cot Y
        .SetCellBorder .ColLetterToNumber("Y"), rowStartSpread1, .ColLetterToNumber("Y"), (lrowCount + rowStartSpread1), 15, &O0, CellBorderStyleSolid

        '        3. copy du lieu Text tu spread2 sang spread1
        fpSpread2.Row = rowStartSpread2
          If isFirstRow = False Then
            rowStartSpread1 = rowStartSpread1 + 1
        End If
        
        Do While fpSpread2.Row < lrowCount + 3
            DoEvents
            ProgressBar1.value = a
            
            fpSpread2.Row = rowStartSpread2
            
            .Row = rowStartSpread1
            .RowHeight(-2) = 14.5
            

            .Col = .ColLetterToNumber("C")
            .Text = fparray(a, 0)

            .Col = .ColLetterToNumber("D")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 1), IIf(InStr(1, fparray(a, 1), ".") <> 0, InStr(1, fparray(a, 1), ".") - 1, Len(fparray(a, 1))))  'Replace(fparray(a, 1), ".", "")
            Else
                .Text = fparray(a, 1)
            End If

            .Col = .ColLetterToNumber("E")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 2)) Then
                .Text = Left(fparray(a, 2), IIf(InStr(1, fparray(a, 2), ".") <> 0, InStr(1, fparray(a, 2), ".") - 1, Len(fparray(a, 2)))) 'Replace(fparray(a, 2), ".", "")
            Else
                .Text = fparray(a, 2)
            End If

            .Col = .ColLetterToNumber("F")
            .Text = fparray(a, 3)

            .Col = .ColLetterToNumber("G")

            If IsNumeric(fparray(a, 4)) Then
                If Val(fparray(a, 4)) > 0 Then
                    .Text = Round(fparray(a, 4))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("H")
            
            If IsNumeric(fparray(a, 5)) Then
                If Val(fparray(a, 5)) > 0 Then
                    .Text = Round(fparray(a, 5))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("I")
           
            If IsNumeric(fparray(a, 6)) Then
                If Val(fparray(a, 6)) > 0 Then
                    .Text = Round(fparray(a, 6))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
         
            .Col = .ColLetterToNumber("J")

'            If IsNumeric(fparray(a, 7)) Then
'                .Text = Round(fparray(a, 7))
'            Else
            .Text = 0
'            End If

            a = a + 1
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread2 = rowStartSpread2 + 1

        Loop
            
        ' Truong hop khai 1 dong
        If lrowCount = 1 Then
            DoEvents
            ProgressBar1.value = a
            
            fpSpread2.Row = rowStartSpread2
            
            .Row = rowStartSpread1
            .RowHeight(-2) = 14.5
            

            .Col = .ColLetterToNumber("C")
            .Text = fparray(a, 0)

            .Col = .ColLetterToNumber("D")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 1), IIf(InStr(1, fparray(a, 1), ".") <> 0, InStr(1, fparray(a, 1), ".") - 1, Len(fparray(a, 1))))  'Replace(fparray(a, 1), ".", "")
            Else
                .Text = fparray(a, 1)
            End If

            .Col = .ColLetterToNumber("E")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 2)) Then
                .Text = Left(fparray(a, 2), IIf(InStr(1, fparray(a, 2), ".") <> 0, InStr(1, fparray(a, 2), ".") - 1, Len(fparray(a, 2)))) 'Replace(fparray(a, 2), ".", "")
            Else
                .Text = fparray(a, 2)
            End If

            .Col = .ColLetterToNumber("F")
            .Text = fparray(a, 3)

            .Col = .ColLetterToNumber("G")

            If IsNumeric(fparray(a, 4)) Then
                If Val(fparray(a, 4)) > 0 Then
                    .Text = Round(fparray(a, 4))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("H")
            
            If IsNumeric(fparray(a, 5)) Then
                If Val(fparray(a, 5)) > 0 Then
                    .Text = Round(fparray(a, 5))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("I")
           
            If IsNumeric(fparray(a, 6)) Then
                If Val(fparray(a, 6)) > 0 Then
                    .Text = Round(fparray(a, 6))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
         
            .Col = .ColLetterToNumber("J")

'            If IsNumeric(fparray(a, 7)) Then
'                .Text = Round(fparray(a, 7))
'            Else
            .Text = 0
'            End If

            a = a + 1
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread2 = rowStartSpread2 + 1
        End If
        ' 4. Set format cho Grid
        
        'format chi tieu [8],[D]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("D")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("D")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .BlockMode = False
        
        'format chi tieu [9],[E]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("E")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .TypeMaxEditLen = 60
        .BlockMode = False
        
        'format chi tieu [8]- [9],[D]->[E]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("D")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .TypeHAlign = TypeHAlignLeft
        .BlockMode = False

        'format chi tieu [10]- cot [F]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("F")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("F")
        .BlockMode = True
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False

        'format tu chi tieu [11] den [14],[G]->[J]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("G")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("J")
        .BlockMode = True
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 0
        .TypeNumberSeparator = "."
        .TypeNumberShowSep = True
        
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False

        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("J")
        .BlockMode = True
        .FontSize = 8
        .Lock = False
        .BlockMode = False
        
        ' set lock cot J
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("J")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("J")
        .BlockMode = True
        .Lock = True
        .BlockMode = False
        
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("Y")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("Y")
        .BlockMode = True
        .CellType = CellTypeCheckBox
        .Lock = False
        .BlockMode = False
        
        '5. Copy Fomulas tu dong rowStartSpread1 cho cac dong con lai
        
            If rowStartSpread11 > 22 Then
                copyFormulasSheet3 lrowCount + 1, fpSpread1, rowStartSpread11
            Else
                If isFirstRow = False Then
                    copyFormulasSheet3 lrowCount + 1, fpSpread1, rowStartSpread11
                Else
                    If lrowCount > 1 Then
                        copyFormulasSheet3 lrowCount, fpSpread1, rowStartSpread11
                    End If
                End If
            End If
            
     

        .EventEnabled(0) = True
'        .AutoCalc = False
'        .ReCalc
'        .AutoCalc = True
    End With

    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "moveDataNKH", Err.Number, Err.Description
    Debug.Print "Erros: " & Err.Description
End Sub

Private Sub gridData05A(rowStartSpread1 As Long, _
                        rowStartSpread2 As Long, _
                        lrowCount As Long, _
                        numSheet As Integer, isFirstRow As Boolean)
    ReDim fparray(lrowCount - 1, 10) As Variant
    fpSpread2.GetArray fpSpread2.ColLetterToNumber("B"), rowStartSpread2, fparray
    Dim a                As Long
    Dim rowStartSpread11 As Long
    a = 0
    rowStartSpread11 = rowStartSpread1

    With fpSpread1
        .sheet = numSheet
        .EventEnabled(0) = False

        ' do hai phu luc A,B co dong bat dau 22, truong hop nay them du lieu vao grid da co du lieu
        If rowStartSpread1 > 22 Then
           
            .MaxRows = lrowCount + .MaxRows
            ' 1. Insert row them cac dong trong
            .InsertRows rowStartSpread1 + 1, lrowCount
          
            rowStartSpread11 = rowStartSpread1
        Else
            .MaxRows = lrowCount + .MaxRows - 1
            ' 1. Insert row them cac dong trong
            If isFirstRow = True Then
                .InsertRows rowStartSpread1 + 1, lrowCount - 1
            Else
                
                .InsertRows rowStartSpread1 + 1, lrowCount
            End If
        End If

        '2. Set border cho grid
        .SetCellBorder .ColLetterToNumber("B"), rowStartSpread1, .ColLetterToNumber("S"), (lrowCount + rowStartSpread1), 15, &O0, CellBorderStyleSolid
        
        '3. copy du lieu Text tu spread2 sang spread1
        fpSpread2.Row = rowStartSpread2
        If isFirstRow = False Then
            rowStartSpread1 = rowStartSpread1 + 1
        End If

        Do While fpSpread2.Row < lrowCount + 4
            DoEvents
            ProgressBar1.value = a
            fpSpread2.Row = rowStartSpread2
            
            .Row = rowStartSpread1
            .RowHeight(-2) = 14.5
            
            .Col = .ColLetterToNumber("D")
                
            .Text = fparray(a, 0)
            .Col = .ColLetterToNumber("E")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 1), IIf(InStr(1, fparray(a, 1), ".") <> 0, InStr(1, fparray(a, 1), ".") - 1, Len(fparray(a, 1)))) 'Replace(fparray(a, 1), ".", "")
            Else
                .Text = fparray(a, 1)
            End If
        
            .Col = .ColLetterToNumber("F")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 2)) Then
                .Text = Left(fparray(a, 2), IIf(InStr(1, fparray(a, 2), ".") <> 0, InStr(1, fparray(a, 2), ".") - 1, Len(fparray(a, 2))))   'Replace(fparray(a, 2), ".", "")
            Else
                .Text = fparray(a, 2)
            End If
                                        
            .Col = .ColLetterToNumber("G")
            .Text = fparray(a, 3)
                        
            .Col = .ColLetterToNumber("H")

            If IsNumeric(fparray(a, 4)) Then
                If Val(fparray(a, 4)) > 0 Then
                    .Text = Round(fparray(a, 4))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                       
            .Col = .ColLetterToNumber("I")

            If IsNumeric(fparray(a, 5)) Then
                If Val(fparray(a, 5)) > 0 Then
                    .Text = Round(fparray(a, 5))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                                           
            .Col = .ColLetterToNumber("J")

            If IsNumeric(fparray(a, 6)) Then
                If Val(fparray(a, 6)) > 0 Then
                    .Text = Round(fparray(a, 6))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
        
            .Col = .ColLetterToNumber("K")

            If IsNumeric(fparray(a, 7)) Then
                If Val(fparray(a, 7)) > 0 Then
                    .Text = Round(fparray(a, 7))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                              
            .Col = .ColLetterToNumber("L")

            If IsNumeric(fparray(a, 8)) Then
                If Val(fparray(a, 8)) > 0 Then
                    .Text = Round(fparray(a, 8))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("N")
            
            If IsNumeric(fparray(a, 10)) Then
                If Val(fparray(a, 10)) > 0 Then
                    .Text = Round(fparray(a, 10))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
            
            .Col = .ColLetterToNumber("O")
            .Text = 0

            a = a + 1
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread2 = rowStartSpread2 + 1

        Loop
        ' Truong hop tai 1 dong du lieu
        If lrowCount = 1 Then
            DoEvents
            ProgressBar1.value = a
            fpSpread2.Row = rowStartSpread2
            
            .Row = rowStartSpread1
            .RowHeight(-2) = 14.5
            
            .Col = .ColLetterToNumber("D")
                
            .Text = fparray(a, 0)
            .Col = .ColLetterToNumber("E")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 1), IIf(InStr(1, fparray(a, 1), ".") <> 0, InStr(1, fparray(a, 1), ".") - 1, Len(fparray(a, 1)))) 'Replace(fparray(a, 1), ".", "")
            Else
                .Text = fparray(a, 1)
            End If
        
            .Col = .ColLetterToNumber("F")
            ' Replace dau "." doi voi cac truong hop format khong co not comment mau xanh tren file excel
            If Not IsNull(fparray(a, 2)) Then
                .Text = Left(fparray(a, 2), IIf(InStr(1, fparray(a, 2), ".") <> 0, InStr(1, fparray(a, 2), ".") - 1, Len(fparray(a, 2))))   'Replace(fparray(a, 2), ".", "")
            Else
                .Text = fparray(a, 2)
            End If
                                        
            .Col = .ColLetterToNumber("G")
            .Text = fparray(a, 3)
                        
            .Col = .ColLetterToNumber("H")

            If IsNumeric(fparray(a, 4)) Then
                If Val(fparray(a, 4)) > 0 Then
                    .Text = Round(fparray(a, 4))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                       
            .Col = .ColLetterToNumber("I")

            If IsNumeric(fparray(a, 5)) Then
                If Val(fparray(a, 5)) > 0 Then
                    .Text = Round(fparray(a, 5))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                                           
            .Col = .ColLetterToNumber("J")

            If IsNumeric(fparray(a, 6)) Then
                If Val(fparray(a, 6)) > 0 Then
                    .Text = Round(fparray(a, 6))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
        
            .Col = .ColLetterToNumber("K")

            If IsNumeric(fparray(a, 7)) Then
                If Val(fparray(a, 7)) > 0 Then
                    .Text = Round(fparray(a, 7))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                              
            .Col = .ColLetterToNumber("L")

            If IsNumeric(fparray(a, 8)) Then
                If Val(fparray(a, 8)) > 0 Then
                    .Text = Round(fparray(a, 8))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("N")
            
            If IsNumeric(fparray(a, 10)) Then
                If Val(fparray(a, 10)) > 0 Then
                    .Text = Round(fparray(a, 10))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
            
            .Col = .ColLetterToNumber("O")
            .Text = 0

            a = a + 1
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread2 = rowStartSpread2 + 1
        End If
        ' end
'               4. Set format cho Grid

'               'format max lenght cot [E]
                .Row = rowStartSpread11
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("C")
        .BlockMode = True
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False

        '               'format max lenght cot [E]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("E")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .BlockMode = False
                
                'format max lenght cot [F]
                .Row = rowStartSpread11
                .Col = .ColLetterToNumber("F")
                .Row2 = lrowCount + rowStartSpread11
                .Col2 = .ColLetterToNumber("F")
                .BlockMode = True
                .TypeMaxEditLen = 60
                .BlockMode = False

'               format chi tieu [8]- [9],[E]->[F]
                .Row = rowStartSpread11
                .Col = .ColLetterToNumber("E")
                .Row2 = lrowCount + rowStartSpread11
                .Col2 = .ColLetterToNumber("F")
                .BlockMode = True
                .TypeHAlign = TypeHAlignLeft
                .BlockMode = False
        
                'format chi tieu [10]- cot [G]
                .Row = rowStartSpread11
                .Col = .ColLetterToNumber("G")
                .Row2 = lrowCount + rowStartSpread11
                .Col2 = .ColLetterToNumber("G")
                .BlockMode = True
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .BlockMode = False
        
        'format tu chi tieu [11] den [21],[H]->[S]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("H")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("S")
        .BlockMode = True
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 0
        .TypeNumberSeparator = "."
        .TypeNumberShowSep = True
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False

        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("S")
        .BlockMode = True
        .TypeVAlign = TypeVAlignCenter
        .FontSize = 8
        .FontName = "Tahoma"
        .Lock = False
        .BlockMode = False
        
        ' set lock cot O
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("O")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("O")
        .BlockMode = True
        .Lock = True
        .BlockMode = False
        
        ' set lock cot R, S
         .Row = rowStartSpread11
        .Col = .ColLetterToNumber("R")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("S")
        .BlockMode = True
        .Lock = True
        .BlockMode = False
        
        '5. Copy Fomulas tu dong rowStartSpread1 cho cac dong con lai
        If rowStartSpread11 > 22 Then
            copyFormulasSheet2 lrowCount + 1, fpSpread1, rowStartSpread11
        Else
            If isFirstRow = False Then
                copyFormulasSheet2 lrowCount + 1, fpSpread1, rowStartSpread11
            Else
                If lrowCount > 1 Then
                    copyFormulasSheet2 lrowCount, fpSpread1, rowStartSpread11
                End If
            End If
        End If

       .EventEnabled(0) = True
       
'    .AutoCalc = False
'    .ReCalc
'    .AutoCalc = True
        
    End With

End Sub

Private Sub gridData06TNCN(rowStartSpread1 As Long, _
                           rowStartSpread2 As Long, _
                           lrowCount As Long, _
                           numSheet As Integer)
                           
    ReDim fparray(lrowCount, 4) As Variant
    fpSpread2.GetArray fpSpread2.ColLetterToNumber("B"), rowStartSpread2, fparray
    Dim a                As Integer
    Dim rowStartSpread11 As Long
    a = 0
    rowStartSpread11 = rowStartSpread1

    With fpSpread1
        .sheet = numSheet
        .EventEnabled(0) = False

        ' do hai phu luc A,B co dong bat dau 22, truong hop nay them du lieu vao grid da co du lieu
        If rowStartSpread1 > 22 Then
           
            .MaxRows = lrowCount + .MaxRows
            ' 1. Insert row them cac dong trong
            .InsertRows rowStartSpread1 + 1, lrowCount
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread11 = rowStartSpread1
        Else
            .MaxRows = lrowCount + .MaxRows - 1
            ' 1. Insert row them cac dong trong
            .InsertRows rowStartSpread1 + 1, lrowCount - 1
        End If

        '2. Set border cho grid
        .SetCellBorder .ColLetterToNumber("B"), rowStartSpread1, .ColLetterToNumber("H"), (lrowCount + rowStartSpread1), 15, &O0, CellBorderStyleSolid
        
        '3. copy du lieu Text tu spread2 sang spread1
        fpSpread2.Row = rowStartSpread2

        Do While fpSpread2.Row < lrowCount + 2
            DoEvents
            ProgressBar1.value = a
            fpSpread2.Row = rowStartSpread2
            
            .Row = rowStartSpread1
            .RowHeight(-2) = 14.5
            

            .Col = .ColLetterToNumber("D")
                
            .Text = fparray(a, 0)
            .Col = .ColLetterToNumber("E")
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 1), IIf(InStr(1, fparray(a, 1), ".") <> 0, InStr(1, fparray(a, 1), ".") - 1, Len(fparray(a, 1))))
            Else
                .Text = fparray(a, 1)
            End If
        
            .Col = .ColLetterToNumber("F")
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 2), IIf(InStr(1, fparray(a, 2), ".") <> 0, InStr(1, fparray(a, 2), ".") - 1, Len(fparray(a, 2))))
            Else
                .Text = fparray(a, 2)
            End If
                                        
            .Col = .ColLetterToNumber("G")

            If IsNumeric(fparray(a, 3)) Then
                If Val(fparray(a, 3)) > 0 Then
                    .Text = Round(fparray(a, 3))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                        
            .Col = .ColLetterToNumber("H")

            If IsNumeric(fparray(a, 4)) Then
                If Val(fparray(a, 4)) > 0 Then
                    .Text = Round(fparray(a, 4))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If
                       
            

            a = a + 1
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread2 = rowStartSpread2 + 1

        Loop
            
        '               4. Set format cho Grid
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("C")
        .BlockMode = True
         .CellType = CellTypeCheckBox
        .BlockMode = False

        '               'format max lenght cot [E]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("E")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .BlockMode = False
                
        'format max lenght cot [F]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("F")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("F")
        .BlockMode = True
        .TypeMaxEditLen = 60
        .BlockMode = False

        '               format chi tieu [8]- [9],[E]->[F]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("E")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("F")
        .BlockMode = True
        .TypeHAlign = TypeHAlignLeft
        .BlockMode = False
        

        'format tu chi tieu [10] den [11],[G]->[H]
        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("G")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("H")
        .BlockMode = True
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 0
        .TypeNumberSeparator = "."
        .TypeNumberShowSep = True
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False

        .Row = rowStartSpread11
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount + rowStartSpread11
        .Col2 = .ColLetterToNumber("H")
        .BlockMode = True
        .TypeVAlign = TypeVAlignCenter
        .FontSize = 8
        .FontName = "Tahoma"
        .Lock = False
        .BlockMode = False
        
        '5. Copy Fomulas tu dong rowStartSpread1 cho cac dong con lai
        If rowStartSpread11 > 22 Then
            copyFormulas06_TNCN lrowCount + 1, fpSpread1, rowStartSpread11 - 1
        Else
            copyFormulas06_TNCN lrowCount, fpSpread1, rowStartSpread11
        End If

        .EventEnabled(0) = True
       
    End With
                           
                              
End Sub
Private Sub gridData01NTNN(rowStartSpread1 As Long, _
                           rowStartSpread2 As Long, _
                           lrowCount As Long, _
                           numSheet As Integer)
                           
    ReDim fparray(lrowCount, 11) As Variant
    fpSpread2.GetArray fpSpread2.ColLetterToNumber("B"), rowStartSpread2, fparray
    Dim a                As Integer
    Dim rowStartSpread11 As Long
    a = 0
    rowStartSpread11 = rowStartSpread1

    With fpSpread1
        .sheet = numSheet
        .EventEnabled(0) = False

        ' do hai phu luc A,B co dong bat dau 22, truong hop nay them du lieu vao grid da co du lieu
        If rowStartSpread1 > 55 Then
           
            .MaxRows = lrowCount + .MaxRows
            ' 1. Insert row them cac dong trong
            .InsertRows rowStartSpread1 + 1, lrowCount
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread11 = rowStartSpread1
        Else
            .MaxRows = lrowCount + .MaxRows - 1
            ' 1. Insert row them cac dong trong
            .InsertRows rowStartSpread1 + 1, lrowCount - 1
        End If

        '2. Set border cho grid
        ' .SetCellBorder .ColLetterToNumber("C"), rowStartSpread1, .ColLetterToNumber("BQ"), (lrowCount + rowStartSpread1), 15, &O0, CellBorderStyleSolid
        
        '5. Copy Fomulas tu dong rowStartSpread1 cho cac dong con lai
        If rowStartSpread11 > 55 Then
            copyFormulas01_NTNN lrowCount + 1, fpSpread1, rowStartSpread11 - 1
        Else
            copyFormulas01_NTNN lrowCount, fpSpread1, rowStartSpread11
        End If

        .EventEnabled(0) = True
        
        '3. copy du lieu Text tu spread2 sang spread1
        fpSpread2.Row = rowStartSpread2
        
        Dim arrStr() As String
        Dim sDate    As String

        Do While fpSpread2.Row < lrowCount + 13
            DoEvents
            ProgressBar1.value = a
            fpSpread2.Row = rowStartSpread2

            .RowHeight(-2) = 14.5
            .Row = rowStartSpread1

            .Col = .ColLetterToNumber("C")
            .Text = fparray(a, 0)

            .Col = .ColLetterToNumber("L")
            If Not IsNull(fparray(a, 1)) Then
                .Text = Left(fparray(a, 1), IIf(InStr(1, fparray(a, 1), ".") <> 0, InStr(1, fparray(a, 1), ".") - 1, Len(fparray(a, 1))))
            Else
                .Text = fparray(a, 1)
            End If
            .Col = .ColLetterToNumber("R")
            .Text = fparray(a, 2)

            .Col = .ColLetterToNumber("X")

            If IsNumeric(fparray(a, 3)) Then
                If Val(fparray(a, 3)) > 0 Then
                    .Text = Round(fparray(a, 3))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("AD")

'            If InStr(1, fparray(a, 4), "-") <> 0 Then
'                arrStr = Split(fparray(a, 4), "-")
'            Else
'                arrStr = Split(fparray(a, 4), "/")
'            End If
'
'            sDate = Right("00" & arrStr(0), 2) & "/" & Right("00" & arrStr(1), 2) & "/" & Right("20" & arrStr(2), 4)
            
            .Text = fparray(a, 4)  'sDate

            .Col = .ColLetterToNumber("AI")

            If IsNumeric(fparray(a, 5)) Then
                If Val(fparray(a, 5)) > 0 Then
                    .Text = Round(fparray(a, 5))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("AM")

            If IsNumeric(fparray(a, 6)) And CDbl(fparray(a, 6)) < 100 Then
                If Val(fparray(a, 6)) > 0 Then
                    .Text = Round(fparray(a, 6), 2)
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("AQ")

            If IsNumeric(fparray(a, 7)) And CDbl(fparray(a, 7)) < 100 Then
                If Val(fparray(a, 7)) > 0 Then
                    .Text = Round(fparray(a, 7))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("AY")

            If IsNumeric(fparray(a, 9)) Then
                If Val(fparray(a, 9)) > 0 Then
                    .Text = Round(fparray(a, 9))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("BC")

            If IsNumeric(fparray(a, 10)) And CDbl(fparray(a, 10)) < 100 Then
                If Val(fparray(a, 10)) > 0 Then
                    .Text = Round(fparray(a, 10), 2)
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            .Col = .ColLetterToNumber("BG")

            If IsNumeric(fparray(a, 11)) Then
                If Val(fparray(a, 11)) > 0 Then
                    .Text = Round(fparray(a, 11))
                Else
                    .Text = 0
                End If
            Else
                .Text = 0
            End If

            a = a + 1
            rowStartSpread1 = rowStartSpread1 + 1
            rowStartSpread2 = rowStartSpread2 + 1

        Loop
            
        '4. Set format cho Grid

        '               'format max lenght cot [L]
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("L")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("L")
        '        .BlockMode = True
        '        .TypeMaxEditLen = 14
        '        .TypeHAlign = TypeHAlignCenter
        '        .BlockMode = False
        '
        '        'format max lenght cot [R]
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("R")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("R")
        '        .BlockMode = True
        '        .TypeHAlign = TypeHAlignRight
        '        .TypeMaxEditLen = 50
        '        .BlockMode = False
        '
        '        'format max lenght cot [AD]
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("AD")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("AD")
        '        .BlockMode = True
        '        .TypeHAlign = TypeHAlignCenter
        '        .BlockMode = False
        '
        '        'format max lenght cot [AQ]
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("AQ")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("AQ")
        '        .BlockMode = True
        '        .TypeMaxEditLen = 2
        '        .BlockMode = False
        '
        '
        '
        '        ' cot 4
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("X")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("X")
        '        .BlockMode = True
        '        .CellType = CellTypeNumber
        '        .TypeNumberDecPlaces = 0
        '        .TypeNumberSeparator = "."
        '        .TypeNumberShowSep = True
        '        .TypeHAlign = TypeHAlignRight
        '        .BlockMode = False
        '         ' cot 6 - 14
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("AI")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("BQ")
        '        .BlockMode = True
        '        .CellType = CellTypeNumber
        '        .TypeNumberDecPlaces = 0
        '        .TypeNumberSeparator = "."
        '        .TypeNumberShowSep = True
        '        .TypeHAlign = TypeHAlignRight
        '        .BlockMode = False
        '
        '
        '        .Row = rowStartSpread11
        '        .Col = .ColLetterToNumber("C")
        '        .Row2 = lrowCount + rowStartSpread11
        '        .Col2 = .ColLetterToNumber("BQ")
        '        .BlockMode = True
        '        .TypeVAlign = TypeVAlignCenter
        '        .FontSize = 8
        '        .FontName = "Tahoma"
        '        .Lock = False
        '        .BlockMode = False
       
    End With
                              
End Sub

' ham get formula tinh so tien phat nop cham
Private Function getFormulaTienPNC(t As Long, soTien As Double, strColRow As String) As String
    Dim soNgayNopCham As Long
    Dim soNgayNopChamTruocHl As Long
    Dim arrDate() As String
    Dim dHanNop As Date
    Dim dNgayBs As Date
    Dim dHieuLuc As Date
    
    Dim result As String
    
    soNgayNopCham = getSoNgay(hanNopTk, ngayLapTkBs)
    soNgayNopChamTruocHl = getSoNgay(hanNopTk, "01/07/2013") - 1
    If hanNopTk <> "" Then
        arrDate = Split(hanNopTk, "/")
        dHanNop = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
    End If
    
    If ngayLapTkBs <> "" Then
        arrDate = Split(ngayLapTkBs, "/")
        dNgayBs = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
    End If
    
    dHieuLuc = DateSerial(2013, 7, 1)
    If DateDiff("D", dHanNop, dHieuLuc) > 0 And DateDiff("D", dNgayBs, dHieuLuc) < 0 Then
        ' neu ngay phat sinh khoan no truoc 01/07/2013
        If soNgayNopCham - soNgayNopChamTruocHl <= 90 Then
            result = soNgayNopCham & "*" & strColRow & "* 0.05 / 100"
        Else
            result = (soNgayNopChamTruocHl + 90) & "*" & strColRow & "* 0.05 / 100 +" & (soNgayNopCham - soNgayNopChamTruocHl - 90) & "*" & strColRow & "* 0.07 / 100"
        End If
    Else
        ' neu ngay phat sinh khoan no sau 01/07/2013
        If soNgayNopCham <= 90 Then
            result = soNgayNopCham & "*" & strColRow & "*0.05/100"
        Else
            result = 90 & "*" & strColRow & "*0.05/100+" & (soNgayNopCham - 90) & "*" & strColRow & "*0.07/100"
        End If
    End If
    getFormulaTienPNC = "IF(" & result & ">0;ROUND(" & result & ";0);0)"  'result
    Exit Function
End Function
