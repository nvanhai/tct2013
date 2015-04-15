VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmAddSheet 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4815
   ControlBox      =   0   'False
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
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
      TabIndex        =   4
      Top             =   1890
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
      TabIndex        =   3
      Top             =   1890
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   30
      TabIndex        =   5
      Top             =   450
      Width           =   4755
      Begin VB.CheckBox chkSelectAll 
         Height          =   195
         HelpContextID   =   81211
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   195
      End
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   825
         Left            =   60
         TabIndex        =   2
         Top             =   480
         Width           =   4635
         _Version        =   458752
         _ExtentX        =   8176
         _ExtentY        =   1455
         _StockProps     =   64
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
         MaxCols         =   1
         MaxRows         =   1
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmAddSheet.frx":0000
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
         TabIndex        =   1
         Top             =   210
         Width           =   2355
      End
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Thªm phô lôc kª khai"
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
      TabIndex        =   6
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
Attribute VB_Name = "frmAddSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Center Name       : FIS (Financial Insurance Solution)
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmAddSheet
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
Private strSelectedSheets As String
Private strSheets As String
Private blnFPChange As Boolean

Public Function SheetSelections(ByVal strAllSheet As String, ByVal strCurrentSheets As String) As String
    strSheets = strAllSheet
    strSelectedSheets = strCurrentSheets
    
    Me.Show vbModal
    
    SheetSelections = strSelectedSheets
    
End Function

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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ' BCTC kiem tra chi chon PL LCTTGT hoac LCTTTT
    Dim idxPL As Long
    Dim countPL As Integer
    If TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "69" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "19" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "20" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "22" Then
        With fpSpread1
            .Col = 1
            For idxPL = 2 To .MaxRows
                .Row = idxPL
                If (.value = 1 Or .value = 2) And (idxPL = 3 Or idxPL = 4) Then
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
    ' end
    GetSelectedSheets
    Unload Me
End Sub

Private Sub Form_Load()
    LoadGrid
End Sub

Private Sub GetSelectedSheets()
    Dim intCtrl As Integer
    With fpSpread1
        .Col = 1
        For intCtrl = 1 To .MaxRows
            .Row = intCtrl
            If .value = 1 Then
                strSelectedSheets = strSelectedSheets & "," & .TypeCheckText
            End If
        Next intCtrl
    End With
End Sub

Private Sub LoadGrid()
    Dim intCtrl As Integer
    Dim arrStrSheetNames() As String
    
    arrStrSheetNames = Split(strSheets, ",")
    With fpSpread1
        .SheetCount = 1
        .MaxCols = 1
        .MaxRows = UBound(arrStrSheetNames) + 1
        .NoBeep = True
        .TabStripPolicy = TabStripPolicyNever
        .RowHeadersShow = False
        .ColHeadersShow = False
        .ScrollBars = ScrollBarsVertical
        .EditModePermanent = True
        .CursorStyle = CursorStyleArrow
        .EventEnabled(EventButtonClicked) = False
        .Col = 1
        .ColWidth(1) = 37
        For intCtrl = 1 To .MaxRows
            .Row = intCtrl
            .CellType = CellTypeCheckBox
            .TypeCheckText = arrStrSheetNames(intCtrl - 1)
            If InStr(1, "," & strSelectedSheets & ",", "," & arrStrSheetNames(intCtrl - 1) & ",") <> 0 Then
                .TypeCheckType = TypeCheckTypeThreeState
                .value = 2
                .Lock = True
            Else
                .TypeCheckType = TypeCheckTypeNormal
                .value = 0
            End If
            .RowHeight(intCtrl) = .MaxTextRowHeight(intCtrl)
        Next intCtrl
        
        'An sheet to khai
        .Row = 1
        .RowHidden = True
        .EventEnabled(EventButtonClicked) = True
        
        'Call event for chkSelectAll
        blnFPChange = True
        fpSpread1_ButtonClicked 1, 1, 1
        blnFPChange = False
    End With
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    
End Sub

Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim lCtrl As Long
    Dim intSelectAll As Integer
    'dhdang sua
    Dim rowcheck As Integer
    
    If Not blnFPChange Then
        Exit Sub
    End If
    
    intSelectAll = 2
    fpSpread1.Col = 1
    For lCtrl = 2 To fpSpread1.MaxRows
        fpSpread1.Row = lCtrl
        If fpSpread1.value = 0 Then
            intSelectAll = 0
            Exit For
        ElseIf fpSpread1.value = 1 Then
            intSelectAll = 1
        End If
    Next lCtrl
    
    ' dhdang
    ' BC26 chi cho phep chon 1 phu luc
    For lCtrl = 2 To fpSpread1.MaxRows
        fpSpread1.Row = lCtrl
        If fpSpread1.value = 1 Then
            rowcheck = lCtrl
        ElseIf fpSpread1.value = 2 Then
            rowcheck = -1
        End If
    Next lCtrl

    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "68" Then
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

Private Sub fpSpread1_GotFocus()
    blnFPChange = True
End Sub

Private Sub fpSpread1_LostFocus()
    blnFPChange = False
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
    If chkSelectAll.Enabled Then
        chkSelectAll.SetFocus
    Else
        cmdOK.SetFocus
    End If
End Sub
