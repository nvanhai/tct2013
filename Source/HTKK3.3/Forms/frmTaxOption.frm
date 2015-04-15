VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTaxOption 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   0
      TabIndex        =   3
      Top             =   300
      Width           =   4935
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   1575
         Left            =   60
         TabIndex        =   2
         Top             =   510
         Width           =   4815
         _Version        =   458752
         _ExtentX        =   8493
         _ExtentY        =   2778
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmTaxOption.frx":0000
      End
      Begin MSForms.Label lblSelection 
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   150
         Width           =   4815
         Size            =   "8493;450"
         FontName        =   "Tahoma"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.Label lblCaption 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   2295
      ForeColor       =   -2147483634
      Size            =   "4048;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2490
      Width           =   1305
      Caption         =   "Exit"
      Size            =   "2302;661"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2490
      Width           =   1305
      Caption         =   "OK"
      Size            =   "2302;661"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmTaxOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2) - 200
    
    SetControlCaption Me
    SetDefaultActiveProperties
    LoadGrid
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
    frmTreeviewMenu.Show
End Sub

Private Sub cmdOK_Click()
    SetActiveValue
    Unload Me
    frmInterfaces.Show
End Sub

Private Sub Form_Resize()
    'SetFormCaption Me, imgCaption, lblCaption
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
    SetAttribute TAX_Utilities.NodeValidity.childNodes(lCtrl - 1), "Active", fpSpread1.Value
Next lCtrl
End Sub

'*****************************************************
'Description: LoadGrid procedure initialize and setup value to grid
'Author:ThanhDX
'Date:10/11/2005
'Input:
'Output:
'Return:
'*****************************************************
Private Sub LoadGrid()
Dim xmlNode As MSXML.IXMLDOMNode
Dim fso As New FileSystemObject
Dim lCtrl As Long, lRow As Long
Dim strDataFileName As String
Dim bNewData As Boolean

On Error GoTo ErrHandle
    lRow = 1
    bNewData = True
    
    With fpSpread1
        .col = 1
        For Each xmlNode In TAX_Utilities.NodeValidity.childNodes
                ' Get name of data file
                If Val(TAX_Utilities.month) <> 0 Then
                    strDataFileName = GetAttribute(xmlNode, "Folder") & GetAttribute(xmlNode, "DataFile") & "_" & TAX_Utilities.month & TAX_Utilities.Year & ".xml"
                Else
                    strDataFileName = GetAttribute(xmlNode, "Folder") & GetAttribute(xmlNode, "DataFile") & "_0" & TAX_Utilities.ThreeMonths & TAX_Utilities.Year & ".xml"
                End If
                
                'By default number of row is one
                'If it has more than one row
                If lRow > 1 Then 'Insert new row
                    .MaxRows = .MaxRows + 1
                    .InsertRows lRow, 1
                End If
                
                .Row = lRow
                .CellType = CellTypeCheckBox
                .TypeCheckType = TypeCheckTypeNormal
                .TypeCheckTextAlign = TypeCheckTextAlignRight
                .TypeCheckText = GetAttribute(xmlNode, "Caption")
                
                ' Check the exist of data file -> Set value to Checkbox
                If fso.FileExists(strDataFileName) Then
                    .TypeCheckType = TypeCheckTypeThreeState
                    .Value = 2
                    .Lock = True
                    bNewData = False
                End If
                'Resize row height: Auto fit with content
                .RowHeight(lRow) = .MaxTextRowHeight(lRow)
            
            lRow = lRow + 1
        Next
        
        'New tax -> Set default value by Menu.xml
        If bNewData Then
            For lCtrl = 1 To lRow - 1
                .Row = lCtrl
                .Value = GetAttribute(TAX_Utilities.NodeValidity.childNodes(lCtrl - 1), "Active")
            Next lCtrl
            'Do not allow first row to edit
            .Row = 1
            .TypeCheckType = TypeCheckTypeThreeState
            .Lock = True
            .Value = 2
        End If
        
        'Set cursor style and edit mode to fpSpread
        .CursorStyle = CursorStyleArrow
        .EditModePermanent = True
        .GrayAreaBackColor = vbButtonFace
    End With
Exit Sub

ErrHandle:
    SaveErrorLog Me.Name, "LoadGrid", Err.Number, Err.Description
End Sub

Private Sub SetDefaultActiveProperties()
Dim xmlDom As New MSXML.DOMDocument
Dim xmlNodeMenu As MSXML.IXMLDOMNode, xmlNodeValidity As MSXML.IXMLDOMNode
Dim strTemp As String, lCtrl As Long

xmlDom.Load App.path & "\Menu.xml"

strTemp = GetAttribute(TAX_Utilities.NodeMenu, "ID")
For Each xmlNodeMenu In xmlDom.getElementsByTagName("Menu")
    If GetAttribute(xmlNodeMenu, "ID") = strTemp Then _
        Exit For
Next

strTemp = GetAttribute(TAX_Utilities.NodeValidity, "StartDate")
For Each xmlNodeValidity In xmlNodeMenu.childNodes
    If GetAttribute(xmlNodeValidity, "StartDate") = strTemp Then _
        Exit For
Next

For lCtrl = 0 To TAX_Utilities.NodeValidity.childNodes.length - 1
    SetAttribute TAX_Utilities.NodeValidity.childNodes(lCtrl), "Active", _
        GetAttribute(xmlNodeValidity.childNodes(lCtrl), "Active")
Next lCtrl

Set xmlNodeMenu = Nothing
Set xmlNodeValidity = Nothing
Set xmlDom = Nothing

End Sub
