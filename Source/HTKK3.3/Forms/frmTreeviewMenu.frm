VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTreeviewMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9870
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPUSpreadADO.fpSpread fpSpread1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   6135
      _Version        =   458752
      _ExtentX        =   10821
      _ExtentY        =   1085
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmTreeviewMenu.frx":0000
   End
   Begin FPUSpreadADO.fpSpread sstv 
      Height          =   5100
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6015
      _Version        =   458752
      _ExtentX        =   10610
      _ExtentY        =   8996
      _StockProps     =   64
      BorderStyle     =   0
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmTreeviewMenu.frx":026C
      VirtualScrollBuffer=   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmTreeviewMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmTreeviewMenu
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
Private Const HH_HELP_CONTEXT = &HF
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3
Private Const INIT_ROW = 1
Private Const INIT_HEIGHT = 5
Private Const MENU_BGCOLOR = 1 'RGB(53, 78, 171)
Private Const MENU_FORE_COLOR = &H140664


'Dim pluspict As Picture
Dim pluspict1 As Picture
Dim pluspict2 As Picture
Dim pluspict3 As Picture
'Dim minuspict As Picture
Dim minuspict1 As Picture
Dim minuspict2 As Picture
Dim minuspict3 As Picture
Dim subline As Picture
Dim subline1 As Picture
Dim fillerline As Picture
Dim endline As Picture
Dim prevbnum As Long, prevprow As Long
Dim prevsel(0, 1) As Long
Dim arrnodemenu(114, 4) As String    'Store the demo info
Dim isend As Boolean
Dim actRow As Long

'****************************************************
'Description:Form_Load procedure initialize the values of controls
'   Step 1: Load TreeviewMenu
'   Step 2: Load other information
'   Step 3: Load the interface default


'Date:02/11/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub Form_Load()
    Dim fso As New FileSystemObject
    
    If fso.FileExists("..\InterfaceTemplates\Template.xls") Then
'        If fpSpread1.IsExcelFile("..\InterfaceTemplates\Template.xls") Then
'            fpSpread1.EventEnabled(EventAllEvents) = False
'            fpSpread1.ImportExcelBook GetAbsolutePath("..\InterfaceTemplates\Template.xls"), vbNullString
'            fpSpread1.EventEnabled(EventAllEvents) = True
'        End If
        fpSpread1.LoadFromFile "..\InterfaceTemplates\Template.xls"
    End If
    
    Set fso = Nothing
    
    'Load Menu
    LoadTreeViewMenu
            
    'Load other informations
    LoadOtherInfor
    
    HighlightItem sstv.MaxCols, 2
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = 0
    
End Sub

'****************************************************
'Description:LoadTreeViewMenu procedure load Menu
'   Step 1: Set init paramter
'   Step 2: Load nodes menu
'   Step 3: Set finish paramter
'****************************************************

Private Sub LoadTreeViewMenu()
On Error GoTo ErrorHandle
    'Init the spread tree
    BeginfpTreeView
    
    'Set up the nodes
    LoadNodeTreeView
    
    'Must call this sub when finishing the tree
    EndfpTreeView 1
    
    hideAllRow 7   ' Focus vao group thi se active dong dau tien cua group
    ' Init
    InitActiveForm
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "LoadTreeViewMenu", Err.Number, Err.Description
End Sub
'****************************************************
'Description:LoadOtherInfors procedure load others informations
'   Step 1: Load the message from Message.xml
'   Step 2: Load the header from Header.xml
'   Step 3: Set path value for help file
'****************************************************

Private Sub LoadOtherInfor()
On Error GoTo ErrorHandle
    
    'Load file header
    LoadHeaderFile

    'Set path for help file
    App.HelpFile = App.path & "\HTKK.chm"
    
    'Set background for menu
    'Me.Picture = LoadPicture(GetAbsolutePath("..\Pictures\menu_bg.gif"))
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "LoadOtherInfor", Err.Number, Err.Description
End Sub
'****************************************************
'Description:LoadNodeTreeView procedure load others informations
'****************************************************

Private Sub LoadNodeTreeView()
    On Error GoTo ErrorHandle
    
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
    xmlDocument.Load App.path & "\Menu.xml"
    Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
    
    For Each xmlNode In xmlNodeListMenu
        GetMenuNode xmlNode
    Next
    
    Set xmlNode = Nothing
    Set xmlDocument = Nothing
    
    Exit Sub
 
ErrorHandle:
    SaveErrorLog Me.Name, "LoadNodeTreeView", Err.Number, Err.Description
End Sub
'****************************************************
'Description:GetMenuNode procedure create node menu for TreeviewMenu
'   from file Menu.xml
'   Step 1: Read information from the data node
'   Step 2: Create node menu if it has child or it has
'           file template in Datafiles directory
'****************************************************

Private Sub GetMenuNode(pxmlNodeMenu As MSXML.IXMLDOMNode)
    On Error GoTo ErrorHandle
    
    Dim id As String
    Dim parent, PopID As String
    Dim caption As String
    Dim icon  As String
    
    id = GetAttribute(pxmlNodeMenu, "ID")
    caption = GetAttribute(pxmlNodeMenu, "Caption")
    parent = GetAttribute(pxmlNodeMenu, "ParentID")
    PopID = GetAttribute(pxmlNodeMenu, "PopID")
    '****************************
    ' removed
    'icon = GetAttribute(pxmlNodeMenu, "Icon")
    '****************************
    '****************************
    
    icon = GetAbsolutePath(GetAttribute(pxmlNodeMenu, "Icon"))
    
    '****************************
    If Len(parent) = 0 And Len(PopID) = 0 Then
        AddHeaderNode id, caption, "", icon, False
    ElseIf Len(parent) <> 0 And Len(PopID) = 0 Then
        AddHeaderNode1 id, caption, "", icon, False
    Else
        If HasInterfaceTemplate(pxmlNodeMenu) Then
            AddSubNode id, caption, "", icon
        End If
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "GetMenuNode", Err.Number, Err.Description
End Sub
'****************************************************
'Description:AddHeaderNode procedure add a parent menu node
'****************************************************

Private Sub AddHeaderNode(NodeID As String, NodeText As String, TipText As String, icon As String, hidenode As Boolean)
On Error GoTo ErrorHandle
'Add a new header mode
  Dim Col As Integer
  Col = 1
    With sstv
    
        .Row = .Row + 1
        .Col = Col
                
        'Add the end picture
        .Row = .Row - 1
        .Col = Col + 1
        If .Row > INIT_ROW Then
            .TypePictPicture = endline
        End If
        .Row = .Row + 1
        .Col = Col
        'Plus/Minus Picture
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignRight
        If hidenode = True Then
            If NodeID = "100" Then
                .TypePictPicture = minuspict1
                
                
            ElseIf NodeID = "102" Then
                .TypePictPicture = minuspict3
            Else
                .TypePictPicture = minuspict2
            End If
        Else
            If NodeID = "100" Then
                .TypePictPicture = pluspict1
            ElseIf NodeID = "102" Then
                .TypePictPicture = pluspict3
            Else
                .TypePictPicture = pluspict2
            End If
        End If
'        If col <> 1 Then .ColWidth(.col) = 2.375
        
'        .BackColor = RGB(53, 78, 171)
        
        'Image List Picture
        .Col = .Col + 1
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignTop
        .RowHeight(.Row) = 12
        .TypePictStretch = True
        .TypePictMaintainScale = True
        .TypePictPicture = LoadPicture(icon)
 '       If col <> 1 Then .colwidth(.col) = 3.125
        
'        .BackColor = RGB(53, 78, 171)
        
        'Node text
        .Col = .Col + 1
        .Col2 = .MaxCols
        .Row = .Row
        .Row2 = .Row
        .TypeHAlign = TypeHAlignRight
        .BlockMode = True
            .CellType = CellTypeStaticText
'            .BackColor = &HC0FFFF
            .ForeColor = RGB(0, 0, 0) 'vbBlack
            '.FontItalic = True
        .BlockMode = False
        
'        .BackColor = RGB(53, 78, 171)
        
        'Set the text
        .Text = NodeText
        .FontBold = True
        
        
        'Text tip text
        arrnodemenu(.Row, 0) = NodeID
        arrnodemenu(.Row, 1) = TipText

        
        'Set col widths
'        If col <> 1 Then .colwidth(.col) = 1.75
        
'        .col = .col + 1
'        .CellType = CellTypeStaticText

'        If col <> 1 Then .colwidth(.col) = 8
    End With
    If NodeID = "102" Then
        isend = True
    End If
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "AddHeaderNode", Err.Number, Err.Description
End Sub
'****************************************************
'Description:AddSubNode procedure add a childrent menu node
'****************************************************
Private Sub AddHeaderNode1(NodeID As String, NodeText As String, TipText As String, icon As String, hidenode As Boolean)
On Error GoTo ErrorHandle
'Add a new header mode
  Dim Col As Integer
  Col = 1
    With sstv
        .Row = .Row + 1
        If Not isend Then
            .Col = Col
            .CellType = CellTypePicture
            .TypePictCenter = True
            .TypeHAlign = TypeHAlignRight
            .TypeVAlign = TypeVAlignTop
            .TypePictPicture = subline1
        End If
        Col = 2
        .Col = Col + 1
                
        'Add the end picture
        .Row = .Row - 1
        .Col = Col + 1
        If .Row > INIT_ROW Then
            .TypePictPicture = endline
        End If
        .Row = .Row + 1
        .Col = Col
        'Plus/Minus Picture
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignCenter
        If hidenode = True Then
            If NodeID = "100" Then
                .TypePictPicture = minuspict1
            ElseIf NodeID = "102" Then
                .TypePictPicture = minuspict3
            Else
                .TypePictPicture = minuspict2
            End If
        Else
            If NodeID = "100" Then
                .TypePictPicture = pluspict1
            ElseIf NodeID = "102" Then
                .TypePictPicture = pluspict3
            Else
                .TypePictPicture = pluspict2
            End If
        End If
'        If col <> 1 Then .ColWidth(.col) = 2.375
        
'        .BackColor = RGB(53, 78, 171)
        
        'Image List Picture
        .Col = .Col + 1
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignTop
        .RowHeight(.Row) = 12
        .TypePictStretch = True
        .TypePictMaintainScale = True
        .TypePictPicture = LoadPicture(icon)
 '       If col <> 1 Then .colwidth(.col) = 3.125
        
'        .BackColor = RGB(53, 78, 171)
        
        'Node text
        .Col = .Col + 1
        .Col2 = .MaxCols
        .Row = .Row
        .Row2 = .Row
        .TypeHAlign = TypeHAlignRight
        .BlockMode = True
            .CellType = CellTypeStaticText
'            .BackColor = &HC0FFFF
            .ForeColor = RGB(0, 0, 0) 'vbBlack
            '.FontItalic = True
        .BlockMode = False
        
'        .BackColor = RGB(53, 78, 171)
        
        'Set the text
        .Text = NodeText
        .FontBold = True
        
        
        'Text tip text
        arrnodemenu(.Row, 0) = NodeID
        arrnodemenu(.Row, 1) = TipText

        
        'Set col widths
'        If col <> 1 Then .colwidth(.col) = 1.75
        
'        .col = .col + 1
'        .CellType = CellTypeStaticText

'        If col <> 1 Then .colwidth(.col) = 8
    End With
    If NodeID = "102" Then
        isend = True
    End If
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "AddHeaderNode", Err.Number, Err.Description
End Sub
'****************************************************
'Description:AddSubNode procedure add a childrent menu node
'****************************************************

Private Sub AddSubNode(NodeID As String, NodeText As String, TipText As String, icon As String)
On Error GoTo ErrorHandle
'Add a sub node
Dim Col As Long
Col = 1
    With sstv
'        .BackColor = RGB(53, 78, 171)
        .Row = .Row + 1
        If Not isend Then
            .Col = Col
            .CellType = CellTypePicture
            .TypePictCenter = True
            .TypeHAlign = TypeHAlignRight
            .TypeVAlign = TypeVAlignTop
            .TypePictPicture = subline1
        End If
        
        .Col = Col + 1
        
        'SubLine
        .CellType = CellTypePicture
        .TypePictCenter = True
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignTop
        .TypePictPicture = subline
        .RowHeight(.Row) = 13
        
        'Node text
        .Col = .Col + 1
        .Col2 = .MaxCols
        .Row = .Row
        .Row2 = .Row
        .BlockMode = True
        .CellType = CellTypeStaticText
        .BlockMode = False
        .Text = NodeText
        .ForeColor = RGB(0, 0, 0) 'vbBlack
'        .BackColor = RGB(53, 78, 171)
        
        'Text tip text
        arrnodemenu(.Row, 0) = NodeID
        arrnodemenu(.Row, 1) = TipText
        
        
    End With
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "AddSubNode", Err.Number, Err.Description
End Sub



'****************************************************
'Description:EndfpTreeView procedure set final paramter
'****************************************************

Private Sub EndfpTreeView(Col As Long)
On Error GoTo ErrorHandle
    Dim i As Long, j As Long, z As Long, ret As Long
    Dim ctarray(10) As Integer
    Dim lastcol As Integer
    'Finish the tree
    
    sstv.MaxCols = sstv.DataColCnt + 2 'Set the maximum number of columns
    sstv.MaxRows = sstv.DataRowCnt + 1 'Set the maximum number of rows
    
    'Add the end picture
    sstv.Col = Col + 1
    sstv.TypePictPicture = endline
    
    'Loop through all rows and columns
    'If a header row, store the number of child rows that belong to it using the SetRowItemData function
    
    'Init the array
    For i = 0 To UBound(ctarray) - 1
        ctarray(i) = 0
    Next i
    
    For i = 1 To sstv.DataRowCnt
        For j = 1 To sstv.MaxCols
            sstv.Row = i
            sstv.Col = j
            If sstv.CellType = CellTypePicture And IsNumeric(sstv.TypePictPicture) Then
                If sstv.TypePictPicture = minuspict1 Or sstv.TypePictPicture = minuspict2 Or sstv.TypePictPicture = minuspict3 Or _
                sstv.TypePictPicture = pluspict1 Or sstv.TypePictPicture = pluspict2 Or sstv.TypePictPicture = pluspict3 Then
                    If ctarray(j) = 0 Then
                        'First time in. Save row number
                        ctarray(j) = i
                        lastcol = j
                    Else
                        'Set the data
                        sstv.SetRowItemData ctarray(j), i - ctarray(j)
                        ctarray(j) = i
                        lastcol = j
                        'Update any columns after this one
                        For z = j To UBound(ctarray)
                            If ctarray(z) <> 0 Then
                                sstv.SetRowItemData ctarray(z), i - ctarray(z)
                                ctarray(z) = i
                            End If
                        Next z
                        
                    End If
                End If
            End If
        Next j
    Next i
    
    'Save the last item
   sstv.SetRowItemData ctarray(lastcol), sstv.DataRowCnt - ctarray(lastcol) + 1
   
    'Show/hide rows
'    For i = 1 To sstv.DataRowCnt
'        ret = sstv.GetRowItemData(i)
'        If ret <> 0 Then
'            'Is a header row
'            'Show or hide the child rows
'            ShowHideRows i, ret
'        End If
'    Next i
    
    'Change col width of last column
    sstv.ColWidth(sstv.MaxCols) = 17.125 '19.875
    'MsgBox sstv.Height
    sstv.RowHeight(sstv.MaxRows) = 500
    'sstv.InsertRows
    
    For i = 1 To sstv.MaxRows
        sstv.Row = i
        sstv.ScrollBarShowMax = False
        sstv.ScrollBarTrack = ScrollBarTrackOff
    Next
    SetBackColorForMenu
 
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "EndfpTreeView", Err.Number, Err.Description
End Sub



'****************************************************
'Description:BeginfpTreeView procedure set init paramters
'****************************************************
Private Sub BeginfpTreeView()
On Error GoTo ErrorHandle
    'Init the spread tree view control
    Dim X As Boolean

    'Init the pictures
'    Set pluspict = LoadPicture(GetAbsolutePath("..\Pictures\out_pluscol.bmp"))
    Set pluspict1 = LoadPicture(GetAbsolutePath("..\Pictures\out_pluscol1.bmp"))
    Set pluspict2 = LoadPicture(GetAbsolutePath("..\Pictures\out_pluscol2.bmp"))
    Set pluspict3 = LoadPicture(GetAbsolutePath("..\Pictures\out_pluscol3.bmp"))
'    Set minuspict = LoadPicture(GetAbsolutePath("..\Pictures\out_minuscol.bmp"))
    Set minuspict1 = LoadPicture(GetAbsolutePath("..\Pictures\out_minuscol1.bmp"))
    Set minuspict2 = LoadPicture(GetAbsolutePath("..\Pictures\out_minuscol2.bmp"))
    Set minuspict3 = LoadPicture(GetAbsolutePath("..\Pictures\out_minuscol3.bmp"))
    Set subline = LoadPicture(GetAbsolutePath("..\Pictures\out_subline.bmp"))
    Set subline1 = LoadPicture(GetAbsolutePath("..\Pictures\out_subline1.bmp"))
    Set fillerline = LoadPicture(GetAbsolutePath("..\Pictures\out_filler.bmp"))
    Set endline = LoadPicture(GetAbsolutePath("..\Pictures\out_oneline.bmp"))
    
    With sstv
        .Appearance = 0 'Appearance3D  '3D appearance
        .EditModePermanent = True   'Do not show the highlight box around a selected cell
        .AllowCellOverflow = True   'Allow text to flow into adjacent cells
        .GridShowHoriz = False      'Turn off Horizontal gridlines
        .GridShowVert = False       'Turn off Vertical gridlines
        .ColHeadersShow = False   'Turn off column headers
        .RowHeadersShow = False  'Turn off row headers
        .ScrollBarExtMode = True    'Show scroll bars when needed
        .ArrowsExitEditMode = True
        .GrayAreaBackColor = RGB(244, 238, 202) 'vbWhite
        .MaxCols = 10  'Set the maximum number of columns
        .MaxRows = 114  'Set the maximum number of rows
        .GridSolid = True
        .BackColorStyle = BackColorStyleOverGrid
        .ScrollBars = ScrollBarsNone
        .VScrollSpecialType = VScrollSpecialTypeNoLineUpDown
        .CursorStyle = CursorStyleArrow
        .NoBeep = True
        
        'Set column widths
        .ColWidth(1) = 2.375
        .ColWidth(2) = 2.525
        .ColWidth(3) = 1.75
        .ColWidth(4) = 36
        
        'Initialize every cell as a Picture cell
        .Col = 1
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .BlockMode = True
            .CellType = CellTypePicture
        .BlockMode = False
        
        'Init TextTip appearance
        .TextTip = TextTipFixed
        X = .SetTextTipAppearance("Tahoma", "8", False, False, &HC0FFFF, &H800000)
   
    End With
    
    'Init the row var that was last "rolled over"
    prevbnum = 0    'for borders
    prevprow = 0    'for header pictures
    
    'Init the row, col array that was last clicked on (selected demo)
    prevsel(0, 0) = 0   'Row
    prevsel(0, 1) = 0   'Col
    
    'Init the row
    sstv.Row = INIT_ROW
    sstv.RowHeight(INIT_ROW) = INIT_HEIGHT
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "BeginfpTreeView", Err.Number, Err.Description
End Sub


Private Sub Form_Resize()
    On Error GoTo ErrorHandle
    'Me.Visible = False
    Me.Top = 0
    Me.Left = 0
    Me.Width = 6150 'frmSystem.ScaleWidth / 4
    Me.Height = 8200 'frmSystem.ScaleHeight

    sstv.Top = 0
    sstv.Left = 0
    sstv.Width = Me.ScaleWidth - 25
    sstv.Height = Me.ScaleHeight
'    fpSpread1.Top = 0
'    fpSpread1.Left = 0
    fpSpread1.Width = Me.ScaleWidth - 25
    fpSpread1.Height = Me.ScaleHeight
    Me.BackColor = RGB(244, 238, 202)
    'LoadNodeTreeView
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Resize", Err.Number, Err.Description
End Sub

'****************************************************
'Description:sstv_Click procedure invoke event click
'****************************************************
Private Sub sstv_Click(ByVal Col As Long, ByVal Row As Long)
    'Select the item or hide/show the rows

    GetMenuAction Row
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub
'****************************************************
'Description:GetMenuAction procedure process the actions of menu
'****************************************************
Private Sub GetMenuAction(Row As Long)
On Error GoTo ErrorHandle
    'Select the item or hide/show the rows
    Dim ret As Long
    Dim Col As Long
    
    'Get the row item data
    ret = sstv.GetRowItemData(Row)
    
    If ret = 0 Then
    'MsgBox "1"
        'Not a header row
        'Select the item
        SelectItem sstv.MaxCols, Row
        
        'Process action
        ProcessMenuAction arrnodemenu(Row, 0)
    Else
        'Is a header row
        'Show or hide the child rows

       ' ShowHideRows Row, ret
       hideAllRow Row
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub
'****************************************************
'Description:ShowHideRows procedure hidden rows menu
'****************************************************
Private Sub ShowHideRows(startrow As Long, rownum As Long)
On Error GoTo ErrorHandle
    Dim rowcnt As Long, Col As Long
    Dim i As Long, showtype As Integer
    'Show or hide the child rows

    Col = GetPlusMinusCell(startrow)
    If Col = -1 Then Exit Sub   'Not a header row
    
    'Turn off redraw
    sstv.ReDraw = False
    
    'Get the current picture
    sstv.Col = Col
    sstv.Row = startrow
    Select Case sstv.TypePictPicture
        'Hide rows
        Case pluspict1
            sstv.TypePictPicture = minuspict1
            showtype = 1
        Case pluspict2
            sstv.TypePictPicture = minuspict2
            showtype = 1
        Case pluspict3
            sstv.TypePictPicture = minuspict3
            showtype = 1
        
        'Show rows
        Case minuspict1
            sstv.TypePictPicture = pluspict1
            showtype = 0
        Case minuspict2
            sstv.TypePictPicture = pluspict2
            showtype = 0
        Case minuspict3
            sstv.TypePictPicture = pluspict3
            showtype = 0
    End Select
        
    
    'Turn off all borders
'    SetCellBorder 1, -1, SS_BORDER_TYPE_NONE, SS_BORDER_STYLE_DEFAULT, 0
    
    
    'Set vars for showing/hiding rows
    rowcnt = startrow + rownum - 1
    startrow = startrow + 1
    
    'Show or hide the rows
    For i = startrow To rowcnt
        sstv.Row = i
        If showtype = 0 Then
            'Hide Rows
            sstv.RowHidden = True
        Else
            'Show Rows
            sstv.RowHidden = False
        End If
    Next i
    
    'Turn on redraw
    sstv.ReDraw = True
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowHideRows", Err.Number, Err.Description
End Sub
'****************************************************
'Description:ShowHideRows procedure hidden rows menu
'****************************************************
Private Sub hideAllRow(Row As Long)
On Error GoTo ErrorHandle
    Dim i As Long
    If Row = 6 Then
        'Row = Row + 3
        Row = Row + 1   ' Bo KHBS nen row giam 2 dong
     End If
    'hide all row
     For i = 1 To sstv.MaxRows
        sstv.Row = i
        If sstv.GetRowItemData(i) = 0 Then
            sstv.RowHidden = True
        Else
            sstv.Col = GetPlusMinusCell(i)
            Select Case sstv.TypePictPicture
                'Hide rows
                Case minuspict1
                    sstv.TypePictPicture = pluspict1
                Case minuspict2
                    sstv.TypePictPicture = pluspict2
                Case minuspict3
                    sstv.TypePictPicture = pluspict3
             End Select
        End If
     Next i
    'Show or hide the rows
    Dim numRow As Long
    numRow = Row + sstv.GetRowItemData(Row)
    
    For i = Row To numRow
    sstv.Row = i
        If sstv.GetRowItemData(i) = 0 Then
            sstv.RowHidden = False
        Else
            sstv.Col = GetPlusMinusCell(Row)
            Select Case sstv.TypePictPicture
                'Hide rows
                Case pluspict1
                    sstv.TypePictPicture = minuspict1
                Case pluspict2
                    sstv.TypePictPicture = minuspict2
                Case pluspict3
                    sstv.TypePictPicture = minuspict3
             End Select
        End If
    Next i
    'Turn on redraw
    sstv.ReDraw = True
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowHideRows", Err.Number, Err.Description
End Sub
'****************************************************
'Description:sstv_KeyPress procedure invoke event KeyPress
'****************************************************
Private Sub sstv_KeyPress(KeyAscii As Integer)
'Using keyboard navigation
    If KeyAscii = 13 Or KeyAscii = 32 Then  '13 = Enter key, 32 = Space bar
        If actRow <> sstv.ActiveRow Then
            GetMenuAction actRow
        Else
            GetMenuAction sstv.ActiveRow ' actRow
        End If
    End If
End Sub

'****************************************************
'Description:sstv_LeaveCell procedure load others informations
'****************************************************
Private Sub sstv_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Using keyboard navigation
    If NewRow = -1 Then Exit Sub
    actRow = NewRow
    HighlightItem sstv.MaxCols, NewRow
    
End Sub
'****************************************************
'Description:sstv_MouseDown procedure load others informations
'****************************************************
Private Sub sstv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Col As Long, Row As Long
    Dim ret As Long

    'Clicked on an item
    'Change border to solid line
    
    If sstv.CellType = CellTypeStaticText Then
        'Get the row and column currently over
        sstv.GetCellFromScreenCoord Col, Row, X, Y
        ret = sstv.GetRowItemData(Row)
        If ret <> 0 Then Exit Sub   'Is a header row: exit sub
        
        'Update var for tracking last row was "rolled over"
        prevbnum = Row
    
        'Add cell border
        Col = GetColWithText(Col, Row)
        SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_FINE_DOT, &HC0C0C0
        
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub

'****************************************************
'Description:sstv_MouseMove procedure load others informations
'****************************************************
Private Sub sstv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long, Row As Long

    'Rolling over an item
    'Highlight with dotted border style
    
    'Get the row and column currently over
    sstv.GetCellFromScreenCoord Col, Row, X, Y
    actRow = Row
'    sstv.ActiveRow = Row
    HighlightItem Col, Row
        
    
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub

'****************************************************
'Description:sstv_MouseUp procedure load others informations
'****************************************************
Private Sub sstv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Col As Long, Row As Long

    'Clicked on an item
    'Change border to solid line
    
    'Get the row and column currently over
    sstv.GetCellFromScreenCoord Col, Row, X, Y
       
    'Select the item
    SelectItem Col, Row
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub

'****************************************************
'Description:SelectItem procedure load others informations
'****************************************************
Private Sub SelectItem(Col As Long, Row As Long)
On Error GoTo ErrorHandle
    'Select the item
    Dim ret As Long
    Dim demotext As Variant

    sstv.Col = Col
    sstv.Row = Row
    If sstv.CellType = CellTypeStaticText Then
        
        ret = sstv.GetRowItemData(Row)
        If ret <> 0 Then Exit Sub       'Not a header row. Exit
        
        'Get the column number that contains the header text
        Col = GetColWithText(Col, Row)
        
        'Clear the previously selected item
        sstv.Col = prevsel(0, 1)
        sstv.Col2 = sstv.MaxCols
        sstv.Row = prevsel(0, 0)
        sstv.Row2 = prevsel(0, 0)
        sstv.BlockMode = True
        'Set selection color
'        sstv.BackColor = vbWhite
        sstv.ForeColor = RGB(0, 0, 0) ' RGB(172, 172, 172)
        sstv.BlockMode = False
        
        'Save new row,col number
        prevsel(0, 0) = Row   'Row
        prevsel(0, 1) = Col   'Col
    
        'Set border and colors for the selected item
'        SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_SOLID, vbRed
        
        sstv.Col = Col
        sstv.Col2 = sstv.MaxCols
        sstv.Row = Row
        sstv.Row2 = Row
        sstv.BlockMode = True
            'Set selection color
'            sstv.BackColor = RGB(0, 67, 123) 'vbBlue
'            sstv.ForeColor = vbBlue  'vbWhite
        sstv.BlockMode = False
        
        
        ret = GetColWithText(sstv.MaxCols, Row)
        sstv.GetText ret, Row, demotext
        
        
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SelectItem", Err.Number, Err.Description
End Sub

'****************************************************
'Description:GetColWithText procedure load others informations
'****************************************************
Private Function GetColWithText(Col As Long, Row As Long) As Long
On Error GoTo ErrorHandle
    'Return the column number that contains the text
    'AllowCellOverflow = true. The cell may display the data
    '  but not necessarily contain the data.  Get the cell that contains the text
    Dim i As Integer
    
    GetColWithText = 0
    
    sstv.Row = Row
    'Loop through all columns
    For i = Col To 1 Step -1
        sstv.Col = i
        If sstv.Text <> "" Then
            'Contains the text.
            GetColWithText = i
            Exit For
        End If
    Next i
    
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "GetColWithText", Err.Number, Err.Description
End Function

'****************************************************
'Description:GetPlusMinusCell procedure load others informations
'****************************************************
Private Function GetPlusMinusCell(Row As Long) As Long
On Error GoTo ErrorHandle
Dim i As Long
    'Returns the column number that contains the plus,minus picture

    GetPlusMinusCell = -1
    
    sstv.Row = Row
    'Loop through all columns
    For i = 1 To sstv.MaxCols
        sstv.Col = i
        If sstv.CellType = CellTypePicture And IsNumeric(sstv.TypePictPicture) Then
            If sstv.TypePictPicture = minuspict1 Or sstv.TypePictPicture = minuspict2 Or sstv.TypePictPicture = minuspict3 Or _
             sstv.TypePictPicture = pluspict1 Or sstv.TypePictPicture = pluspict2 Or sstv.TypePictPicture = pluspict3 Then
                'Found the header cell picture
                GetPlusMinusCell = i
                Exit For
            End If
        End If
    Next i
    
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "GetPlusMinusCell", Err.Number, Err.Description
End Function

'****************************************************
'Description:SetCellBorder procedure load others informations
'****************************************************
Private Sub SetCellBorder(Col As Long, Row As Long, BorderType As Integer, BorderStyle As Integer, BorderColor As Long)
On Error GoTo ErrorHandle
    'Set the cell borders
    
    sstv.SetCellBorder Col, Row, sstv.MaxCols - 1, Row, BorderType, BorderColor, BorderStyle
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetCellBorder", Err.Number, Err.Description
End Sub

'****************************************************
'Description:HighlightItem procedure load others informations
'****************************************************
Private Sub HighlightItem(Col As Long, Row As Long)
On Error GoTo ErrorHandle
    Dim ret As Long, invalidpic As Boolean
    'Rolling over an item
    'Highlight with dotted border style
    
    sstv.ReDraw = False
    
    sstv.Row = Row
    sstv.Col = Col
    
    Select Case sstv.CellType
        Case CellTypePicture
            
            'Change picture and cursor
            'See if a header picture
            invalidpic = False
            If Not IsNumeric(sstv.TypePictPicture) Then
                invalidpic = True
            ElseIf sstv.TypePictPicture <> minuspict1 And sstv.TypePictPicture <> minuspict2 And sstv.TypePictPicture <> minuspict3 Or _
                sstv.TypePictPicture <> pluspict1 And sstv.TypePictPicture <> pluspict2 And sstv.TypePictPicture <> pluspict3 Then
'            ElseIf sstv.TypePictPicture <> pluspict And sstv.TypePictPicture <> minuspict Then
                invalidpic = True
            End If
            
            'Not a header picture. Exit
            If invalidpic = True Then
'                If sstv.CursorStyle <> CursorStyleDefault Then sstv.CursorStyle = CursorStyleDefault
                Exit Sub
            End If
           
            If prevprow = Row Then Exit Sub
            
            'Header picture.
            'Change cursor
'            If sstv.CursorStyle <> CursorStyleArrow Then sstv.CursorStyle = CursorStyleArrow
            
     Case CellTypeStaticText
        'Draw a rollover border
        If prevbnum = Row Then Exit Sub   'See if same row
         
        'Remove previous border
        SetCellBorder 1, prevbnum, SS_BORDER_TYPE_NONE, SS_BORDER_STYLE_DEFAULT, 0
        
        'Get the beginning column that contains text
        Col = GetColWithText(Col, Row)
        
        'See if a header row
        ret = sstv.GetRowItemData(Row)
        If ret <> 0 Then
            'Set Header Border
            SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_SOLID, RGB(168, 9, 13)
        Else
            'Set demo border
            SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_FINE_DOT, RGB(168, 9, 13)
        End If
        
        'Update row holder
        prevbnum = Row
        
     End Select
     
     sstv.ReDraw = True
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "HighlightItem", Err.Number, Err.Description
End Sub

'Private Sub sstv_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'Private Sub sstv_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
'Display the text tip
    
'    If demoarray(Row, 0) = "" Then Exit Sub
'
'    If CLng(demoarray(Row, 0)) <= Col Then
'        TipText = demoarray(Row, 1)   'Set the text tip text
'        MultiLine = 1
'        TipWidth = 1500
'        ShowTip = True
'    End If
    
'End Sub


'****************************************************
'Description:HasInterfaceTemplate procedure check if file interfase exists in Interfaces directory
'Input: pxmlNodeMenu: node menu
'Output:
'Return: True - if file template exits, otherwise it is False
'****************************************************
Private Function HasInterfaceTemplate(pxmlNodeMenu As MSXML.IXMLDOMNode) As Boolean
    On Error GoTo ErrorHandle
    
    Dim xmlListValidity As MSXML.IXMLDOMNodeList
    Dim xmlNodeValidity As MSXML.IXMLDOMNode
    Dim strTemplateFile  As String
    Dim bIn As Boolean
    Dim fso As New FileSystemObject
    
    bIn = False
    Set xmlListValidity = pxmlNodeMenu.selectNodes("Validity")
    
    ' function system
    If xmlListValidity.length = 0 Then
        HasInterfaceTemplate = True
        Exit Function
    End If
    
    ' function business
    For Each xmlNodeValidity In xmlListValidity
        strTemplateFile = GetAbsolutePath(GetAttribute(xmlNodeValidity, "InterfaceTemplate"))
        If fso.FileExists(strTemplateFile) Then
            bIn = True
            Exit For
        End If
    Next
    
    HasInterfaceTemplate = bIn
    
    Set xmlNodeValidity = Nothing
    Set xmlListValidity = Nothing
    Set fso = Nothing
    
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "HasInterfaceTemplate", Err.Number, Err.Description
End Function

'****************************************************
'Description:InitActiveForm procedure initialize the status for
'   the functions of menu. The status of all function
'   is false (not show)
'****************************************************
Private Sub InitActiveForm()
    On Error GoTo ErrorHandle
    
    Dim lID As String
    Dim lInterface As String
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim lnode As MSXML.IXMLDOMNode
    Dim i As Integer
    
    i = 0
    ReDim arrActiveForm(0)
    For Each xmlNode In xmlNodeListMenu
        lID = xmlNode.Attributes.getNamedItem("ID").nodeValue
        Set lnode = xmlNode.selectSingleNode("Validity")
        
        If Not lnode Is Nothing Then
            i = i + 1
            ReDim Preserve arrActiveForm(i) As activeForm
            arrActiveForm(i).id = lID
            arrActiveForm(i).showed = False
        End If
    Next
    Set xmlNode = Nothing
    
    hasActiveForm = False
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "initActiveForm", Err.Number, Err.Description
End Sub

'****************************************************
'Description:LoadHeaderFile procedure load header form Header.xml
'****************************************************

Public Sub LoadHeaderFile()
    On Error GoTo ErrorHandle
    
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeValidity As MSXML.IXMLDOMNode
    Dim xmlNodeSheet As MSXML.IXMLDOMNode
    Dim strDataFileName  As String
    
    For Each xmlNode In xmlNodeListMenu
        If xmlNode.Attributes.getNamedItem("ID").nodeValue = "100_1" Then
            Exit For
        End If
    Next
    
    Set xmlNodeValidity = xmlNode.selectNodes("Validity").Item(0)
    Set xmlNodeSheet = xmlNodeValidity.selectNodes("Sheet").Item(0)
        
    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(xmlNodeSheet, "DataFile") & ".xml"
    
    xmlHeaderData.Load strDataFileName
        
    Set xmlNode = Nothing
    Set xmlNodeValidity = Nothing
    Set xmlNodeSheet = Nothing
    
    Exit Sub
 
ErrorHandle:
    SaveErrorLog Me.Name, "LoadHeaderFile", Err.Number, Err.Description
End Sub

'****************************************************
'Description:ProcessMenuAction procedure process the action of menu
'****************************************************
Public Sub ProcessMenuAction(pID As String)
    On Error GoTo ErrorHandle
    
    Dim lnode As MSXML.IXMLDOMNode
    
    Set lnode = getNode(pID)
    If lnode Is Nothing Then Exit Sub
    
    Me.Hide
    Select Case pID
        Case "100_2" ' sao luu
            'BackupRestorData True
'***********************
' added
'Date: 08/01/2006
            BackupData
            Me.Show
'***********************
        Case "100_3" ' phuc hoi
'***********************
' added
'Date: 08/01/2006
            RestoreData
            Me.Show
'***********************
            'BackupRestorData False
        Case "100_4" ' thoat
            CloseApplication
            Exit Sub
        Case "100_5" ' profile
            ChangeProfile
        Case "100_6" ' Import data
            ImportTaxReport
'***********************
' added
'Date: 13/04/06
        ' Show searching form
        Case "100_8" ' Import data
            ShowSearchingDgl
'***********************
' added
'Date: 13/04/06
        ' Show searching form
         Case "1"
            ShowKhaibosung
            
'***********************
        
        Case "102_1" ' gioi thieu
            Me.Show
            ShowAboutDgl
            Exit Sub
        Case "102_2" ' huong dan su dung
            Me.Show
            ShowHelpDlg
            Exit Sub
         Case "102_3" ' bang ke ban ra
            Me.Show
            ShowBangkebanra
            Exit Sub
          Case "102_4" ' mua vao
            Me.Show
            ShowBangkemuavao
            Exit Sub
            'bang ke TTDB
            'dhdang sua ngay 15/01/2011
'          Case "102_9" ' 06B/TNCN
'            Me.Show
'            ShowBangke06BTNCN
'          Exit Sub
          Case "102_5" ' TNCN04
            Me.Show
            'ShowBangkeTNCN
            ShowBangkebanraTTDB
            Exit Sub
          Case "102_6" ' TNDN05
            Me.Show
            'ShowMau05TNDN
            ShowBangkemuavaoTTDB
            Exit Sub
          Case "102_7" ' 01-7/GTGT
            Me.Show
            '01-7/GTGT
            ShowBangke01_7GTGT
            Exit Sub
          Case "102_8" ' 25/MGT-TNCN
            Me.Show
            '25/MGT-TNCN
            ShowBangke25_MGT_TNCN
            Exit Sub
          Case "102_9" ' BKGH
            Me.Show
            'BKGH
            ShowBangkeGH_TNDN
            Exit Sub
          Case "102_10" ' 01/NTNN
            Me.Show
            '01NTNN
            ShowBangke01_NTNN
            Exit Sub
          Case "102_11" ' 02/MT-GTGT
            Me.Show
            '02/MT-GTGT
            ShowPLMienThueTT140 "PL_02MT_GTGT_TT140_MaSoThue.xls"
            Exit Sub
          Case "102_12" ' 01/MGT-TNDN
            Me.Show
            '01/MGT-TNDN
            ShowPLMienThueTT140 "PL_01MGT_TNDN_TT140_MaSoThue.xls"
            Exit Sub
          Case "102_13" ' 26/MT-TNCN
            Me.Show
            '01/TNDN TT199
            ShowPLMienThueTT140 "PL_01_TNDN_TT199_MaSoThue.xls"
            Exit Sub
          Case "102_14" ' 01/GH-TNDN
            Me.Show
            '01/GH-TNDN
            ShowPLMienThueTT140 "PL_01GH_TNDN_TT16_MaSoThue.xls"
            Exit Sub
          Case "102_16" ' 02/GH-GTGT
            Me.Show
            '02/GH-GTGT
            ShowPLMienThueTT140 "PL_02GH_GTGT_TT16_MaSoThue.xls"
            Exit Sub
          Case "102_17" ' bang ke ban ra tK 04/GTGT
            Me.Show
            ShowPLMienThueTT140 "Bangkebanra_04GTGT.xls"
            Exit Sub
          Case "102_18" ' bang ke ban ra tK 01-3/GTGT
            Me.Show
            ShowPLMienThueTT140 "Bangkebanra_01_3.xls"
            Exit Sub
          Case "102_19" ' bang ke 02-1/TNDN
            Me.Show
            ShowPLMienThueTT140 "Bangke_02_1TNDN.xls"
            Exit Sub
            
        Case Else
            ShowFormFunction lnode
    End Select
        
    'frmSystem.Picture = LoadPicture(GetAbsolutePath("..\Pictures\bg1.gif"))
    
    Set lnode = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ProcessMenuAction", Err.Number, Err.Description
End Sub


'****************************************************
'Description:ShowFormFunction procedure show form function
'   Step 1: Show frmPeriod to user can chose the priod
'   Step 2: Show frmInterfases
'****************************************************

Private Sub ShowFormFunction(pNode As MSXML.IXMLDOMNode)
    On Error GoTo ErrorHandle
    
    Dim frmTK As frmInterfaces
    Dim frPeriod As frmPeriod
    Dim i As Integer
    Dim sYear As String
    
    If hasActiveForm Then
        Exit Sub
    End If
    
    i = getFormIndex(pNode.Attributes.getNamedItem("ID").nodeValue)
    If arrActiveForm(i).showed Then
        Exit Sub
    Else
        TAX_Utilities_v2.NodeMenu = pNode
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "FinanceYear") = "1" Then
            strNgayTaiChinh = GetNgayBatDauNamTaiChinh
            If Not KiemTraNgayTaiChinh(strNgayTaiChinh) Then
                frmTreeviewMenu.Show
                Exit Sub
            End If
        End If
        
        If (pNode.Attributes.getNamedItem("Year").nodeValue = "0") Then
            Set frmTK = New frmInterfaces
            frmTK.Show
            
        Else
            Set frPeriod = New frmPeriod
            frPeriod.Show
            
            ' Neu la in mau bia ho so quyet toan thi vao luon ma ko can click Ok
            If pNode.Attributes.getNamedItem("ID").nodeValue = "52" Or pNode.Attributes.getNamedItem("ID").nodeValue = "66" Or pNode.Attributes.getNamedItem("ID").nodeValue = "67" Or pNode.Attributes.getNamedItem("ID").nodeValue = "10" Or pNode.Attributes.getNamedItem("ID").nodeValue = "09" Then
                frPeriod.Hide
                frPeriod.cmdOK_Click
            End If
            
            
        End If
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "showFormFunction", Err.Number, Err.Description
    
End Sub
'****************************************************
'Description:BackupRestorData procedure backup or restore data
'****************************************************

Public Sub BackupRestorData(pBackup As Boolean)
    On Error GoTo ErrorHandle
    
    If hasActiveForm Then
        Exit Sub
    End If
    Dim frmBackup As New frmBackup_Restore
    frmBackup.bIsBackup = pBackup
    frmBackup.Show
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "BackupRestorData", Err.Number, Err.Description
End Sub
'****************************************************
'Description:ChangeProfile procedure change profile
'****************************************************

Public Sub ChangeProfile()
    On Error GoTo ErrorHandle
    
    If hasActiveForm Then
        Exit Sub
    End If
    Dim fLogin As New frmLogin
    fLogin.Show
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "BackupRestorData", Err.Number, Err.Description
End Sub
'****************************************************
'Description:ShowHelpDlg procedure show help
'****************************************************

Private Sub ShowHelpDlg()
    On Error GoTo ErrorHandle

    HTMLHelp Me.hwnd, App.HelpFile, HH_DISPLAY_TOC, CLng(0)

    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowHelpDlg", Err.Number, Err.Description
End Sub

Private Sub ShowBangkemuavaoTTDB()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "bangkemuavao_01TTDB.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub

Private Sub ShowBangke01_7GTGT()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "PL_01_7GTGT_TT154_MaSoThue.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub

' Phu luc mien thue theo TT140
Private Sub ShowPLMienThueTT140(FileName As String)
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & FileName, "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowPLMienThueTT140", Err.Number, Err.Description
End Sub


Private Sub ShowBangke25_MGT_TNCN()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "PL_25MGT_TNCN_TT154_MaSoThue.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub


Private Sub ShowBangkeGH_TNDN()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "BKGiaHan_TT170_MaSoThue.doc", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub


Private Sub ShowBangke01_NTNN()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Tokhai_01NTNN.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub

Private Sub ShowBangke06BTNCN()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bangke_06BTNCN.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub
Private Sub ShowBangkemuavao()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "bangkemuavao.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub
Private Sub ShowBangkebanraTTDB()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "bangkebanra_01TTDB.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub
Private Sub ShowBangkebanra()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "bangkebanra.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub

Private Sub ShowBangkeTNCN()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "Bangke_04TNCN.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub
Private Sub ShowMau05TNDN()
    On Error GoTo ErrorHandle
    Call ShellExecute(hwnd, "Open", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel\" & "hoahongdaily_05_TNDN.xls", "", Mid$(App.path, 1, InStrRev(App.path, "\")) & "InterfaceTemplates\excel", 3)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowBangke", Err.Number, Err.Description
End Sub

'****************************************************
'Description:ShowAboutDgl procedure show form about
'****************************************************

Private Sub ShowAboutDgl()
    On Error GoTo ErrorHandle
    
    'Exit Sub

'    Dim aboutDlg As New frmAbout
'    aboutDlg.Show
    Call ShellExecute(hwnd, "Open", GetAbsolutePath("\gioithieu.htm"), "", App.path, 3)
    'Shell GetAbsolutePath("\IEXPLORE.exe") & " " & GetAbsolutePath("\gioithieu.htm"), vbMaximizedFocus
'    DisplayMessage "0001", msYesNoCancel, miWarning
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowAboutDgl", Err.Number, Err.Description
End Sub

'****************************************************
'Description:getNode procedure return Node of ID from ListNode
'****************************************************

Public Function getNode(pID As String) As MSXML.IXMLDOMNode
    On Error GoTo ErrorHandle
    
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim lID As String
    
    For Each xmlNode In xmlNodeListMenu
        lID = xmlNode.Attributes.getNamedItem("ID").nodeValue
        If lID = pID Then
            Exit For
        End If
    Next
    
    Set getNode = xmlNode
    Set xmlNode = Nothing
    
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "getNode", Err.Number, Err.Description
End Function
'****************************************************
'Description:getFormIndex procedure return index of form has Id is pID
'****************************************************

Public Function getFormIndex(pID As String) As Integer
On Error GoTo ErrorHandle
    On Error GoTo ErrorHandle
    
    Dim i As Long
    
    For i = 1 To UBound(arrActiveForm)
        If arrActiveForm(i).id = pID Then
            getFormIndex = i
            Exit For
        End If
    Next
    
    Exit Function
ErrorHandle:
    SaveErrorLog Me.Name, "getFormIndex", Err.Number, Err.Description
End Function
'****************************************************
'Description:CloseApplication procedure exit application
'****************************************************
Public Sub CloseApplication()
    On Error GoTo ErrorHandle
    
    frmSystem.clickexit = True
    Unload Me
    Unload frmSystem
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "CloseApplication", Err.Number, Err.Description
End Sub
Private Sub SetBackColorForMenu()
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To sstv.MaxRows ' sstv.DataRowCnt
        sstv.Row = i
        For j = 1 To 5
            sstv.Col = j
            sstv.BackColor = RGB(244, 238, 202)
            'RGB(244, 238, 202)
            If (i = 1) Or (i = sstv.MaxRows) Or (j = 5) Then sstv.Lock = True
        Next j
    Next i
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "SetBackColorForMenu", Err.Number, Err.Description
End Sub

Private Sub sstv_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
sstv.Top = 0
sstv.TopRow = 1
End Sub

Private Sub ImportTaxReport()
    Dim strFileName As String
    Dim strContentOfFile As String
    
    On Error GoTo DialogError
    With CommonDialog1
        .CancelError = True
        .Filter = "Text file (*.txt)|*.txt"
        .FilterIndex = 1
        .DialogTitle = "Select a text file to import data"
        .ShowOpen
        strFileName = .FileName
    End With
        
    On Error GoTo ErrHandle
    If CheckPrefixContentOfFile(strFileName, strContentOfFile) Then
        If RestoreDataFile(strContentOfFile) Then
            strHiddenFormName = "ImportTaxReport"
            frmInterfaces.Show
        Else
            TAX_Utilities_v2.NodeValidity = Nothing
            Me.Show
        End If
    Else
        TAX_Utilities_v2.NodeValidity = Nothing
        Me.Show
    End If
    Exit Sub
DialogError:
    Me.Show
    Exit Sub
ErrHandle:
    TAX_Utilities_v2.NodeValidity = Nothing
    Me.Show
    SaveErrorLog Me.Name, "cmdExport_Click", Err.Number, Err.Description
End Sub

'****************************
'Description: CheckPrefixContentOfFile check whether
'             file is belong to.
'Input:
'       strBarcodeData: Data string.
'OutPut:
'Return: True if restore data file successfully
'        False if the otherwise.
'****************************
Private Function CheckPrefixContentOfFile(ByVal strFileName As String, ByRef strFileData As String) As Boolean
    Dim fso As New FileSystemObject
    Dim tstFile As TextStream, strDataFileName As String
    Dim strPrefix As String, strValidDate As String
    Dim lCtrl As Long
    Dim udtDateUtils As DateUtils
    Dim dNgayDauKy, dNgayCuoiKy As Date
    Dim dNgayDau, dNgayCuoi
    
    On Error GoTo ErrHandle
    
    'Initial parameters
    TAX_Utilities_v2.month = ""
    TAX_Utilities_v2.ThreeMonths = ""
    TAX_Utilities_v2.Year = ""
    TAX_Utilities_v2.FirstDay = ""
    TAX_Utilities_v2.LastDay = ""
    
    strFileData = ""
    Set tstFile = fso.OpenTextFile(strFileName, ForReading, False, TristateTrue)
    While Not tstFile.AtEndOfStream
        strFileData = strFileData & tstFile.ReadLine
    Wend
    tstFile.Close
    
    'Get Prefix infor
    strPrefix = Left$(strFileData, 49)
    
    'Check tax id
    If strTaxIdString <> Trim(Mid$(strPrefix, 3, 13)) Then
        DisplayMessage "0050", msOKOnly, miInformation
        Exit Function
    End If
    
    'Set node menu
    TAX_Utilities_v2.NodeMenu = getNode(Mid$(strPrefix, 1, 2))
    
    'Lay ngay bat dau nam tai chinh
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "FinanceYear") = "1" Then
        strNgayTaiChinh = GetNgayBatDauNamTaiChinh
        If Not KiemTraNgayTaiChinh(strNgayTaiChinh) Then
            Exit Function
        Else
            iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
            iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
        End If
    Else
        strNgayTaiChinh = "01/01"
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
    End If
    
    'Check validity of period
    If Not CheckPeriod(Mid$(strPrefix, 16, 2), Mid$(strPrefix, 18, 4)) Then
        Exit Function
    End If
    
    'Set Month or three month attr
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
        TAX_Utilities_v2.month = Mid$(strPrefix, 16, 2)
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
        TAX_Utilities_v2.ThreeMonths = CInt(Mid$(strPrefix, 16, 2))
    End If
    
    'Set year attr
    TAX_Utilities_v2.Year = Mid$(strPrefix, 18, 4)
    
    '********************************
'    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "1" Then
'        strNgayTaiChinh = GetNgayBatDauNamTaiChinh
'        If Not KiemTraNgayTaiChinh(strNgayTaiChinh) Then
'            Exit Function
'        Else
'            iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
'            iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
'        End If
'    End If
    
    '********************************
    
    'Set validity menu
    TAX_Utilities_v2.NodeValidity = GetValidityNode
    
    'Set first day and last day
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "0" Then
        TAX_Utilities_v2.FirstDay = Mid$(strPrefix, 30, 10)
        TAX_Utilities_v2.LastDay = Mid$(strPrefix, 40, 10)
    End If
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "0" Then
        Set udtDateUtils = New DateUtils
        dNgayDau = udtDateUtils.ToDate(TAX_Utilities_v2.FirstDay, "DD/MM/YYYY")
        dNgayCuoi = udtDateUtils.ToDate(TAX_Utilities_v2.LastDay, "DD/MM/YYYY")
        
        'Check valid of date
        
        If IsNull(dNgayDau) Or _
            IsNull(dNgayCuoi) Then
                DisplayMessage "0076", msOKOnly, miCriticalError
                Exit Function
        End If
    End If
'**************************************

    'Check validity of start date.
    strValidDate = GetAttribute(TAX_Utilities_v2.NodeValidity, "StartDate")
    
On Error GoTo ErrTypeMismatch
    If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))), _
            DateSerial(CInt(Mid$(strPrefix, 26, 4)), CInt(Mid$(strPrefix, 24, 2)), _
            CInt(Mid$(strPrefix, 22, 2)))) = 0 Then
        DisplayMessage "0054", msOKOnly, miInformation
        TAX_Utilities_v2.NodeValidity = Nothing
        Exit Function
    End If
    
On Error GoTo ErrHandle
'**************************************
    'Kiem tra hop le cua ngay bat dau va ngay ket thuc
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "0" Then
        dNgayDauKy = GetNgayDauQuy(4, CInt(TAX_Utilities_v2.Year) - 1, iNgayTaiChinh, iThangTaiChinh)
        dNgayCuoiKy = GetNgayCuoiQuy(1, CInt(TAX_Utilities_v2.Year) + 1, iNgayTaiChinh, iThangTaiChinh)
        
        If dNgayDau < dNgayDauKy Then
            DisplayMessage "0065", msOKOnly, miInformation
            Exit Function
        End If
        If dNgayCuoi > dNgayCuoiKy Then
            DisplayMessage "0066", msOKOnly, miInformation
            Exit Function
        End If
        If DateDiff("M", dNgayDau, dNgayCuoi) + 1 > 15 Then
            DisplayMessage "0068", msOKOnly, miInformation
            Exit Function
        End If
        If dNgayCuoi < dNgayDau Then
            DisplayMessage "0069", msOKOnly, miInformation
            Exit Function
        End If
    End If
'**************************************
    
    'Get main content
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "0" Then
        strFileData = Mid$(strFileData, 50)
    Else
        strFileData = Mid$(strFileData, 30)
    End If
    
    'Get file name of current tax report
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & IIf(Left$(TAX_Utilities_v2.ThreeMonths, 1) = "0", TAX_Utilities_v2.ThreeMonths, "0" & TAX_Utilities_v2.ThreeMonths) & TAX_Utilities_v2.Year & ".xml"
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" Then
        'Data file contain Day from and to.
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
        & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
    '********************************
    '  added
    ' Date: 04/04/06
    Else
        'Data file not contain Day from and to.
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
        & TAX_Utilities_v2.Year & ".xml"
    '********************************
    End If
    
    'Ky ke khai da ton tai
    If fso.FileExists(strDataFileName) Then
        If DisplayMessage("0074", msYesNo, miQuestion, , mrNo) = mrYes Then
            For lCtrl = 1 To TAX_Utilities_v2.NodeValidity.childNodes.length
                If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lCtrl - 1), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lCtrl - 1), "DataFile") & "_" & IIf(Left$(TAX_Utilities_v2.ThreeMonths, 1) = "0", TAX_Utilities_v2.ThreeMonths, "0" & TAX_Utilities_v2.ThreeMonths) & TAX_Utilities_v2.Year & ".xml"
                ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" Then
                    'Data file contain Day from and to.
                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lCtrl - 1), "DataFile") & "_" _
                    & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                '********************************
                '  added
                ' Date: 04/04/06
                Else
                    'Data file not contain Day from and to.
                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lCtrl - 1), "DataFile") & "_" _
                    & TAX_Utilities_v2.Year & ".xml"
                '********************************
                End If
                If fso.FileExists(strDataFileName) Then
                    fso.DeleteFile strDataFileName, True
                End If
            Next
        Else
            Exit Function
        End If
    End If
    Set fso = Nothing
    CheckPrefixContentOfFile = True
    Exit Function
ErrTypeMismatch:
    DisplayMessage "0072", msOKOnly, miCriticalError
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "CheckPrefixContentOfFile", Err.Number, Err.Description
    DisplayMessage "0072", msOKOnly, miCriticalError
End Function

'****************************
'Description: RestoreDataFile function restore
'             data files from data string.
'   Step 1: Cut data string into sheet datas
'   Step 2: Load content of sheet datas to DOM, load template to DOM
'   Step 3: Generate xml string and save it to xml file
'Input:
'       strBarcodeData: Data string.
'OutPut:
'Return: True if restore data file successfully
'        False if the otherwise.
'****************************
Private Function RestoreDataFile(ByVal strBarcodeData As String) As Boolean ', rsTaxInfor As ADODB.Recordset)
    Dim blnValidData As Boolean
    Dim strDataRestore As String, strFileName As String
    Dim lIndex As Long, lCtrl As Long, arrStrData() As String
    Dim xmlData As New MSXML.DOMDocument, xmlTemplate As New MSXML.DOMDocument
    Dim fso As New FileSystemObject, tstFile As TextStream
    
On Error GoTo ErrHandle
    arrStrData = GetSheetDatas(strBarcodeData)
    
'Kiem tra su ton tai cua to khai
    If Trim(arrStrData(1)) = vbNullString Then
        DisplayMessage "0072", msOKOnly, miCriticalError
        RestoreDataFile = False
        Exit Function
    End If
'****************************
    
    If UBound(arrStrData) < TAX_Utilities_v2.NodeValidity.childNodes.length Then
        DisplayMessage "0072", msOKOnly, miCriticalError
        RestoreDataFile = False
        Exit Function
    End If
    
    For lIndex = 1 To UBound(arrStrData())
        xmlTemplate.Load GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lIndex - 1), _
            "TemplateFolder")) & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lIndex - 1), "DataFile") & ".xml"
        
        If TAX_Utilities_v2.month <> "" Then
            strFileName = GetAbsolutePath(TAX_Utilities_v2.DataFolder) _
                & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" _
                & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
        ElseIf TAX_Utilities_v2.ThreeMonths <> "" Then
            strFileName = GetAbsolutePath(TAX_Utilities_v2.DataFolder) _
                & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" _
                & format(TAX_Utilities_v2.ThreeMonths, "0#") & TAX_Utilities_v2.Year & ".xml"
        ElseIf TAX_Utilities_v2.FirstDay <> "" And TAX_Utilities_v2.LastDay <> "" Then
                strFileName = GetAbsolutePath(TAX_Utilities_v2.DataFolder) _
                    & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" _
                    & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" _
                    & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
        Else
                strFileName = GetAbsolutePath(TAX_Utilities_v2.DataFolder) _
                    & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" _
                    & TAX_Utilities_v2.Year & ".xml"
        End If
        If arrStrData(lIndex) <> vbNullString Then
            If Not xmlData.loadXML(arrStrData(lIndex)) Then
                DisplayMessage "0072", msOKOnly, miCriticalError
                RestoreDataFile = False
                Exit Function
            End If
            
            'Get data string and structure
            'strDataRestore = GetSections(xmlData.firstChild, xmlTemplate.getElementsByTagName("Sections")(0), blnValidData)
            GetSections xmlTemplate.getElementsByTagName("Sections")(0), xmlData.firstChild, blnValidData
            If Not blnValidData Then
                RestoreDataFile = False
                Exit Function
            Else
                xmlTemplate.save strFileName
            End If
        Else
            'xmlTemplate.save strFileName
            'strDataRestore = xmlTemplate.xml
        End If

    Next lIndex
    
    Set xmlData = Nothing
    Set xmlTemplate = Nothing
    Set fso = Nothing
    
    RestoreDataFile = True
    
    Exit Function
ErrHandle:
    'DisplayMessage "0072", msOKOnly, miCriticalError
    SaveErrorLog Me.Name, "RestoreDataFile", Err.Number, Err.Description
End Function

'****************************
'Description: GetSheetDatas function divide data string into sheet datas.
'Author:
'Date:23/11/2005
'Input:strBarcodeData: Data string.
'Output:
'Return: array of data sheets.
'****************************
Private Function GetSheetDatas(ByVal strBarcodeData As String) As String()
    Dim arrStrData() As String ', strSheetId As String , strTemp As String
    Dim intIndex As Integer, intLoc1 As Integer, intLoc2 As Integer
    Dim xmlNode As MSXML.IXMLDOMNode
    
On Error GoTo ErrHandle
    For Each xmlNode In TAX_Utilities_v2.NodeValidity.childNodes
        SetAttribute xmlNode, "Active", "0"
    Next
    
    ReDim arrStrData(0)
   
    For Each xmlNode In TAX_Utilities_v2.NodeValidity.childNodes
            
        intLoc1 = InStr(1, strBarcodeData, "<S" & GetAttribute(xmlNode, "ID") & ">")
        
        If intLoc1 = 0 Then
            intIndex = intIndex + 1
            ReDim Preserve arrStrData(intIndex)
        Else
            intLoc2 = InStr(1, strBarcodeData, "</S" & GetAttribute(xmlNode, "ID") & ">")
            If intLoc2 > intLoc1 Then
                SetAttribute xmlNode, "Active", "1"
                intIndex = intIndex + 1
                ReDim Preserve arrStrData(intIndex)
                arrStrData(intIndex) = Mid$(strBarcodeData, intLoc1, intLoc2 + 5)
                strBarcodeData = Replace(strBarcodeData, arrStrData(intIndex), "")
            End If
        End If
    Next
    
    If strBarcodeData = "" Then
        If UBound(arrStrData) < TAX_Utilities_v2.NodeValidity.childNodes.length Then
            ReDim Preserve arrStrData(TAX_Utilities_v2.NodeValidity.childNodes.length)
        End If
    Else
        ReDim arrStrData(0)
    End If
    
    GetSheetDatas = arrStrData()
    Set xmlNode = Nothing

    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetSheetDatas", Err.Number, Err.Description
End Function

Private Sub GetCells(xmlSectionTemplate As MSXML.IXMLDOMNode, arrStrValue() As String)
    On Error GoTo ErrHandler
    Dim lCtrl As Long, lCtrl2 As Long
    
    'Fill data from array of data to Cell node
    While lCtrl <= UBound(arrStrValue) And Not xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2) Is Nothing
        If GetAttribute(xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2), "Receive") <> "0" Then
            SetAttribute xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2), "Value", _
                arrStrValue(lCtrl)
        Else
            lCtrl = lCtrl - 1
        End If
        lCtrl = lCtrl + 1
        lCtrl2 = lCtrl2 + 1
    Wend
    
    Exit Sub
ErrHandler:
    SaveErrorLog Me.Name, "GetCells", Err.Number, Err.Description
End Sub

'*******************************************
'Description: GetSection procedure convert data from data string
'               to Dom data.
'   xmlSectionTemplate: Section template node
'   xmlSectionData : Section data node
'*******************************************
Private Sub GetSection(xmlSectionTemplate As MSXML.IXMLDOMNode, xmlSectionData As MSXML.IXMLDOMNode, blnValidData As Boolean)

On Error GoTo ErrHandler

    Dim lCtrl As Long, lElementsNo As Long
    Dim lngMaxRow As Long, lngRows As Long
    Dim arrStrValue() As String
    
    lElementsNo = GetElementsNo(xmlSectionTemplate.childNodes(0))
    'Get array of data units
    arrStrValue = Split(xmlSectionData.Text, "~")
    If GetAttribute(xmlSectionTemplate, "Dynamic") = "0" Then
        'Static data
        If UBound(arrStrValue) + 1 > lElementsNo Then
            blnValidData = False
            DisplayMessage "0070", msOKOnly, miCriticalError
            Exit Sub
        End If
    Else
        ' Dynamic data
        '************************************
        'Check amount of dynamic row
        lngMaxRow = CLng(GetAttribute(xmlSectionTemplate, "MaxRows"))
        
        If lngMaxRow > 0 And ((UBound(arrStrValue) + 1) / lElementsNo > lngMaxRow) Then
            DisplayMessage "0102", msOKOnly, miWarning, , mrOK
            DoEvents
            lngRows = lngMaxRow
        Else
            lngRows = IIf((UBound(arrStrValue) + 1) Mod lElementsNo = 0, _
                         (UBound(arrStrValue) + 1) / lElementsNo, (UBound(arrStrValue) + 1) \ lElementsNo + 1)
        End If
        
        '************************************
        For lCtrl = 2 To lngRows
            'Insert nodes
            InsertNode xmlSectionTemplate
        Next lCtrl
    End If
    
    GetCells xmlSectionTemplate, arrStrValue
    
    blnValidData = True
    
    Exit Sub
ErrHandler:
    blnValidData = False
    SaveErrorLog Me.Name, "GetSection", Err.Number, Err.Description
End Sub

'*******************************************
'Description: GetSections procedure convert data from data string
'               to Dom data.
'   xmlSectionsTemplate: Sections template node
'   xmlSectionsData : Sections data node
'*******************************************
Private Sub GetSections(xmlSectionsTemplate As MSXML.IXMLDOMNode, xmlSectionsData As MSXML.IXMLDOMNode, blnValidData As Boolean)

On Error GoTo ErrHandler

    Dim xmlSectionNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    Dim intDifIndex As Integer              'Su khac nhau ve so thu tu
    
    If xmlSectionsData.childNodes.length > xmlSectionsTemplate.childNodes.length Then
        blnValidData = False
        DisplayMessage "0072", msOKOnly, miCriticalError
        Exit Sub
    End If

    For lCtrl = 1 To xmlSectionsTemplate.childNodes.length
        If GetAttribute(xmlSectionsTemplate.childNodes(lCtrl - 1), "Receive") <> "0" Then
            GetSection xmlSectionsTemplate.childNodes(lCtrl - 1), xmlSectionsData.childNodes(lCtrl - 1 - intDifIndex), blnValidData
            If Not blnValidData Then
                blnValidData = False
                Exit Sub
            End If
        Else
            intDifIndex = intDifIndex + 1
        End If
    Next
    blnValidData = True
    
    Exit Sub
ErrHandler:
    blnValidData = False
    SaveErrorLog Me.Name, "GetSections", Err.Number, Err.Description
End Sub

'*******************************************
'Description: BackupData procedure open a dialog and allows user
'             choose a zip file to backup the DataFiles folder.
'*******************************************
Private Sub BackupData()
    Dim strFileName As String
    Dim fso As New FileSystemObject
    Dim tst As TextStream, fld As Folder
    Dim objProcess As New clsProcessRunning
        
    On Error GoTo DialogError
Dialog:
    With CommonDialog1
        .CancelError = True
        .Filter = "zip file (*.zip)|*.zip"
        .FilterIndex = 1
        .DialogTitle = "Select a zip file to backup data"
        .ShowSave
        If Right$(.FileName, 4) <> ".zip" Then
            strFileName = .FileName & ".zip"
        Else
            strFileName = .FileName
        End If
    End With
    
    On Error GoTo ErrAccess
    
    If fso.FileExists(strFileName) Then
        If DisplayMessage("0052", msYesNo, miQuestion, , mrNo) = mrYes Then
            fso.DeleteFile strFileName, True
        Else
            GoTo Dialog
        End If
    End If
    
    'Set mouse pointer
    Screen.MousePointer = vbHourglass
    
    'Check protection of drive or folder
    fso.CreateTextFile Mid$(strFileName, 1, InStrRev(strFileName, "\")) & "Test.txt"
    
    'Check size of drive to backup data
    If fso.GetDrive(Mid$(strFileName, 1, InStr(1, strFileName, ":"))).FreeSpace < _
       fso.GetFolder(GetAbsolutePath("..\DataFiles")).Size Then
        Screen.MousePointer = vbDefault
        DisplayMessage "0037", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    'Create info file
    Set tst = fso.CreateTextFile(GetAbsolutePath("..\DataFiles\Info.txt"), True)
    For Each fld In fso.GetFolder(GetAbsolutePath("..\DataFiles")).SubFolders
        tst.WriteLine fld.Name
    Next
    tst.Close
    
    'Zip file
    On Error GoTo ErrHanlde
    Shell """" & GetAbsolutePath() & "\pkzip45""" & " -add -silent -dir=relative -temp=""" & Left$(GetAbsolutePath(TAX_Utilities_v2.DataFolder), Len(GetAbsolutePath(TAX_Utilities_v2.DataFolder)) - 1) & """ " & """" & strFileName & """ " & """" & GetAbsolutePath("..\DataFiles\*") & """"
            
    Do
        DoEvents
    Loop Until Not objProcess.ProcessRunning("pkzip45.exe")
    
    fso.DeleteFile Mid$(strFileName, 1, InStrRev(strFileName, "\")) & "Test.txt", True
'    'Wait for creating file
'    Dim lCtrl1 As Long, lCtrl2 As Long
'    Do While (Not fso.FileExists(GetAbsolutePath(strFileName))) And (lCtrl1 < 10000)
'        lCtrl1 = lCtrl1 + 1
'    Loop
'
'    'Delay
'    Do While (lCtrl2 < 10000000)
'        lCtrl2 = lCtrl2 + 1
'    Loop
    
'    If Not fso.FileExists(GetAbsolutePath(strFileName)) And (lCtrl1 >= 100000) Then
'        DisplayMessage "", msOKOnly, miCriticalError
'        Exit Sub
'    End If
    
    'Success
    Screen.MousePointer = vbDefault
    If fso.FileExists(GetAbsolutePath(strFileName)) Then
        DisplayMessage "0007", msOKOnly, miInformation
    Else
        DisplayMessage "0059", msOKOnly, miInformation
    End If
    
    'Delete Info file
    fso.DeleteFile GetAbsolutePath("..\DataFiles\Info.txt"), True
    
    Set objProcess = Nothing
    Set fso = Nothing
Exit Sub
DialogError:
    Screen.MousePointer = vbDefault
    Exit Sub
ErrAccess:
    'Access denied
    Screen.MousePointer = vbDefault
    If Err.Number = 70 Then
        DisplayMessage "0021", msOKOnly, miCriticalError
    ElseIf Err.Number = 5 Then 'Permission denied
        DisplayMessage "0025", msOKOnly, miCriticalError
    End If
    Exit Sub
ErrHanlde:
    'fso.DeleteFile GetAbsolutePath("..\DataFiles\Info.txt"), True
    Screen.MousePointer = vbDefault
    If Err.Number = 70 Then
        DisplayMessage "0021", msOKOnly, miCriticalError
    ElseIf Err.Number = 5 Then
        DisplayMessage "0025", msOKOnly, miCriticalError
    ElseIf Err.Number = 53 Then
        DisplayMessage "0058", msOKOnly, miCriticalError
    End If
    SaveErrorLog Me.Name, "BackupData", Err.Number, Err.Description
End Sub

'*******************************************
'Description: RestoreData procedure open a dialog and allows user
'             choose a zip file to restore the DataFiles folder.

'*******************************************
Private Sub RestoreData()
    Dim strFileName As String
    Dim fso As New FileSystemObject
    Dim fld As Folder, tst As TextStream
    Dim fle As file, strTemp As String
    Dim lCtrl1 As Long, lCtrl2 As Long
    Dim objProcess As New clsProcessRunning
    
    'Open dilalog
    On Error GoTo DialogError
    With CommonDialog1
        .CancelError = True
        .Filter = "zip file (*.zip)|*.zip"
        .FilterIndex = 1
        .DialogTitle = "Select a zip file to restore data"
        .ShowOpen
        If Right$(.FileName, 4) <> ".zip" Then
            strFileName = .FileName & ".zip"
        Else
            strFileName = .FileName
        End If
    End With
    
    'Selected file not exist
    If Not fso.FileExists(strFileName) Then
        DisplayMessage "0053", msOKOnly, miInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHanlde
'        'Copy zip file to ZIP temporary file
'        fso.CopyFile strFileName, Replace(strFileName, ".tct", ".zip"), True
'        strFileName = Left$(strFileName, Len(strFileName) - 4) & ".zip"
        
        'Delete info file if it exist
        If fso.FileExists(GetAbsolutePath("..\DataFiles\Info.txt")) Then
            fso.DeleteFile GetAbsolutePath("..\DataFiles\Info.txt"), True
        End If
        
        'Wait for deleting of info file
        Do While (fso.FileExists(GetAbsolutePath("..\DataFiles\Info.txt"))) And (lCtrl1 < 100000)
            lCtrl1 = lCtrl1 + 1
        Loop
        
        'Delay
        Do While (lCtrl2 < 100000)
            lCtrl2 = lCtrl2 + 1
            DoEvents
        Loop
        
        'If can not delete info file
        If fso.FileExists(GetAbsolutePath("..\DataFiles\Info.txt")) Then
            Screen.MousePointer = vbDefault
            DisplayMessage "0056", msOKOnly, miCriticalError
            Exit Sub
        End If
        
        'Extract to temporary folder
        Shell """" & GetAbsolutePath() & "\pkzip45""" & " -Extract -silent -Over=All -include=Info.txt " & """" & strFileName & """ " & """" & GetAbsolutePath("..\DataFiles") & """"
        
        'Wait for process finish
        Do
            DoEvents
        Loop Until Not objProcess.ProcessRunning("pkzip45.exe")
        
        'Wait for info file
        lCtrl1 = 0
        Do While (Not fso.FileExists(GetAbsolutePath("..\DataFiles\Info.txt"))) And (lCtrl1 < 10000)
            lCtrl1 = lCtrl1 + 1
        Loop
        
        'Delay
        lCtrl2 = 0
        Do While (lCtrl2 < 10000)
            lCtrl2 = lCtrl2 + 1
            DoEvents
        Loop
        
        'If info file not exit
        If Not fso.FileExists(GetAbsolutePath("..\DataFiles\Info.txt")) Then
            Screen.MousePointer = vbDefault
            DisplayMessage "0055", msOKOnly, miCriticalError
            Exit Sub
        End If
        
        'Read content of info file
        Set tst = fso.OpenTextFile(GetAbsolutePath("..\DataFiles\Info.txt"))
        
        'Delete exist folder
        While Not tst.AtEndOfStream
            strTemp = Trim(tst.ReadLine)
            If strTemp <> "" Then
                If fso.FolderExists(GetAbsolutePath("..\DataFiles") & "\" & strTemp) Then
                    fso.DeleteFolder GetAbsolutePath("..\DataFiles") & "\" & strTemp, True
                End If
            End If
        Wend
        tst.Close
        
        'Extract to Datafiles
        Shell """" & GetAbsolutePath() & "\pkzip45""" & " -Extract -silent -Over=All -dir=Relative " & """" & strFileName & """ " & """" & GetAbsolutePath("..\DataFiles") & """"
                
        'Wait for process finish
        Do
            DoEvents
        Loop Until Not objProcess.ProcessRunning("pkzip45.exe")
        
        'Success
        Screen.MousePointer = vbDefault
        DisplayMessage "0008", msOKOnly, miInformation

        Set fso = Nothing
        Set objProcess = Nothing
    Exit Sub
DialogError:
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHanlde:
    Screen.MousePointer = vbDefault
    SaveErrorLog Me.Name, "RestoreData", Err.Number, Err.Description
    On Error GoTo ErrorExit
    Screen.MousePointer = vbDefault
    If fso.FileExists(GetAbsolutePath("..\DataFiles\Info.txt")) Then
        fso.DeleteFile GetAbsolutePath("..\DataFiles\Info.txt"), True
    End If
    
    If Right$(strFileName, 4) = ".zip" Then _
            fso.DeleteFile strFileName, True
    Exit Sub
ErrorExit:
    SaveErrorLog Me.Name, "RestoreData", Err.Number, Err.Description
End Sub


Private Sub InsertNode(xmlSectionTemplate As MSXML.IXMLDOMNode)
    Dim xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode, xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim lRows As Long, lRow2s As Long
    Dim lRowLBound As Long, lRowUbound As Long
    Dim lRow As Long, lCol As Long
    
    Set xmlCellsNode = xmlSectionTemplate.lastChild
    lRows = GetDynRowCount(fpSpread1, xmlCellsNode, lRow2s, lRowLBound, lRowUbound)
    
    'Increase row value on each cell in Dom data
    IncreaseRowInDOM fpSpread1, xmlSectionTemplate.parentNode.parentNode, lRowUbound + 1, lRows, lRow2s
    
    Set xmlNodeNewCells = xmlCellsNode.CloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
        
        ' Set new ID2 for node (CellID2)
        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
    Next
    
    ' Insert new node to DOM object
    xmlCellsNode.parentNode.appendChild xmlNodeNewCells
End Sub

Private Function GetElementsNo(xmlCellsNode As MSXML.IXMLDOMNode) As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCntElementsNo As Long
    
    For Each xmlCellNode In xmlCellsNode.childNodes
        If GetAttribute(xmlCellNode, "Receive") <> "0" Then
            lCntElementsNo = lCntElementsNo + 1
        End If
    Next
    GetElementsNo = lCntElementsNo
End Function

Private Sub ShowSearchingDgl()
 'Dim frmSearching As New frmTraCuu
 'frmSearching.Show
 frmTraCuu.Show
 'Me.Show
End Sub
Private Sub ShowKhaibosung()
    frmKhaibosung.Show
End Sub



