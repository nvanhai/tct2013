VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmTreeviewMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin FPUSpreadADO.fpSpread sstv 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _Version        =   458752
      _ExtentX        =   6800
      _ExtentY        =   5106
      _StockProps     =   64
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmTreeviewMenu.frx":0000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   4620
      X2              =   4620
      Y1              =   60
      Y2              =   4500
   End
End
Attribute VB_Name = "frmTreeviewMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : FIS - CMC Software Solution
' Project           : Du an ho tro ke khai thue
' Package           : Interface
' Form, Module
'   or Class name   : frmTreeviewMenu
' Descriptions      : Report sh
' Start date        : 11/10/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : TuanLM
' Integrate         :
' Project manager   : ThietKN
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
Private Const MENU_COLOR = 1 'RGB(53, 78, 171)
Private Const MENU_FORE_COLOR = &HFFFFFF
Private Const MENU_FORE_COLOR1 = &HFFFFFF


Dim pluspict1 As Picture
Dim pluspict2 As Picture
Dim pluspict3 As Picture
Dim minuspict1 As Picture
Dim minuspict2 As Picture
Dim minuspict3 As Picture
Dim subline As Picture
Dim subline1 As Picture
Dim fillerline As Picture
Dim endline As Picture
Dim prevbnum As Long, prevprow As Long
Dim prevsel(0, 1) As Long
Dim arrnodemenu(50, 4) As String    'Store the demo info
Dim isend As Boolean
Dim actRow As Long

Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long



'****************************************************
'Description:Form_Load procedure initialize the values of controls
'   Step 1: Load TreeviewMenu
'   Step 2: Load other information
'   Step 3: Load the interface default
'Author:TuanLM
'Modify by:
'Date:02/11/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub Form_Load()
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = 3000 'frmSystem.ScaleWidth / 4
    Me.Height = frmSystem.ScaleHeight
    
    'Load Menu
    LoadTreeViewMenu
            
    'Load other informations
    LoadOtherInfor
    
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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub LoadTreeViewMenu()
On Error GoTo ErrorHandle
    'Init the spread tree
    BeginfpTreeView
    
    'Set up the nodes
    LoadNodeTreeView
    
    'Must call this sub when finishing the tree
    EndfpTreeView 1
                
    ' Init
'    InitActiveForm
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "LoadTreeViewMenu", Err.Number, Err.Description
End Sub


'****************************************************
'Description:LoadOtherInfors procedure load others informations
'   Step 1: Load the message from Message.xml
'   Step 2: Load the header from Header.xml
'   Step 3: Set path value for help file
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub LoadOtherInfor()
On Error GoTo ErrorHandle

    'Load list of messages
    'LoadListMessage
    
    'Load file header
    'LoadHeaderFile

    'Set path for help file
    App.HelpFile = App.path & "\HTKK_CQT.chm"
    
    'Set background for menu
'    Me.Picture = LoadPicture("..\Pictures\menu_bg.gif")
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "LoadOtherInfor", Err.Number, Err.Description
End Sub



'****************************************************
'Description:LoadNodeTreeView procedure load others informations
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub LoadNodeTreeView()
    On Error GoTo ErrorHandle
    
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
    xmlDocument.Load App.path & "\Menu_CQT.xml"
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
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub GetMenuNode(pxmlNodeMenu As MSXML.IXMLDOMNode)
    On Error GoTo ErrorHandle
    
    Dim id As String
    Dim parent As String
    Dim caption As String
    Dim icon  As String
    
    id = GetAttribute(pxmlNodeMenu, "ID")
    caption = GetAttribute(pxmlNodeMenu, "Caption")
    parent = GetAttribute(pxmlNodeMenu, "ParentID")
    icon = GetAbsolutePath(GetAttribute(pxmlNodeMenu, "Icon"))
    
    If Len(parent) = 0 Then
        AddHeaderNode id, caption, "", icon, False
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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
            ElseIf NodeID = "103" Then
                .TypePictPicture = minuspict3
            Else
                .TypePictPicture = minuspict2
            End If
        Else
            If NodeID = "100" Then
                .TypePictPicture = pluspict1
            ElseIf NodeID = "103" Then
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
            .ForeColor = MENU_FORE_COLOR1 'vbBlack
            '.FontItalic = True
        .BlockMode = False
        
'        .BackColor = RGB(53, 78, 171)
        
        'Set the text
        .FontBold = True
        .Text = NodeText
        
        'Text tip text
        arrnodemenu(.Row, 0) = NodeID
        arrnodemenu(.Row, 1) = TipText

        
        'Set col widths
'        If col <> 1 Then .colwidth(.col) = 1.75
        
'        .col = .col + 1
'        .CellType = CellTypeStaticText

'        If col <> 1 Then .colwidth(.col) = 8
    End With
    If NodeID = "103" Then
        isend = True
    End If
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "AddHeaderNode", Err.Number, Err.Description
End Sub



'****************************************************
'Description:AddSubNode procedure add a childrent menu node
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignTop
        .TypePictPicture = subline
        .RowHeight(.Row) = 13
        '.BackColor = MENU_COLOR
        
        'Filler line
'        .Col = .Col + 1
'        .CellType = CellTypePicture
'        .TypePictCenter = True
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignTop
'        .TypePictStretch = True
'        .TypePictMaintainScale = True
'        .RowHeight(.Row) = 14
'        .TypePictPicture = LoadPicture(icon) 'fillerline
'        .BackColor = RGB(53, 78, 171)
        
        'Node text
        .Col = .Col + 1
        .Col2 = .MaxCols
        .Row = .Row
        .Row2 = .Row
        .BlockMode = True
        .CellType = CellTypeStaticText
        .BlockMode = False
        .Text = NodeText
        .ForeColor = MENU_FORE_COLOR 'vbBlack
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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
            
'                If sstv.TypePictPicture = minuspict Or sstv.TypePictPicture = pluspict Then
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
    For i = 1 To sstv.DataRowCnt
        ret = sstv.GetRowItemData(i)
        If ret <> 0 Then
            'Is a header row
            'Show or hide the child rows
            ShowHideRows i, ret
        End If
    Next i
    
    'Change col width of last column
    'sstv.ColWidth(sstv.MaxCols) = 17.025 '19.875
    
    sstv.RowHeight(sstv.MaxRows) = 5000
    'sstv.InsertRows
    
    SetBackColorForMenu
 
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "EndfpTreeView", Err.Number, Err.Description
End Sub



'****************************************************
'Description:BeginfpTreeView procedure set init paramters
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
        .GrayAreaBackColor = RGB(74, 121, 198) 'vbWhite
        .MaxCols = 10  'Set the maximum number of columns
        .MaxRows = 100  'Set the maximum number of rows
        .GridSolid = True
        .BackColorStyle = BackColorStyleOverGrid
        .ScrollBars = ScrollBarsNone
        .CursorStyle = CursorStyleArrow
        .NoBeep = True
        
        'Set column widths
        .ColWidth(1) = 2.375
        .ColWidth(2) = 2.125
        .ColWidth(3) = 1.75
        .ColWidth(4) = 15.725
        
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
    Me.Top = 0
    Me.Left = 0
    Me.Width = 3150 'frmSystem.ScaleWidth / 4
    Me.Height = frmSystem.ScaleHeight - 400
    
    sstv.Top = 0
    sstv.Left = 0
    sstv.Width = Me.ScaleWidth - 25
    sstv.Height = Me.ScaleHeight
    
    Me.BackColor = RGB(74, 121, 198)
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Resize", Err.Number, Err.Description

End Sub



'****************************************************
'Description:sstv_Click procedure invoke event click
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

'****************************************************
Private Sub GetMenuAction(Row As Long)
On Error GoTo ErrorHandle
    'Select the item or hide/show the rows
    Dim ret As Long
    Dim Col As Long
    
    'Get the row item data
    ret = sstv.GetRowItemData(Row)
    
    If ret = 0 Then
        'Not a header row
        'Select the item
        SelectItem sstv.MaxCols, Row
        
        'Process action
        ProcessMenuAction arrnodemenu(Row, 0)
    Else
        'Is a header row
        'Show or hide the child rows

        ShowHideRows Row, ret
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub


'****************************************************
'Description:ShowHideRows procedure hidden rows menu
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Description:sstv_KeyPress procedure invoke event KeyPress
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

'****************************************************
Private Sub sstv_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Using keyboard navigation

    If NewRow = -1 Then Exit Sub
    actRow = NewRow
    HighlightItem sstv.MaxCols, NewRow
    
End Sub

'****************************************************
'Description:sstv_MouseDown procedure load others informations
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

'****************************************************
Private Sub sstv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long, Row As Long

    'Rolling over an item
    'Highlight with dotted border style

    'Get the row and column currently over
    sstv.GetCellFromScreenCoord Col, Row, X, Y

    HighlightItem Col, Row

    actRow = Row
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "setupMenuData", Err.Number, Err.Description
End Sub

'****************************************************
'Description:sstv_MouseUp procedure load others informations
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
        sstv.ForeColor = MENU_FORE_COLOR  ' RGB(172, 172, 172)
        sstv.BlockMode = False
        
        'Save new row,col number
        prevsel(0, 0) = Row   'Row
        prevsel(0, 1) = Col   'Col
    
        'Set border and colors for the selected item
        SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_SOLID, vbRed
        
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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'            If sstv.TypePictPicture = pluspict Or sstv.TypePictPicture = minuspict Then
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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
'Author:TuanLM
'Modify by:
'Date:03/11/2005
'Input:
'Output:
'Return:

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
            SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_SOLID, vbGreen
        Else
            'Set demo border
            SetCellBorder Col, Row, SS_BORDER_TYPE_OUTLINE, SS_BORDER_STYLE_FINE_DOT, vbRed
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
'Author:TuanLM
'Modify by:
'Date:29/10/2005
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
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************
Private Sub InitActiveForm()
    On Error GoTo ErrorHandle
    
    Dim lID As String
    Dim lInterface As String
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim lnode As MSXML.IXMLDOMNode
    Dim i As Integer
    
    i = 0
    For Each xmlNode In xmlNodeListMenu
        lID = xmlNode.Attributes.getNamedItem("ID").nodeValue
'        Set lnode = xmlNode.selectSingleNode("Validity")
        
'        If Not lnode Is Nothing Then
            i = i + 1
            ReDim Preserve arrActiveForm(i) As activeForm
            arrActiveForm(i).id = lID
            arrActiveForm(i).showed = False
'        End If
    Next
    Set xmlNode = Nothing
    
    hasActiveForm = False
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "initActiveForm", Err.Number, Err.Description
End Sub



'****************************************************
'Description:LoadHeaderFile procedure load header form Header.xml
'
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

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
        
    strDataFileName = GetAbsolutePath(GetAttribute(xmlNodeSheet, "Folder")) & GetAttribute(xmlNodeSheet, "DataFile") & ".xml"
    
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
'Author:TuanLM
'Modify by:
'Date:03/10/2005
'Input:
'Output:
'Return:

'****************************************************
Private Sub ProcessMenuAction(pID As String)
    On Error GoTo ErrorHandle
    
    Dim lnode As MSXML.IXMLDOMNode
    Dim arrStrFileNames() As String
    Dim blnMenuHide As Boolean
    
    blnMenuHide = True
    Select Case pID
        Case "100_1" ' thoat
            CloseApplication
            Exit Sub
        Case "100_2" ' Tham so
            frmThamso.Show
            Exit Sub
        
       
        Case "100_3" ' Mat khau
            frmChangepass.Show
            Exit Sub
        Case "100_4" ' Tham so QHSCC
            frmThamsoQHSCC.Show
            Exit Sub
        Case "100_5" ' Tra cuu AC
            frmTraCuuAC.Show vbModeless
            Exit Sub
        Case "100_6" ' Tra cuu AC loi
            frmTraCuuACError.Show
            Exit Sub
            
        Case "102_1" ' thoat
            Exit Sub
        Case "102_2" ' thong ke
            Exit Sub
        Case "103_1" ' thong ke
            ShowAboutDgl
            Exit Sub
        Case "103_2" ' huong dan su dung
            ShowHelpDlg
            Exit Sub
        Case "101_1"
            If PortNotOnpened Then
                MSComm1.PortOpen = False
                frmInterfaces.SetReceiveByBarcode True
                frmInterfaces.Show
            Else
                blnMenuHide = False
            End If
        Case "101_3"
            If checkActivePIT = True Then
                ShowTruyennhanTkDgl
            Else
                DisplayMessage "0136", msOKOnly, miInformation
            End If
            Exit Sub
        Case "101_4"
            ShowNhanLaiTkDgl
            Exit Sub
        Case "100_5"
            frmTheodoiTK.Show
            Exit Sub
        Case "104_1"
            frmTraCuu.Show
        Case "101_2"
            frmInterfaces.SetReceiveByBarcode False
            arrStrFileNames = frmBrowser.GetFileNames()
            
            If UBound(arrStrFileNames) > 0 Then
                frmInterfaces.SetArrayElements arrStrFileNames
                frmInterfaces.Show
            Else
                blnMenuHide = False
            End If
        Case Else
            Exit Sub
    End Select
    
    If blnMenuHide Then
        Me.Hide
    End If
    
    Set lnode = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ProcessMenuAction", Err.Number, Err.Description
End Sub

'****************************************************
'Description:ShowHelpDlg procedure show help
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub ShowHelpDlg()
    On Error GoTo ErrorHandle
    
    HTMLHelp Me.hwnd, App.HelpFile, HH_DISPLAY_TOC, CLng(0)
        
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowHelpDlg", Err.Number, Err.Description
End Sub

'****************************************************
'Description:ShowTruyennhanTkDgl procedure show TruyenNhan
'Author:nvanhai
'Modify by:
'Date:14/10/2010
'Input:
'Output:
'Return:

'****************************************************
Private Sub ShowTruyennhanTkDgl()
 'Dim frmSearching As New frmTraCuu
 'frmSearching.Show
 frmTruyennhanTK.Show
 'Me.Show
End Sub

'****************************************************
'Description:ShowNhanLaiTkDgl procedure show NhanLaiTK
'Author:nvanhai
'Modify by:
'Date:14/10/2010
'Input:
'Output:
'Return:

'****************************************************
Private Sub ShowNhanLaiTkDgl()
 'Dim frmSearching As New frmTraCuu
 'frmSearching.Show
 frmNhanLaiTk.Show
 'Me.Show
End Sub

'****************************************************
'Description:ShowAboutDgl procedure show form about
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************

Private Sub ShowAboutDgl()
    On Error GoTo ErrorHandle
    
    'Exit Sub

    Dim aboutDlg As New frmAbout
    aboutDlg.Show
    
'    DisplayMessage "0001", msYesNoCancel, miWarning
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "ShowAboutDgl", Err.Number, Err.Description
End Sub



'****************************************************
'Description:getNode procedure return Node of ID from ListNode
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input: pID: id of node
'Output:
'Return: Node

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
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:pID: id of form
'Output:
'Return: index of form

'****************************************************

Public Function getFormIndex(pID As String) As Integer

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
'Author:TuanLM
'Modify by:
'Date:1/11/2005
'Input: pxmlNodeMenu: node menu
'Output:
'Return:

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
Public Sub TienichTracuu()
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
            sstv.BackColor = RGB(74, 121, 198)
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

'******************************
'Description: PortNotOnpened function check whether port in use by another program
'Author:ThanhDX
'Date: 22/12/2005
'Return: true if port not in use
'        false otherwise
'******************************
Private Function PortNotOnpened() As Boolean
    On Error GoTo ErrHandle
    MSComm1.PortOpen = True
    PortNotOnpened = True
    Exit Function
ErrHandle:
    DisplayMessage "0061", msOKOnly, miCriticalError
End Function
