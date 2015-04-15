VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmPreview 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11025
   WindowState     =   2  'Maximized
   Begin FPUSpreadADO.fpSpread fpSpread1 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   11745
      _Version        =   458752
      _ExtentX        =   20717
      _ExtentY        =   556
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
      SpreadDesigner  =   "frmPreview.frx":0000
   End
   Begin FPUSpreadADO.fpSpreadPreview fpspReport 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      _Version        =   458752
      _ExtentX        =   20558
      _ExtentY        =   14208
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Xem tr­íc b¶n in"
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
      TabIndex        =   2
      Top             =   30
      Width           =   3135
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************
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
'*************************************************************************************
Option Explicit
Private intPageCount As Integer
Private intCurrentPage As Integer
Private intLocalCurrPage As Integer

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
On Error GoTo ErrHandle
    If intCurrentPage < intPageCount Then
    'If fpspReport.PageCurrent < frmReportData.fpsReport.PrintPageCount Or frmReportData.fpsReport.Sheet < frmReportData.fpsReport.SheetCount - 1 Then
        Printer.Orientation = frmReportData.fpsReport.PrintOrientation
        If fpspReport.PageCurrent < frmReportData.fpsReport.PrintPageCount Then
            
            fpspReport_PageChange fpspReport.PageCurrent + 1
        Else ' Move next Sheet
            Do
                frmReportData.fpsReport.sheet = frmReportData.fpsReport.sheet + 1
                intLocalCurrPage = 1
            Loop Until GetAttribute(TAX_Utilities_v2.NodeValidity. _
                childNodes(frmReportData.fpsReport.sheet - 1), "Active") <> "0" _
                    Or frmReportData.fpsReport.sheet = frmReportData.fpsReport.SheetCount - 1
            If GetAttribute(TAX_Utilities_v2.NodeValidity. _
                childNodes(frmReportData.fpsReport.sheet - 1), "Active") _
                = "0" Then Exit Sub
            'Update global current page int this form
            intCurrentPage = intCurrentPage + fpspReport.PagesPerScreen
            Form_Activate
        End If
    End If
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdNext_Click", Err.Number, Err.Description
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo ErrHandle
    If 1 < intCurrentPage Then
    'If fpspReport.PageCurrent > 1 Or frmReportData.fpsReport.Sheet > 1 Then
        
        If fpspReport.PageCurrent > 1 Then
            
            fpspReport_PageChange fpspReport.PageCurrent - 1
        
        Else 'Move previous sheet
            Do
                frmReportData.fpsReport.sheet = frmReportData.fpsReport.sheet - 1
                Printer.Orientation = frmReportData.fpsReport.PrintOrientation
                intLocalCurrPage = frmReportData.fpsReport.PrintPageCount
            Loop Until GetAttribute(TAX_Utilities_v2.NodeValidity. _
                childNodes(frmReportData.fpsReport.sheet - 1), "Active") <> "0" Or frmReportData.fpsReport.sheet = 1
            
            If GetAttribute(TAX_Utilities_v2.NodeValidity. _
                childNodes(frmReportData.fpsReport.sheet - 1), "Active") _
                = "0" Then Exit Sub
                
            'Update global current page int this form
            intCurrentPage = intCurrentPage - fpspReport.PagesPerScreen
            Form_Activate
            Printer.Orientation = frmReportData.fpsReport.PrintOrientation
            fpspReport.PageCurrent = frmReportData.fpsReport.PrintPageCount
        End If
    End If
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdPrevious_Click", Err.Number, Err.Description
End Sub

Private Sub cmdPrint_Click()
Dim arrStrPages As Variant
Dim intNumberOfCopies As Integer
Dim udtPrinter As New clsDefaultPrinter
Dim lPrevPaperSize As Long

On Error GoTo ErrHandle

    'Set printer as default to print
    Printer.TrackDefault = True
    lPrevPaperSize = Printer.PaperSize
    If Not udtPrinter.SetPrinterAsDefault(strPrinterName) Then
        'Display message if it has error
        DisplayMessage "0026", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    'Set Printer to default printer of OS
    Printer.TrackDefault = True
    Printer.PaperSize = vbPRPSA4
'     If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(1), "Caption") = "04-1/TNCN" Then
'        Printer.PaperSize = vbPRPSA3
'     End If
        ' BC26
'    If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Caption") = "BC26-AC" Then
'        Printer.PaperSize = vbPRPSA3
'        Printer.Orientation = vbPRORLandscape
'    End If
    ' end

    
    'Check Ready of printer
    If Not frmReports.IsPrinterReady Then
        DisplayMessage "0057", msOKOnly, miCriticalError
        Set udtPrinter = Nothing
        Printer.PaperSize = lPrevPaperSize
        Exit Sub
    End If
    
    'CreateExcelBook
    'frmReportData.SetPrintedPages (arrStrPages)
    For intNumberOfCopies = 1 To CInt(frmReports.txtNumberOfCopies.Text)
            frmReportData.PrintTax
    Next intNumberOfCopies
    Unload frmReportData
    Unload Me
    Unload frmReports
    
    Set udtPrinter = Nothing
    Printer.PaperSize = lPrevPaperSize
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdPrint_Click", Err.Number, Err.Description
End Sub

Private Sub cmdZoom_Click()
fpspReport.ZoomState = 3
fpspReport.SetFocus
End Sub

Private Sub Form_Activate()
On Error GoTo ErrHandle
    'Attach preview control to Spread
    'frmReportData.fpsReport.hDCPrinter = Printer.hDC
    Printer.Orientation = frmReportData.fpsReport.PrintOrientation
    Me.fpspReport.hWndSpread = frmReportData.fpsReport.hwnd
    
    'Update page count listing
    UpdatePageCount
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Active", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim intCtrl As Integer, blnNextPageExist As Boolean
    
    blnNextPageExist = False
    
    SetControlCaption Me
    SetupToolbar
    
    'Set Page Count
    intPageCount = 0
    'frmReportData.fpsReport.sheet = 1
    'frmReportData.fpsReport.hDCPrinter = Printer.hDC
    For intCtrl = 1 To frmReportData.fpsReport.SheetCount - 1
        ' If sheet is active
        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(intCtrl - 1), "Active") <> "0" Then
            ' Has the next page
            If intCtrl > 1 Then _
                blnNextPageExist = True
            frmReportData.fpsReport.sheet = intCtrl
            Printer.Orientation = frmReportData.fpsReport.PrintOrientation
            intPageCount = intPageCount + frmReportData.fpsReport.PrintPageCount
        End If
    Next intCtrl
    
    'Set current page
    intCurrentPage = 1
    intLocalCurrPage = 1
    
    'Set first sheet to preview
    frmReportData.fpsReport.sheet = 1
    
    'Disable Previous button
    DisableButton 4, "LEFT"
        
    'Get the zoom display
'    GetZoom zoomindex
        
    If frmReportData.fpsReport.PrintPageCount = 1 And Not blnNextPageExist Then
        'Disable Next button if only one page
        DisableButton 2, "RIGHT"
    End If

Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
End Sub
Private Sub SetupToolbar()
Dim i As Integer

On Error GoTo ErrHandle
    
    'Specify whether Edit Mode is to remain on when switching between cells
    fpSpread1.EditModePermanent = True

    fpSpread1.Col = -1
    fpSpread1.Row = -1
    fpSpread1.Lock = True
    
    'Set the number of rows in the spreadsheet
    fpSpread1.MaxRows = 1
 
    'Set the height of a selected row
    fpSpread1.RowHeight(1) = 15
   
    'Set the number of columns in the spreadsheet
    fpSpread1.MaxCols = 16
 
    'Set the column widths
    For i = 1 To fpSpread1.MaxCols Step 2
        fpSpread1.ColWidth(i) = 0.3
    Next i
   
    'Resize wide column
    fpSpread1.ColWidth(14) = 15
    
    'Show or hide the column headers
    fpSpread1.DisplayColHeaders = False
    fpSpread1.DisplayRowHeaders = False
    
    'Turn off scroll bars
    fpSpread1.ScrollBars = ScrollBarsNone
    
    'Turn off border
    fpSpread1.BorderStyle = BorderStyleNone
      
    'Select row(s)
    fpSpread1.Row = 1
    fpSpread1.Col = -1

    'Determine the color of background, foreground and border color
    fpSpread1.ForeColor = RGB(0, 0, 0)
    fpSpread1.BackColor = vbButtonFace ' RGB(192, 192, 192)
    fpSpread1.GrayAreaBackColor = vbButtonFace
    fpSpread1.FontName = "Tahoma"
    fpSpread1.FontSize = 8
    fpSpread1.FontBold = False
    
    fpSpread1.TextTip = TextTipFloating
    fpSpread1.SetTextTipAppearance "Tahoma", 8, 0, 0, &HC0FFFF, &H0
    fpSpread1.CursorType = CursorTypeLockedCell
    fpSpread1.CursorStyle = CursorStyleArrow
    fpSpread1.NoBeep = True

    'Select a single cell
    fpSpread1.Col = 2
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = CellTypeButton
    fpSpread1.Lock = False
    'fpSpread1.TypeButtonText = "Next"
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\RIGHT.BMP"))
    fpSpread1.TypeButtonAlign = TypeButtonAlignLeft

    'Select a single cell
    fpSpread1.Col = 4
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = CellTypeButton
    fpSpread1.Lock = False
    'fpSpread1.TypeButtonText = "Previous"
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\LEFT.BMP"))
    fpSpread1.TypeButtonAlign = TypeButtonAlignRight
    
    'Select a single cell
    fpSpread1.Col = 6
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = CellTypeButton
    fpSpread1.Lock = False
    'fpSpread1.TypeButtonText = "Zoom"
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\ZOOM.BMP"))
    fpSpread1.TypeButtonAlign = TypeButtonAlignRight

    'Select a single cell
    fpSpread1.Col = 8
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = CellTypeButton
    fpSpread1.Lock = False
    'fpSpread1.TypeButtonText = "Print"
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\PRINT.BMP"))
    fpSpread1.TypeButtonAlign = TypeButtonAlignRight
    
    'Merge Cells from 9 to 15
    fpSpread1.AddCellSpan 9, 1, 6, 1
    fpSpread1.Col = 9
    fpSpread1.Row = 1
    fpSpread1.TypeHAlign = TypeHAlignCenter
    fpSpread1.Lock = True

    'Select a single cell
    fpSpread1.Col = 16
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = CellTypeButton
    fpSpread1.Lock = False
    'fpSpread1.TypeButtonText = "Close"
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\CLOSE.BMP"))
    fpSpread1.TypeButtonAlign = TypeButtonAlignRight
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetupToolbar", Err.Number, Err.Description
End Sub

Private Sub DisableButton(Col As Long, bitmapdirection As String)
On Error GoTo ErrHandle

'Disable specified button
    fpSpread1.ReDraw = False
    
    fpSpread1.Row = 1
    fpSpread1.Col = Col
    
    fpSpread1.Lock = True
    fpSpread1.TypeButtonTextColor = RGB(128, 128, 128)
    fpSpread1.Protect = True
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\" & bitmapdirection & "DIS.BMP"))
    
    fpSpread1.ReDraw = True
Exit Sub

ErrHandle:
    SaveErrorLog Me.Name, "DisableButton", Err.Number, Err.Description
End Sub

Private Sub EnableButton(Col As Long, bitmapdirection As String)
On Error GoTo ErrHandle

'Enable specified button
    fpSpread1.ReDraw = False
    
    fpSpread1.Row = 1
    fpSpread1.Col = Col
    
    fpSpread1.Lock = False
    fpSpread1.TypeButtonTextColor = RGB(0, 0, 0)
    fpSpread1.Protect = True
    Set fpSpread1.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\" & bitmapdirection & ".BMP"))
    
    fpSpread1.ReDraw = True
Exit Sub

ErrHandle:
    SaveErrorLog Me.Name, "EnableButton", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
On Error GoTo ErrHandle
    'fpSpread1.Move 0, 0, ScaleWidth, fpSpread1.Height
    fpspReport.Move 0, lblCaption.Height + fpSpread1.Height + 150, ScaleWidth, ScaleHeight - fpSpread1.Height - lblCaption.Height - 50
    SetFormCaption Me, imgCaption, lblCaption
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Resize", Err.Number, Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    fpspReport.hWndSpread = Empty
    Unload frmReportData
    frmReports.Enabled = True
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Unload", Err.Number, Err.Description
End Sub

Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim intSheet As Integer, intNumberOfCopies As Integer

    fpSpread1.Col = Col
    fpSpread1.Row = Row
    
    If fpSpread1.CellType = CellTypeButton Then
        Select Case Col
            Case 2  'Next
                cmdNext_Click
            Case 4  'Previous
                cmdPrevious_Click
            Case 6  'Zoom
                cmdZoom_Click
            Case 8  'Print
                cmdPrint_Click
            Case 16 'Close
                Unload Me
        End Select
    End If
End Sub

Private Sub UpdatePageCount()
Dim strPages As String

On Error GoTo ErrHandle
    'Page Count
    strPages = "Trang " & intCurrentPage & " trong " & intPageCount
    fpSpread1.SetText 9, 1, strPages
     
    'Enable - Disable buttons
    If intCurrentPage = 1 And intPageCount > 1 Then
        DisableButton 4, "LEFT"
        EnableButton 2, "RIGHT"
    ElseIf intCurrentPage = intPageCount And intPageCount > 1 Then
        DisableButton 2, "RIGHT"
        EnableButton 4, "LEFT"
    ElseIf intPageCount > 1 Then
        EnableButton 4, "LEFT"
        EnableButton 2, "RIGHT"
    End If
    'Set focus to preview control
    fpspReport.SetFocus
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "UpdatePageCount", Err.Number, Err.Description
End Sub

Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyN Then cmdNext_Click
    If Shift = 2 And KeyCode = vbKeyP Then cmdPrevious_Click
    If Shift = 2 And KeyCode = vbKeyZ Then cmdZoom_Click
    If Shift = 2 And KeyCode = vbKeyI Then cmdPrint_Click
    If Shift = 2 And KeyCode = vbKeyX Then cmdClose_Click
End Sub

Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPUSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Select Case Col
        Case 2 'cmdNext
            Set xmlNode = GetNodeByName(TAX_Utilities_v2.NodeCaption.selectSingleNode(Me.Name).childNodes, "cmdNext")
            If Not xmlNode Is Nothing Then
                'TipWidth = 800
                TipText = GetAttribute(xmlNode, "Caption") & " [Ctrl+N] "
            End If
        Case 4 'cmdPrevious
            Set xmlNode = GetNodeByName(TAX_Utilities_v2.NodeCaption.selectSingleNode(Me.Name).childNodes, "cmdPrevious")
            If Not xmlNode Is Nothing Then
                'TipWidth = 900
                TipText = GetAttribute(xmlNode, "Caption") & " [Ctrl+P] "
            End If
        Case 6 'cmdZoom
            Set xmlNode = GetNodeByName(TAX_Utilities_v2.NodeCaption.selectSingleNode(Me.Name).childNodes, "cmdZoom")
            If Not xmlNode Is Nothing Then
                'TipWidth = 1440
                TipText = GetAttribute(xmlNode, "Caption") & " [Ctrl+Z] "
            End If
        Case 8 'cmdPrint
            Set xmlNode = GetNodeByName(TAX_Utilities_v2.NodeCaption.selectSingleNode(Me.Name).childNodes, "cmdPrint")
            If Not xmlNode Is Nothing Then
                'TipWidth = 620
                TipText = GetAttribute(xmlNode, "Caption") & " [Ctrl+I] "
            End If
        Case 16 'cmdClose
            Set xmlNode = GetNodeByName(TAX_Utilities_v2.NodeCaption.selectSingleNode(Me.Name).childNodes, "cmdClose")
            If Not xmlNode Is Nothing Then
                'TipWidth = 620
                TipText = GetAttribute(xmlNode, "Caption") & " [Ctrl+X] "
            End If
    End Select

    If Not xmlNode Is Nothing Then
        ShowTip = True
        MultiLine = TextTipFetchMultilineAuto
    End If
    Set xmlNode = Nothing
    
End Sub

Private Sub fpspReport_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyN Then cmdNext_Click
    If Shift = 2 And KeyCode = vbKeyP Then cmdPrevious_Click
    If Shift = 2 And KeyCode = vbKeyZ Then cmdZoom_Click
    If Shift = 2 And KeyCode = vbKeyI Then cmdPrint_Click
    If Shift = 2 And KeyCode = vbKeyX Then cmdClose_Click
'    Debug.Print Shift & " : " & KeyCode
End Sub

Private Sub fpspReport_PageChange(ByVal Page As Long)
On Error GoTo ErrHandle
    fpspReport.PageCurrent = Page
    If intLocalCurrPage < Page Then
        intCurrentPage = intCurrentPage + 1
    ElseIf intLocalCurrPage > Page Then
        intCurrentPage = intCurrentPage - 1
    End If
    intLocalCurrPage = Page
    UpdatePageCount

Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "fpsReport_PageChanged", Err.Number, Err.Description
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

Private Function GetNodeByName(ByVal xmlNodeList As MSXML.IXMLDOMNodeList, ByVal strName As String) As MSXML.IXMLDOMNode
Dim xmlReturn As MSXML.IXMLDOMNode
    
For Each xmlReturn In xmlNodeList
    If GetAttribute(xmlReturn, "Name") = strName Then
        Set GetNodeByName = xmlReturn
        Exit Function
    End If
Next
Set GetNodeByName = Nothing
End Function
