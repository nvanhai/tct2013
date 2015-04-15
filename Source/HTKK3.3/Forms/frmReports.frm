VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5775
   ControlBox      =   0   'False
   HelpContextID   =   81211
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Xem tr­íc"
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
      Left            =   3090
      TabIndex        =   3
      Top             =   2070
      Width           =   1305
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   4440
      TabIndex        =   4
      Top             =   2070
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&In"
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
      Left            =   1740
      TabIndex        =   2
      Top             =   2070
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   5775
      Begin VB.CheckBox chkDieuChinh 
         Caption         =   "In th«ng tin ®iÒu chØnh"
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
         Left            =   1440
         TabIndex        =   11
         Top             =   1305
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.ComboBox cboPrinters 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   4185
      End
      Begin VB.TextBox txtPages 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   870
         Width           =   4185
      End
      Begin VB.TextBox txtNumberOfCopies 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Text            =   "1"
         Top             =   540
         Width           =   4185
      End
      Begin VB.Label lblPages 
         Caption         =   "Trang in"
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
         Left            =   150
         TabIndex        =   9
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label lblNumberOfCopies 
         Caption         =   "Sè b¶n in"
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
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lblSelectPrinter 
         Caption         =   "Chän m¸y in"
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
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1275
      End
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "In tê khai"
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
      Left            =   960
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmReportData
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

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Dim udtPrinter As New clsDefaultPrinter
    Dim arrStrPages As Variant
    Dim lPrevPaperSize As Long
        
    ' Truong hop in dieu chinh bo sung cac to khai quyet toan TNCN thi phai check vao danh sach dieu chinh
    ' Neu khong mess de bao cho chon lai danh sach
    ' Neu ko chon la chkDieuChinh thi in tat
    If flgPrintBoSung = False And chkDieuChinh.value = 1 Then
        DisplayMessage "0170", msOKOnly, miCriticalError
        Exit Sub
    End If
    ' Ket thuc kiem tra trong truong hop in dieu chinh cac to khai quyet toan TNCN
        
    If CheckPrinter And CheckPrintedPages(arrStrPages) Then
        ' Set chuoi ma vach duoc in ra PDF ve vbNullString
        ' strBarcodeInPDF Chi dung cho truong hop iHTKK
        strBarcodeInPDF = vbNullString
        
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
'        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(1), "Caption") = "04-1/TNCN" Then
'        Printer.PaperSize = vbPRPSA3
'        End If
        
                ' BC26
'        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Caption") = "BC26-AC" Then
'
'            Printer.PaperSize = vbPRPSA3
'            Printer.Orientation = vbPRORLandscape
'        End If
        ' end

        
        ' Doi voi tat ca cac to khai tru quyet toan TNCN, moi phai chuan bi den ma vach de in
'        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "17" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "41" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "42" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "43" Then
            CreateExcelBook
'        End If
        
        frmReportData.SetPrintedPages (arrStrPages)
        'frmReportData.Show 1
        frmPreview.Show 1
        
                
        Set udtPrinter = Nothing
        Printer.PaperSize = lPrevPaperSize
        
    End If
End Sub
Private Sub cmdPrint_Click()
Dim arrStrPages As Variant
Dim intNumberOfCopies As Integer
Dim udtPrinter As New clsDefaultPrinter
Dim lPrevPaperSize As Long

On Error GoTo ErrHandle
    
    ' Truong hop in dieu chinh bo sung cac to khai quyet toan TNCN thi phai check vao danh sach dieu chinh
    ' Neu khong mess de bao cho chon lai danh sach
    ' Neu ko chon la chkDieuChinh thi in tat
    If flgPrintBoSung = False And chkDieuChinh.value = 1 Then
        DisplayMessage "0170", msOKOnly, miCriticalError
        Exit Sub
    End If
    ' Ket thuc kiem tra trong truong hop in dieu chinh cac to khai quyet toan TNCN
    
    If CheckPrinter And CheckPrintedPages(arrStrPages) Then
        ' Set chuoi ma vach duoc in ra PDF ve vbNullString
        ' strBarcodeInPDF Chi dung cho truong hop iHTKK
        strBarcodeInPDF = vbNullString
        
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
        ' in bao cao BC26 bang giay A3
'        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Caption") = "BC26-AC" Then
'            Printer.PaperSize = vbPRPSA3
'            Printer.Orientation = vbPRORLandscape
'        End If
        
        'Check Ready of printer
        If Not IsPrinterReady Then
            DisplayMessage "0057", msOKOnly, miCriticalError
            Set udtPrinter = Nothing
            Printer.PaperSize = lPrevPaperSize
            Exit Sub
        End If
        ' Doi voi tat ca cac to khai tru quyet toan TNCN, moi phai chuan bi den ma vach de in
'        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "17" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "41" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "42" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "43" Then
            CreateExcelBook
'        End If
        frmReportData.SetPrintedPages (arrStrPages)
        For intNumberOfCopies = 1 To CInt(txtNumberOfCopies.Text)
                frmReportData.PrintTax
        Next intNumberOfCopies
        Unload frmReportData
        Unload Me
        
        Set udtPrinter = Nothing
        Printer.PaperSize = lPrevPaperSize
    End If
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdPrint_Click", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    SetCboPrinter
    SetControlCaption Me
    
    Me.Left = Round((Screen.Width - Me.Width) / 2)
    Me.Top = Round((Screen.Height - Me.Height) / 2) - 200
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub SetCboPrinter()
On Error GoTo ErrHandle
    Dim intCtrl As Integer, intDefault As Integer
    
    intDefault = -1
    For intCtrl = 0 To Printers.count - 1
        cboPrinters.AddItem Printers(intCtrl).DeviceName
        If Printers(intCtrl).DeviceName = Printer.DeviceName Then
            intDefault = intCtrl
        End If
    Next

    cboPrinters.ListIndex = intDefault
    If intDefault = -1 Then
        cmdPreview.Enabled = False
        cmdPrint.Enabled = False
    End If
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetCboPrinter", Err.Number, Err.Description

End Sub

'****************************************************
'Description:CheckPrinter function check and set name of printer
'Return:
'       True if name of printer exist
'       False otherwise
'****************************************************
Public Function CheckPrinter() As Boolean
On Error GoTo ErrHandle
    If Not IsNumeric(txtNumberOfCopies.Text) Then
        DisplayMessage "0033", msOKOnly, miInformation
        CheckPrinter = False
        txtNumberOfCopies.SetFocus
        Exit Function
    End If
    CheckPrinter = True
    strPrinterName = cboPrinters.Text
    
Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "CheckPrinter", Err.Number, Err.Description
End Function

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
'    frmInterfaces.Enabled = True
End Sub

Private Sub txtNumberOfCopies_Change()
       
    On Error GoTo ErrorHandle
    
    Static strValue As String

    If Len(txtNumberOfCopies.Text) <> 0 And Not IsNumeric(txtNumberOfCopies.Text) Then
        txtNumberOfCopies.Text = strValue
    Else
        strValue = txtNumberOfCopies.Text
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "txtNumberOfCopies_Change", Err.Number, Err.Description
End Sub

Private Sub txtNumberOfCopies_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandle
    Dim sNumber As String
    sNumber = "0123456789"
    
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "txtNumberOfCopies_KeyPress", Err.Number, Err.Description
End Sub

Private Sub txtPages_Change()
    On Error GoTo ErrorHandle

    Dim strNumber As String, intCtrl As Integer
    Dim blnValid As Boolean
    Static strValue As String

    strNumber = "0123456789,-"

    blnValid = True
    For intCtrl = 1 To Len(txtPages.Text)
        If InStr(1, strNumber, Mid(txtPages.Text, intCtrl, 1)) = 0 Then
            blnValid = False
            Exit For
        End If
    Next intCtrl

    If Len(txtPages.Text) <> 0 And Not blnValid Then
        txtPages.Text = strValue
    Else
        strValue = txtPages.Text
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "txtPages_Change", Err.Number, Err.Description
End Sub

Private Sub txtPages_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandle
    Dim sNumber As String
    sNumber = "0123456789,-"
    
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "txtPages_KeyPress", Err.Number, Err.Description
End Sub

Public Function CheckPrintedPages(arrStrPages As Variant) As Boolean
Dim strPages As String
Dim intCtrl As Integer, blnReturn As Boolean
Dim arrStrTemp() As String

strPages = Replace(txtPages.Text, " ", "")
If strPages = "" Then
    ReDim arrStrTemp(0)
    arrStrPages = arrStrTemp
    arrStrPages(0) = "All"
    CheckPrintedPages = True
    Exit Function
End If

arrStrTemp = Split(strPages, ",")

blnReturn = True
For intCtrl = 0 To UBound(arrStrTemp)
    If Not IsValidElement(arrStrTemp(intCtrl)) Then
        blnReturn = False
        Exit For
    End If
Next

If Not blnReturn Then
    CheckPrintedPages = False
    DisplayMessage "0040", msOKOnly, miCriticalError
    txtPages.SetFocus
    Exit Function
End If

'arrStrTemp = Split(strPages, ",")
arrStrPages = arrStrTemp

CheckPrintedPages = blnReturn
End Function

Private Function IsValidElement(strValue As String) As Boolean
Dim strValid As String, lCtrl As Long
Dim blnReturn As Boolean
Dim arrStrTemp() As String

strValid = "0123456789-"
blnReturn = True

For lCtrl = 1 To Len(strValue)
    If InStr(1, strValid, Mid$(strValue, lCtrl, 1)) = 0 Then
        blnReturn = False
        Exit For
    End If
Next lCtrl

If strValue = "0" Or strValue = vbNullString Then blnReturn = False
If InStr(1, strValue, "-") <> 0 Then
    arrStrTemp = Split(strValue, "-")
    If UBound(arrStrTemp()) <> 1 Then
        blnReturn = False
        GoTo exitFunction
    End If
    
    If Not (IsValidElement(arrStrTemp(0)) And IsValidElement(arrStrTemp(1))) Then
        blnReturn = False
    Else
        If CInt(arrStrTemp(0)) > CInt(arrStrTemp(1)) Then
            blnReturn = False
        End If
    End If
End If

exitFunction:
If Not blnReturn Then
    IsValidElement = False
    Exit Function
End If

IsValidElement = True
End Function

Public Function IsPrinterReady() As Boolean
    On Error GoTo ErrHandler
    Printer.Print
    PrinterKillDoc
    IsPrinterReady = True
    Exit Function
ErrHandler:
End Function
