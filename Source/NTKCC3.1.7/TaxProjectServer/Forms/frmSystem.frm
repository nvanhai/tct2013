VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.MDIForm frmSystem 
   BackColor       =   &H8000000C&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin MSForms.CheckBox chkSaveQHS 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   3555
         VariousPropertyBits=   746588183
         BackColor       =   -2147483636
         DisplayStyle    =   4
         Size            =   "6271;503"
         Value           =   "1"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkSaveQuestion 
         Height          =   285
         Left            =   8160
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   3555
         VariousPropertyBits=   746588179
         BackColor       =   -2147483636
         DisplayStyle    =   4
         Size            =   "6271;503"
         Value           =   "1"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblUser 
         Height          =   225
         Left            =   210
         TabIndex        =   1
         Top             =   30
         Width           =   3825
         ForeColor       =   -2147483640
         BackColor       =   -2147483636
         VariousPropertyBits=   8388627
         Size            =   "6747;397"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : FIS - CMC Software Solution
' Project           : Du an ho tro ke khai thue
' Package           : Interface
' Form, Module
'   or Class name   : frmSystem
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
Public clickexit As Boolean


Private Sub MDIForm_Load()
On Error GoTo ErrorHandle
    Dim xmlDocCaption As New MSXML.DOMDocument
    
    'Load list of messages
    LoadListMessage
    
    Me.Picture = LoadPicture(GetAbsolutePath("..\Pictures\bg.bmp"))
    Me.BackColor = RGB(74, 121, 198)
    Me.icon = LoadPicture(GetAbsolutePath("..\Pictures\icon.ICO"))
    xmlDocCaption.Load App.path & "\Caption.xml"
    TAX_Utilities_Svr_New.NodeCaption = xmlDocCaption.documentElement
    
    frmLogin.Top = (frmSystem.Height - frmLogin.Height) / 2
    frmLogin.Left = (frmSystem.Width - frmLogin.Width) / 2
    
    'Set caption to controls of MDI form
    SetControlCaption Me
    frmLogin.Show
    Set xmlDocCaption = Nothing
    
    clickexit = False
 Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "MDIForm_QueryUnload", Err.Number, Err.Description
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandle
    
    If clickexit Then
        Exit Sub
    End If
    
    If hasActiveForm = True Then
        Cancel = 1
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "MDIForm_QueryUnload", Err.Number, Err.Description
End Sub


'****************************************************
'Description:MDIForm_Unload release the variable common
'
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:

'****************************************************
Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    Set xmlHeaderData = Nothing
    Set xmlNodeListMenu = Nothing
    TAX_Utilities_Svr_New.NodeMessage = Nothing
    TAX_Utilities_Svr_New.NodeCaption = Nothing
    TAX_Utilities_Svr_New.NodeMenu = Nothing
    TAX_Utilities_Svr_New.NodeValidity = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "MDIForm_Unload", Err.Number, Err.Description
End Sub

'****************************************************
'Description:LoadListMessage procedure load messages form Message.xml
'
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:
'****************************************************
Public Sub LoadListMessage()
    On Error GoTo ErrorHandle
    
    Dim xmlDocument As New MSXML.DOMDocument
    
    xmlDocument.Load App.path & "\Message.xml"
    
    TAX_Utilities_Svr_New.NodeMessage = xmlDocument.getElementsByTagName("Message").Item(0).childNodes
    
    Set xmlDocument = Nothing
    
    Exit Sub
 
ErrorHandle:
    SaveErrorLog Me.Name, "loadListMessage", Err.Number, Err.Description
End Sub

