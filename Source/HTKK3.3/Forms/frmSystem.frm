VERSION 5.00
Begin VB.MDIForm frmSystem 
   BackColor       =   &H8000000C&
   Caption         =   "Hç trî kª khai"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14310
   Icon            =   "frmSystem.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   14310
      TabIndex        =   0
      Top             =   0
      Width           =   14310
      Begin VB.Label lblUserInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "M· sè thuÕ: "
         BeginProperty Font 
            Name            =   "DS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   40
         Width           =   11595
      End
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmSystem
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
Public clickexit As Boolean
Private Sub MDIForm_Load()
    Dim xmlDocCaption As New MSXML.DOMDocument
    Dim apMyApp As App

    Set apMyApp = App
    If apMyApp.PrevInstance Then
        End
    End If
    
    LoadListMessage
    
    Me.Picture = LoadPicture(GetAbsolutePath("..\Pictures\bg.gif"))
    Me.BackColor = RGB(168, 9, 13)
    Picture1.BackColor = RGB(168, 9, 13)
    Me.icon = LoadPicture(GetAbsolutePath("..\Pictures\Desktop.ICO"))
    xmlDocCaption.Load App.path & "\Caption.xml"
    TAX_Utilities_v2.NodeCaption = xmlDocCaption.documentElement
    Set xmlDocCaption = Nothing
    
    SetControlCaption frmSystem
    
    clickexit = False
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
'Description:MDIForm_Resize adjust the size and position
'    of TreeviewMenu when MDI form change the size



'Input:
'Output:
'Return:

'****************************************************
Private Sub MDIForm_Resize()
    On Error GoTo ErrorHandle
    Dim a As Double
'    Dim xmlDocument As New MSXML.DOMDocument
'    Dim xmlNode As MSXML.IXMLDOMNode
'
''    frmTreeviewMenu.Top = 0
''    frmTreeviewMenu.Left = 0
''    frmTreeviewMenu.Width = 3150 'frmSystem.ScaleWidth / 4
''    frmTreeviewMenu.Height = frmSystem.ScaleHeight
'
''    If Screen.activeForm.Name <> frmTreeviewMenu.Name Then
''        Screen.activeForm.Top = (frmSystem.ScaleHeight - Screen.activeForm.Height) \ 2 + 100
''        Screen.activeForm.Left = (frmSystem.Width - Screen.activeForm.Width) \ 2 - 100
''        If Screen.activeForm.Top < 0 Then Screen.activeForm.Top = 0
''        If Screen.activeForm.Left < 0 Then Screen.activeForm.Left = 0
''    End If
'
'    xmlDocument.Load App.path & "\Menu.xml"
'    Set xmlNode = xmlDocument.getElementsByTagName("Root").Item(0)
'
'    If Trim(xmlNode.Attributes.getNamedItem("FirstTimeRunID").nodeValue) <> "" Then
'        frmTreeviewMenu.ProcessMenuAction Trim(xmlNode.Attributes.getNamedItem("FirstTimeRunID").nodeValue)
'        xmlNode.Attributes.getNamedItem("FirstTimeRunID").nodeValue = ""
'        xmlDocument.save App.path & "\Menu.xml"
'    End If
'    Set xmlDocument = Nothing
'    Set xmlNode = Nothing
    If (Me.Width = 15480) Then
        Me.Picture = LoadPicture(GetAbsolutePath("..\Pictures\bg.gif"))
    Else
        Me.Picture = LoadPicture(GetAbsolutePath("..\Pictures\bg.gif"))
    End If
    
    clickexit = False
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "MDIForm_Resize", Err.Number, Err.Description
End Sub
'****************************************************
'Description:MDIForm_Unload release the variable common
'



'Input:
'Output:
'Return:

'****************************************************
Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    Set xmlHeaderData = Nothing
    Set xmlNodeListMenu = Nothing
    TAX_Utilities_v2.NodeMessage = Nothing
    TAX_Utilities_v2.NodeCaption = Nothing
    TAX_Utilities_v2.NodeMenu = Nothing
    TAX_Utilities_v2.NodeValidity = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "MDIForm_Unload", Err.Number, Err.Description
End Sub


'****************************************************
'Description:LoadListMessage procedure load messages form Message.xml
'****************************************************
Public Sub LoadListMessage()
    On Error GoTo ErrorHandle
    
    Dim xmlDocument As New MSXML.DOMDocument
    
    xmlDocument.Load App.path & "\Message.xml"
    
    TAX_Utilities_v2.NodeMessage = xmlDocument.getElementsByTagName("Message").Item(0).childNodes
    
    Set xmlDocument = Nothing
    
    Exit Sub
 
ErrorHandle:
    SaveErrorLog Me.Name, "loadListMessage", Err.Number, Err.Description
End Sub
