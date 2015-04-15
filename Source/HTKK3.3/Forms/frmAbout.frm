VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
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

Private Sub cmdExit_Click()
   'Unload Me
End Sub

Private Sub Form_Load()
    'Me.Top = frmSystem.Top + (frmSystem.Height - Me.Height) / 2
    'Me.Left = frmSystem.Left + (frmSystem.Width - Me.Width) / 2
    'SetControlCaption Me, "frmAbout"
End Sub

Private Sub Form_Resize()
     'SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'TAX_Utilities_v2.NodeValidity = Nothing
    'frmTreeviewMenu.Show
End Sub
