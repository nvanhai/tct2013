VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBrowser 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5505
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   9075
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4425
      End
      Begin VB.FileListBox File1 
         Height          =   4575
         Hidden          =   -1  'True
         Left            =   4545
         MultiSelect     =   2  'Extended
         Pattern         =   "*.ocx;*.dll"
         System          =   -1  'True
         TabIndex        =   3
         Top             =   885
         Width           =   4455
      End
      Begin VB.DirListBox Dir1 
         Height          =   4815
         Left            =   90
         TabIndex        =   1
         Top             =   630
         Width           =   4455
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   270
         Width           =   8895
      End
   End
   Begin MSForms.Label lblCaption 
      Height          =   285
      Left            =   690
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      ForeColor       =   -2147483634
      Size            =   "3836;503"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   345
      Left            =   90
      Top             =   60
      Width           =   5865
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   7740
      TabIndex        =   6
      Top             =   5940
      Width           =   1305
      Caption         =   "Cancel"
      Size            =   "2302;661"
      Accelerator     =   72
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   6390
      TabIndex        =   5
      Top             =   5940
      Width           =   1305
      Caption         =   "Ok"
      Size            =   "2302;661"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrStrFileNames() As String

Private Sub Combo1_Change()
    File1.Pattern = Combo1.Text
End Sub

Private Sub Combo1_Click()
    File1.Pattern = Combo1.Text
End Sub

Private Function FileSelected() As Boolean
    Dim i As Long
    
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) = True Then
            FileSelected = True
            Exit For
        End If
    Next
End Function

Private Sub cmdOk_Click()
    Dim strFolderPath As String
    Dim intCtrl As Integer
    
    If FileSelected = False Then
        DisplayMessage "0058", msOKOnly, miInformation, Me.caption
        Exit Sub
    End If
    
    ReDim arrStrFileNames(0)
    
    strFolderPath = Dir1.path
    If Right(strFolderPath, 1) = "\" Then
        strFolderPath = Left(strFolderPath, Len(strFolderPath) - 1)
    End If
    strFolderPath = strFolderPath & "\"
    
    For intCtrl = 0 To File1.ListCount - 1
        If File1.Selected(intCtrl) Then
            ReDim Preserve arrStrFileNames(UBound(arrStrFileNames) + 1)
            arrStrFileNames(UBound(arrStrFileNames)) = strFolderPath & File1.List(intCtrl)
        End If
    Next intCtrl
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    ReDim arrStrFileNames(0)
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dim strDrive As String

On Error GoTo ErrorHandle
    strDrive = Left(Dir1.path, InStr(Dir1.path, "\") - 1)
    Dir1.path = Drive1.Drive
    Exit Sub
ErrorHandle:
    If Err.Number = 68 Then
        DisplayMessage "0031", msOKOnly, miCriticalError
        Drive1.Drive = strDrive
    ElseIf Err.Number = 419 Then
        DisplayMessage "0060", msOKOnly, miCriticalError
        Drive1.Drive = strDrive
    Else
        SaveErrorLog Me.Name, "Drive1_Change", Err.Number, Err.Description
    End If
    
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Shift = 0) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyExecute) Then
        cmdOk_Click
    End If
End Sub

Private Sub Form_Load()
    SetControlCaption Me
    With Combo1
        .Clear
        .AddItem "*.dat"
        .ListIndex = 0
    End With
    File1.Pattern = Combo1.Text
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBrowser = Nothing
End Sub

Public Function GetFileNames() As String()
    Me.Show vbModal
    GetFileNames = arrStrFileNames
End Function
