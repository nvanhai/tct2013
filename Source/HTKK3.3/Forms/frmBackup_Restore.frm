VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackup_Restore 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4695
   ControlBox      =   0   'False
   HelpContextID   =   81213
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Sao l­u"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1350
      Width           =   1305
   End
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
      Left            =   3270
      TabIndex        =   2
      Top             =   1350
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   4
      Top             =   270
      Width           =   4695
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   345
         Left            =   4110
         TabIndex        =   3
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   540
         Width           =   4005
      End
      Begin VB.Label lblMsg 
         Caption         =   "Chän ®­êng dÉn tíi th­ môc sao l­u"
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
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   4005
      End
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Sao l­u - Phôc håi"
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
      Left            =   90
      TabIndex        =   6
      Top             =   0
      Width           =   3615
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmBackup_Restore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmBackup_Restore
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

Public bIsBackup As Boolean ' True - backup data, False - restore data
Public strPath As String

'****************************************************
'Description:CopyData procedure copy data from source directory
'   ot destination directory
'Input: pFrom - source
'       pTo -destination
'Output:
'Return:

'****************************************************

Public Sub CopyData(pFrom As String, pTo As String)
    On Error GoTo ErrorHandle
    Dim i As Integer
    
    CopyFolder pFrom, pTo
    'pkzip45 -add -silent -dir=relative -temp=c:\temp c:\a.zip D:\Working\WIP\Program\TaxProject\DataFiles\*
    'pkzip45 -Extract -silent -Over=All -dir=Relative -temp=c:\Temp c:\a.zip c:\Temp
    Unload Me
    If bIsBackup Then
        DisplayMessage "0007", msOKOnly, miInformation
    Else
        DisplayMessage "0008", msOKOnly, miInformation
    End If
    
    Exit Sub
    
ErrorHandle:

    If Err.Number = 70 Then
        DisplayMessage "0021", msOKOnly, miCriticalError
        Me.Show
    ElseIf Err.Number = 5 Then
        DisplayMessage "0025", msOKOnly, miCriticalError
        Me.Show
    ElseIf Err.Number = -2147024784 Then
        DisplayMessage "0041", msOKOnly, miCriticalError
        Me.Show
    Else
        SaveErrorLog Me.Name, "CopyData", Err.Number, Err.Description
    End If
    
End Sub

'****************************************************
'Description:CopyData procedure copy data from source directory
'   ot destination directory
'Input: pFrom - source
'       pTo -destination
'****************************************************

Public Sub ChangeAttributes(pFolder As String, pAttr As FileAttribute)
    On Error GoTo ErrorHandle
    Dim drv As Drive
    Dim fd As Folder
    Dim sfd As Folder
    Dim file As file
    Dim fso As New FileSystemObject
        
    ' get folder (rootfolder or folder)
    If fso.DriveExists(pFolder) Then
        Set drv = fso.GetDrive(pFolder)
        Set fd = drv.RootFolder
    Else
        Set fd = fso.GetFolder(pFolder)
    End If
    
    SetDefaultPropertyFolder fso.GetFolder(pFolder), Normal, True
        
    Set fd = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "ChangeAttributes", Err.Number, Err.Description
    
End Sub

Private Sub cmdBackup_Click()
    cmdOK_Click
End Sub

'****************************************************
'Description:cmdBrowser_Click procedure return the path to the
'   directory to backup or restore data
'****************************************************

Private Sub cmdBrowser_Click()
    On Error GoTo ErrorHandle
    
    Dim fBrowser As New frmBrowser
    txtPath.Text = fBrowser.getPath
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "cmdBrowser_Click", Err.Number, Err.Description
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

'****************************************************
'Description:cmdOK_Click procedure backup or restore data
'   Step 1: if backup then backup Datafiles directory
'   Step 2: if restore then restore Datafiles directory
'****************************************************

Private Sub cmdOK_Click()
    On Error GoTo ErrorHandle

    Dim fso As New FileSystemObject
    Dim path As String
    'MsgBox fso.FolderExists(txtPath.Text)

    Me.Hide
    If Len(Trim(txtPath.Text)) = 0 Then
        DisplayMessage "0023", msOKOnly, miInformation
        Me.Show
        txtPath.SetFocus

    ElseIf Not fso.FolderExists(txtPath.Text) Or txtPath.Text = "." Or txtPath.Text = "\" Or txtPath.Text = "/" Or InStr(txtPath, "\.") > 0 _
        Or InStr(txtPath, "/.") > 0 Or txtPath.Text = ".." Or InStr(txtPath, "..") > 0 Or InStr(txtPath, ".\") > 0 Or InStr(txtPath, "./") > 0 Then
        DisplayMessage "0010", msOKOnly, miInformation
        Me.Show
        txtPath.SetFocus
    Else
        If Right(txtPath.Text, 1) = "\" Then
            path = Left(txtPath.Text, Len(txtPath.Text) - 1)
        Else
            path = txtPath.Text
        End If
        If bIsBackup Then
            CopyData GetAbsolutePath("..\DataFiles"), path
        Else
            CopyData path, GetAbsolutePath("..\DataFiles")
        End If

    End If

    Exit Sub

ErrorHandle:

    If Err.Number = 52 Then
        DisplayMessage "0010", msOKOnly, miInformation
        Me.Show
        txtPath.SetFocus
    Else
        SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
    End If
End Sub

Private Sub cmdRestore_Click()
    cmdOK_Click
End Sub

Private Sub Form_Activate()
    txtPath.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    Dim s As String

    Me.Top = (frmSystem.Height - Me.Height) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    If bIsBackup Then
        SetControlCaption Me, "frmBackup"
'        cmdOK.HelpContextID = 81214
'        cmdClose.HelpContextID = 81214
'        txtPath.HelpContextID = 81214
'        cmdBrowser.HelpContextID = 81214
    Else
        SetControlCaption Me, "frmRestore"
'        cmdOK.HelpContextID = 81215
'        cmdClose.HelpContextID = 81215
'        txtPath.HelpContextID = 81215
'        cmdBrowser.HelpContextID = 81215
        
    End If

    hasActiveForm = True
    
    Exit Sub
     
ErrorHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
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

'****************************************************
'Description:Form_Resize procedure set caption for form
'****************************************************

Private Sub Form_Resize()
     SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TAX_Utilities_v2.NodeValidity = Nothing
    hasActiveForm = False
    frmTreeviewMenu.Show
End Sub

'*********************************
'Description: SetDefaultPropertyFolder procedure set attribute to
'             all of file and subfolder in folder.
'Input:
'       fldFolder: folder.
'       fleAttr:   Attribute set to files and subfolders
'*********************************
Private Sub SetDefaultPropertyFolder(ByRef fldFolder As Folder, ByVal fleAttr As FileAttribute, Optional blnDestinationDirectory As Boolean = False)
    
    On Error GoTo ErrHandle
    
    Dim fld As Folder
    Dim fle As file
        
    If Not IsValidFolder(fldFolder.Name) And Not blnDestinationDirectory Then _
        Exit Sub
    
    'Change attr to subfolders
    For Each fld In fldFolder.SubFolders
        SetDefaultPropertyFolder fld, fleAttr
    Next
    
    'Change attr to subFiles
    For Each fle In fldFolder.Files
        fle.Attributes = fleAttr
    Next
    
    'Change attr to folder
    fldFolder.Attributes = fleAttr
        
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "SetDefaultPropertyFolder", Err.Number, Err.Description
End Sub

'******************************
'Description: IsValidFolder function check whether
'             name of folder is valid
'Input: folder name
'******************************
Private Function IsValidFolder(ByVal strFolderName As String) As Boolean

If Not IsNumeric(strFolderName) Then _
    Exit Function
If Len(strFolderName) <> 10 And Len(strFolderName) <> 13 Then _
    Exit Function
If Not IsValidTaxId(strFolderName) Then _
    Exit Function
IsValidFolder = True
End Function

'******************************
'Description: CopyFolder procedure copy a folder from
'             specified location to other location.
'Input:
'       strPathFrom: path folder want to copy from
'       strPathTo  : path folder want to copy to
'******************************
Private Sub CopyFolder(ByVal strPathFrom As String, ByVal strPathTo As String)
    
    Dim fso As New FileSystemObject
    Dim fld As Folder, fle As file, fldFolder As Folder
    
    
    If fso.DriveExists(strPathFrom) Then
        Set fldFolder = fso.GetDrive(strPathFrom).RootFolder
    Else
        Set fldFolder = fso.GetFolder(strPathFrom)
    End If
    
    'Copy all of subfolders
    For Each fld In fldFolder.SubFolders
        If IsValidFolder(fld.Name) Then
            If fso.FolderExists(strPathTo & "\" & fld.Name) Then
                fso.DeleteFolder strPathTo & "\" & fld.Name, True
            End If
            fso.CreateFolder strPathTo & "\" & fld.Name
            CopyFolder fld.path, strPathTo & "\" & fld.Name
        End If
    Next
    
    'Copy all of files
    For Each fle In fldFolder.Files
        If InStr(1, fle.Name, ".xml") > 0 Or InStr(1, fle.Name, ".dtd") > 0 Then
            If fso.FileExists(strPathTo & "\" & fle.Name) Then _
                fso.DeleteFile strPathTo & "\" & fle.Name, True
            fso.CopyFile strPathFrom & "\" & fle.Name, strPathTo & "\" & fle.Name, True
        End If
    Next
    
    Set fso = Nothing
End Sub
