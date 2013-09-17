VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   75
      TabIndex        =   1
      Top             =   390
      Width           =   4980
      Begin MSForms.TextBox txtPassword 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   630
         Width           =   2175
         VariousPropertyBits=   746604571
         Size            =   "3836;556"
         PasswordChar    =   42
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblPassword 
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   690
         Width           =   705
         VariousPropertyBits=   276824091
         Caption         =   "Password"
         Size            =   "1244;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtUsername 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   300
         Width           =   2175
         VariousPropertyBits=   746604571
         Size            =   "3836;556"
         Value           =   "ADMIN"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblUsername 
         Height          =   195
         Left            =   255
         TabIndex        =   2
         Top             =   345
         Width           =   1125
         VariousPropertyBits=   276824091
         Caption         =   "Username"
         Size            =   "1984;344"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.CommandButton cmdVAT 
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   1695
      Width           =   2025
      Caption         =   "VAT"
      Size            =   "3572;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   90
      Width           =   1965
      ForeColor       =   -2147483634
      Size            =   "3466;450"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   315
      Left            =   30
      Top             =   30
      Width           =   3915
   End
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3750
      TabIndex        =   6
      Top             =   1695
      Width           =   1305
      Caption         =   "Exit"
      Size            =   "2302;635"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   360
      Left            =   2295
      TabIndex        =   5
      Top             =   1695
      Width           =   1305
      Caption         =   "Login"
      Size            =   "2302;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmLogin"
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
' Start date        : 21/11/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : TuanLM
' Integrate         :
' Project manager   : ThietKN
' Last modify       :
' Reason of modify  :
'******************************************************
Option Explicit

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lparam As Long
  iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
            
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
            
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Sub cmdClose_Click()
    Unload Me
    Unload frmSystem
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrorHandle
   
    If Len(txtUsername.Text) = 0 Then
        DisplayMessage "0056", msOKOnly, miInformation
        txtUsername.SetFocus
        Exit Sub
    End If
    
    'Quangtv
    
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
     
    Dim path As String
    Dim txt  As String
     
    If Not fs.FileExists(App.path & "\config.txt") Then
        If Len(txtPassword.Text) = 0 Then
            DisplayMessage "0088", msOKOnly, miInformation
            txtPassword.SetFocus
            Exit Sub
        Else
            txt = spathVat
            txt = txt & "," & txtUsername.Text
            txt = txt & "," & txtPassword.Text
            txt = txt & "," & spathQHSCC
            Call WritePathFile(App.path & "\config.txt", txt)
        End If

    Else
        Call ReadPathFile(App.path & "\config.txt")

        If txtUsername.Text <> strFile(1) Or txtPassword.Text <> strFile(2) Then
            DisplayMessage "0089", msOKOnly, miInformation
            Exit Sub
        End If
    End If
            
    '            txt = spathVat
    '            txt = txt & "," & txtUsername.Text
    '            txt = txt & "," & txtPassword.Text
    '            Call WritePathFile(App.path & "\config.txt", txtDir.Text)
    'spathVat = txtDir.Text
    
    '26/12/2011
    'dntai
    'get user login
    If spathVat = "" Then
        DisplayMessage "0092", msOKOnly, miCriticalError
        frmThamso.Show
        Exit Sub
    End If

    strUserID = txtUsername.Text

    Select Case IsValidUser()

        Case 2
            GetDataInfor
            'Set user name to system caption
            frmSystem.lblUser.caption = Mid$(frmSystem.lblUser.caption, 1, InStr(1, frmSystem.lblUser.caption, ":") + 1) & strUserName

            '********************************
            ' ThanhDX added
            ' Date: 27/04/06
            ' Check version of application
            If Not CheckVersion Then
                '                Unload Me
                '                frmLogin.Show
                Exit Sub
            End If

            '********************************
            Unload Me
            frmTreeviewMenu.Show

        Case 1

            If DisplayMessage("0059", msYesNo, miQuestion) = mrYes Then
                txtPassword.SetFocus
            Else
                Unload Me
                Unload frmSystem
                Exit Sub
            End If

        Case 0

            If DisplayMessage("0040", msYesNo, miQuestion) = mrYes Then
                txtPassword.SetFocus
            Else
                Unload Me
                Unload frmSystem
                Exit Sub
            End If

    End Select

    If spathVat = "" Then
        frmThamso.Show
         
    End If

    If spathQHSCC = "" Then
        frmThamsoQHSCC.Show
    End If

    If Not fs.FolderExists(spathVat) Then
        If Trim(spathVat) <> vbNullString Then
            DisplayMessage "0092", msOKOnly, miInformation
        End If

        frmThamso.Show
    End If
    
    If Not fs.FolderExists(spathQHSCC) Then
        If Trim(spathQHSCC) <> vbNullString Then
            DisplayMessage "0119", msOKOnly, miInformation
        End If

        frmThamsoQHSCC.Show
    End If

    If clsDAO.Connected = True Then
    
        clsDAO.Disconnect
    End If

    Exit Sub
ErrorHandle:
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
End Sub

Private Sub cmdOpen_Click()
    Dim path As String
    path = BrowseFolder("")
    'txtDir.Text = path
End Sub
Public Function BrowseFolder(szDialogTitle As String) As String
  Dim X As Long, bi As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer
  
    With bi
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = vbNullString
    End If
End Function

Private Sub cmdVAT_Click()
frmThamso.Show
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandle
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    SetControlCaption Me, "frmLogin"
    If fs.FileExists(App.path & "\config.txt") Then
        Call ReadPathFile(App.path & "\config.txt")
        txtUsername.Text = strFile(1)
        spathVat = strFile(0)
        spathQHSCC = strFile(3)
    End If
    'txtPassword.Text = "admin"
Exit Sub
ErrorHandle:
    MsgBox "loi" & Err.Description
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub

'****************************************************
'Description:IsValidUser function check if user and password are valid
'   Step 1: Show frmPeriod to user can chose the priod
'   Step 2: Show frmInterfases
'Author:TuanLM
'Modify by:
'Date:11/10/2005
'Input:
'Output:
'Return:
'****************************************************
Private Function IsValidUser() As Integer
On Error GoTo ErrorHandle
    Dim userid As String
    Dim password As String
    Dim clsConvert  As New clsUnicodeConvert
    
    Dim rec As ADODB.Recordset
    Dim strSQL As String
    Dim cmd As New ADODB.Command
    
    'connect to database BMT
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "BMT", "LOGIN_USER", "LOGIN_USER"
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
'        clsDAO.Connect
'        Me.MousePointer = vbDefault
'        frmSystem.MousePointer = vbDefault
'    End If
    
    'set key trong BMT, call prc_get_key
'    cmd.ActiveConnection = clsDAO.Connection
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "{call BMT_PCK_BMHT.prc_get_key()}"
'    cmd.Execute
'    Set cmd = Nothing
    
    'create slq query check username and password
'    userid = clsConvert.Convert(txtUsername.Text, UNICODE, TCVN)
'    password = clsConvert.Convert(txtPassword.Text, UNICODE, TCVN)
'    strSQL = "SELECT nvl(BMT_PCK_BMHT.fnc_check_login('" & _
'                userid & "','" & password & "'),-1)  result FROM dual"
    
    'check username and password
'    Set rec = clsDAO.Execute(strSQL)
    
'    If rec.Fields(0).Value > 0 Then
'    '***********************************
'    'ThanhDX modified
'    'Date:18/04/06
'    ' Them truong ten_nguoisudung dua vao QLT
'        'get username
'        strSQL = "SELECT ten_nsd, mo_ta FROM bmt_nsd WHERE ten_nsd='" & userid & "' " & _
'        "AND MA_NSD IN (SELECT MA_NSD FROM bmt_nsd_nhom " & _
'        "WHERE MA_NHOM IN (SELECT MA_NHOM FROM BMT_NHOM_CHUC_NANG " & _
'        "WHERE MA_CHUC_NANG IN (SELECT MA_CHUC_NANG FROM bmt_chuc_nang " & _
'        "WHERE ma_ud = 'HTKK')))"
'        Set rec = clsDAO.Execute(strSQL)
''*******************************
''ThanhDX added
''Modify date: 12/12/2005
'        If rec.Fields.Count > 0 Then
'            IsValidUser = 2
'            'get User ID
'            strUserID = rec.Fields(0).Value
'            'get User name
'            strUserName = clsConvert.Convert(rec.Fields(1).Value, TCVN, UNICODE)
''***********************************
''***********************************
''ThanhDX modified
''Modify date: 01/8/2005
'
''            'get cqt id (Cuc thue)
''            strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_TINH'"
''            Set rec = clsDAO.Execute(strSQL)
''
''            strTaxOfficeId = clsConvert.Convert(rec.Fields(0).Value, TCVN, UNICODE) & "00"
'
'            'Kiem tra ma co quan thue la cua chi cuc hay cuc thue.
'            strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_CQT'"
'            Set rec = clsDAO.Execute(strSQL)
'
'            If rec Is Nothing Then
'                'get cqt id (Cuc thue)
'                strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_TINH'"
'                Set rec = clsDAO.Execute(strSQL)
'            End If
                        
strTaxOfficeId = "123" 'clsConvert.Convert(rec.Fields(0).Value, TCVN, UNICODE)
            
'            If Len(Trim(strTaxOfficeId)) = 3 Then
'                strTaxOfficeId = strTaxOfficeId & "00"
'            End If
''***********************************
'        Else
'            IsValidUser = 1
'        End If
'        '*******************************
'    Else
'        IsValidUser = 0
'    End If
    'Longvh
    IsValidUser = 2
'    rec.Close
'    Set rec = Nothing
    
    Exit Function
ErrorHandle:
    Me.MousePointer = vbDefault
    frmSystem.MousePointer = vbDefault
    rec.Close
    Set rec = Nothing
    SaveErrorLog Me.Name, "IsValidUser", Err.Number, Err.Description
End Function

Private Sub GetDataInfor()
On Error GoTo ErrorHandle
    Dim userid As String
    Dim clsConvert  As New clsUnicodeConvert
    Dim rec As ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim para As New ADODB.Parameter
    
'    'connect to database BMT
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "BMT", "LOGIN_USER", "LOGIN_USER"
'        clsDAO.Connect
'    End If
'
'    'set key trong BMT, call prc_get_key
'    cmd.ActiveConnection = clsDAO.Connection
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BMT_PCK_BMHT.prc_get_key"
'
'
'    cmd.Execute
'    Set cmd = Nothing
'
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = clsDAO.Connection
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BMT_PCK_BMHT.Prc_Get_App_Owner"
'    cmd.Parameters.Append cmd.CreateParameter("P_USER_NAME", adVarChar, adParamOutput, 4000)
'    cmd.Parameters.Append cmd.CreateParameter("P_PASSWORD", adVarChar, adParamOutput, 4000)
'    cmd.Parameters.Append cmd.CreateParameter("P_Ma_UD", adVarChar, adParamInput, 4000)
'    cmd.Parameters("P_Ma_UD").Value = "HTKK"
'    cmd.Execute
    
    strDBUserName = "TuanAnh" 'cmd.Parameters("P_USER_NAME").Value
    strDBPassword = "123456" ' cmd.Parameters("P_PASSWORD").Value

'    Set cmd = Nothing
'    ' Destroy connect to BMT
'    clsDAO.Disconnect
'
'    'connect to database QLT
'    clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'    clsDAO.Connect
    
    Exit Sub
ErrorHandle:

    SaveErrorLog Me.Name, "GetDataInfor", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set clsDAO = Nothing
    Set frmLogin = Nothing
End Sub

Private Sub txtUsername_Change()
    txtUsername.Text = UCase(txtUsername.Text)
End Sub

Private Sub txtUsername_LostFocus()
    If Len(txtUsername.Text) > 0 Then
        txtUsername.Text = UCase(txtUsername.Text)
    End If
End Sub

Private Sub txtUsername_Validate(Cancel As Boolean)
    If txtUsername.Text = vbNullString Then
        DisplayMessage "0056", msOKOnly, miInformation
        Cancel = True
        Exit Sub
    End If
End Sub

Private Function CheckVersion() As Boolean
    Dim rsObj  As ADODB.Recordset
    Dim strSQL As String
    Dim fso    As New FileSystemObject
    Dim cnnStr As String
   
    On Error GoTo ErrHandle
    cnnStr = spathVat & "\NTK_TG\tepmau\"

    If Not fso.FileExists(cnnStr & "cg_ref_codes.dbf") Then
        DisplayMessage "0161", msOKOnly, miCriticalError
        CheckVersion = False
        
        Exit Function
    End If
    
    strSQL = "SELECT rv_low_v From cg_ref_codes WHERE rv_domain = 'HTKK_ABOUT.VERSION'"

    'connect to database BMT
    If clsDAO.Connected = True Then
        clsDAO.Disconnect
    End If

    clsDAO.CreateConnectionString cnnStr
    clsDAO.Connect

    If clsDAO.Connected Then
        Set rsObj = clsDAO.Execute(strSQL)

        If Not rsObj Is Nothing Then
            If CInt(Replace(rsObj.Fields(0).Value, ".", "")) = CInt(Replace(APP_VERSION, ".", "")) Then
                CheckVersion = True
                Exit Function
            Else

                If CInt(Replace(rsObj.Fields(0).Value, ".", "")) > CInt(Replace(APP_VERSION, ".", "")) Then
                    'Versions is differed
                    DisplayMessage "0076", msOKOnly, miCriticalError
                    CheckVersion = False

                    Exit Function
                ElseIf CInt(Replace(rsObj.Fields(0).Value, ".", "")) < CInt(Replace(APP_VERSION, ".", "")) Then
                    DisplayMessage "0075", msOKOnly, miCriticalError
                    CheckVersion = False
                    Exit Function
                Else
                
                    DisplayMessage "0161", msOKOnly, miCriticalError
                    CheckVersion = False
            
                    Exit Function
                End If

            End If

        Else
            DisplayMessage "0161", msOKOnly, miCriticalError
            CheckVersion = False
            
            Exit Function
        End If

    Else
        DisplayMessage "0063", msOKOnly, miCriticalError
        CheckVersion = False

        Exit Function
    End If

    CheckVersion = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "CheckVersion", Err.Number, Err.Description
End Function
