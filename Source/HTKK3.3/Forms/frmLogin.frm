VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3045
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1799.087
   ScaleMode       =   0  'User
   ScaleWidth      =   5182.98
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Quay l¹i"
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
      Left            =   4140
      TabIndex        =   7
      Top             =   2550
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Tho¸t"
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
      HelpContextID   =   8126
      Left            =   4140
      TabIndex        =   4
      Top             =   2550
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "§å&ng ý"
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
      Left            =   2790
      TabIndex        =   3
      Top             =   2550
      Width           =   1305
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&M· sè míi"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   2550
      Width           =   1305
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Xãa"
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
      TabIndex        =   1
      Top             =   2550
      Width           =   1305
   End
   Begin VB.ComboBox cboTaxIdString 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmLogin.frx":0000
      Left            =   1890
      List            =   "frmLogin.frx":0002
      TabIndex        =   0
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Lbname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tªn NNT"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2370
      TabIndex        =   8
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "§¨ng nhËp hÖ thèng"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   6
      Top             =   0
      Width           =   1845
   End
   Begin VB.Label lblTaxIdString 
      Caption         =   "M· sè thuÕ"
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
      Left            =   600
      TabIndex        =   5
      Top             =   2100
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -28.168
      X2              =   5185.797
      Y1              =   921.7
      Y2              =   921.7
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   0
      Picture         =   "frmLogin.frx":0004
      Stretch         =   -1  'True
      Top             =   330
      Width           =   5535
   End
   Begin VB.Image imgCaption 
      Height          =   345
      Left            =   30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmLogin"
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

Option Explicit

Private Sub cboTaxIdString_Change()
    Static strOldValue As String
    On Error GoTo ErrorHandle
    
    If Len(cboTaxIdString.Text) <> 0 And Not IsNumeric(cboTaxIdString.Text) Then
        cboTaxIdString.Text = strOldValue
    Else
        strOldValue = cboTaxIdString.Text
    End If
        
    If cboTaxIdString.Text = strTaxIdString Or cboTaxIdString.Text = "" Then
        cmdDelete.Enabled = False
        Lbname.caption = ""
    Else
        cmdDelete.Enabled = True
    End If
        
    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "cboTaxIdString_Change", Err.Number, Err.Description
End Sub

Private Sub cboTaxIdString_Click()
'dhdang them chuc nang hien thi ten DN tren Form Login
    Dim fso As New FileSystemObject
    Dim xmlDom As New MSXML.DOMDocument, xmlDomHeader As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode, xmlNodeList As MSXML.IXMLDOMNodeList
    Dim intCtrl As Integer
    Dim strTenDoanhNghiep As String
    Dim clsConverter As New clsUnicodeTCVNConverter
        
    strTenDoanhNghiep = ""
    
    If fso.FolderExists(GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text)) Then
       ' Load data header to DOM
        xmlDom.Load (GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text)) & "\Header_01.xml"
        ' Get Cell nodes
        Set xmlNodeList = xmlDom.getElementsByTagName("Cell")
        Set xmlNode = xmlNodeList(13)
        strTenDoanhNghiep = GetAttribute(xmlNode, "Value")
        
        ' Neu la file Header cu, ko co CQT cap Cuc va Chi cuc thue quan ly thi xmlNodeList = 25
        If (xmlNodeList.length = 25) Then
             DisplayMessage "0139", msOKOnly, miInformation
             'blnFirstUse = True
             prepareHeaderInfo xmlNodeList
        Else
            prepareHeaderInfo xmlNodeList
        End If
    End If
    
    If cboTaxIdString.Text = strTaxIdString Or cboTaxIdString.Text = "" Then
        cmdDelete.Enabled = False
        Lbname.caption = clsConverter.Convert(strTenDoanhNghiep, UNICODE, TCVN)
    Else
        cmdDelete.Enabled = True
        Lbname.caption = clsConverter.Convert(strTenDoanhNghiep, UNICODE, TCVN)
    End If
End Sub

Private Sub cboTaxIdString_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    Dim sNumber As String

    sNumber = "0123456789"
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        cmdOK.SetFocus
        Exit Sub
    End If
    
    If Len(cboTaxIdString.Text) = 13 Then KeyAscii = 0
    
    If InStr(1, sNumber, Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If

    Exit Sub

ErrorHandle:
    SaveErrorLog Me.Name, "cboTaxIdString_KeyPress", Err.Number, Err.Description
End Sub

Private Sub cmdBack_Click()
    Unload Me
    frmTreeviewMenu.Show
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandle
    
    Dim lIndex As Long, lLocate As Long
    Dim strTaxID As String
    
    lIndex = cboTaxIdString.ListIndex
    strTaxID = cboTaxIdString.Text
    If Trim(strTaxID) = "" Then 'TaxId is not exist
        DisplayMessage "0051", msOKOnly, miInformation
        Exit Sub
    End If
    If Not ExistItem(strTaxID, lLocate) Then ' TaxId is not exist
        DisplayMessage "0048", msOKOnly, miInformation
        Exit Sub
    End If
    
    If DisplayMessage("0049", msYesNo, miQuestion, , mrNo) = mrYes Then
        cboTaxIdString.RemoveItem lLocate
        If lIndex > cboTaxIdString.listcount - 1 Then
            lIndex = cboTaxIdString.listcount - 1
        End If
        cboTaxIdString.ListIndex = lIndex
        DeleteFolder GetAbsolutePath("..\DataFiles\" & strTaxID)
    End If
    '*******************************
    If cboTaxIdString.Text = "" Then
        cmdDelete.Enabled = False
    End If
    '*******************************
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdDelete_Click", Err.Number, Err.Description
End Sub

Private Sub cmdNew_Click()
    cboTaxIdString.SetFocus
    cboTaxIdString.Text = ""
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandle
    Dim blnFirstUse As Boolean, fle As file
    Dim fso As New FileSystemObject
    Dim xmlDom As New MSXML.DOMDocument, xmlDomHeader As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode, xmlNodeList As MSXML.IXMLDOMNodeList
    Dim intCtrl As Integer
    Dim strTenDoanhNghiep As String
    Dim clsConverter As New clsUnicodeTCVNConverter
        
    strTenDoanhNghiep = ""
    'Check null tax id
    If Trim(cboTaxIdString.Text) = "" Then
        DisplayMessage "0051", msOKOnly, miInformation
        cboTaxIdString.SetFocus
        Exit Sub
    End If
    'Check validity of tax id
    If Not IsValidTaxId(cboTaxIdString.Text) Then
        DisplayMessage "0047", msOKOnly, miInformation
        cboTaxIdString.SetFocus
        Exit Sub
    End If
    '**********************************
    If fso.FolderExists(GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text)) Then
       ' Load data header to DOM
        xmlDom.Load (GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text)) & "\Header_01.xml"
        ' Get Cell nodes
        Set xmlNodeList = xmlDom.getElementsByTagName("Cell")
        Set xmlNode = xmlNodeList(13)
        strTenDoanhNghiep = GetAttribute(xmlNode, "Value")
        
        ' Neu la file Header cu, ko co CQT cap Cuc va Chi cuc thue quan ly thi xmlNodeList = 25
        If (xmlNodeList.length = 25) Then
             DisplayMessage "0139", msOKOnly, miInformation
             blnFirstUse = True
             prepareHeaderInfo xmlNodeList
        Else
            prepareHeaderInfo xmlNodeList
        End If
    End If
    '**********************************
    'Set tax id to system caption
    frmSystem.lblUserInfo.caption = Mid$(cboTaxIdString.Text, 1, 10) & _
        IIf(Len(cboTaxIdString.Text) = 13, " - " & Mid$(cboTaxIdString.Text, 11, 3), "") & " : " & clsConverter.Convert(strTenDoanhNghiep, UNICODE, TCVN)
        
        
    'Create new folder if it's not exist
    If Not ExistItem(cboTaxIdString.Text) Then
        blnFirstUse = True
        fso.CreateFolder GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text)
'        For Each fle In fso.GetFolder(GetAbsolutePath("..\InterfaceTemplates")).Files
'            If Right$(fle.Name, 4) = ".dtd" Then
'                fso.CopyFile fle.path, GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text & "\" & fle.Name)
'            End If
'        Next
        '********************************
        'Create Session file
        'fso.CreateTextFile GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text & "\" & "Session.dat")
        '********************************
        SetHeaderInfo cboTaxIdString.Text
        cboTaxIdString.AddItem cboTaxIdString.Text
    Else
        'If Header file does not exist, create it.
        If Not fso.FileExists(GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text & "\Header_01.xml")) Then
            SetHeaderInfo cboTaxIdString.Text
            xmlDom.Load (GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text)) & "\Header_01.xml"
        End If
        
        'Check Version of Header
        xmlDomHeader.Load GetAbsolutePath("..\InterfaceTemplates\xml\Header_01.xml")
        
        If xmlDom.getElementsByTagName("Sections")(0).Attributes.getNamedItem("Version") Is Nothing And _
            Not xmlDomHeader.getElementsByTagName("Sections")(0).Attributes.getNamedItem("Version") Is Nothing Then
                UpdateHeader xmlDom, xmlDomHeader
        ElseIf CDbl(GetAttribute(xmlDomHeader.getElementsByTagName("Sections")(0), "Version")) > _
            CDbl(GetAttribute(xmlDom.getElementsByTagName("Sections")(0), "Version")) Then
                UpdateHeader xmlDom, xmlDomHeader
        End If
        
        Set xmlDomHeader = Nothing
    End If
    '**********************************
    'Longvh added
    'Date:
    '
    Set xmlDom = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    '**********************************
    strTaxIdString = cboTaxIdString.Text
    TAX_Utilities_v2.DataFolder = GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text & "\")
    
    Set fso = Nothing
    Unload Me
    
    If blnFirstUse Then
        SetFirstUse
    Else
        frmTreeviewMenu.Show
    End If
    Exit Sub
ErrHandle:
    Set fso = Nothing
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
End Sub

Private Sub cmdRestore_Click()
    'frmTreeviewMenu.RestoreData
End Sub

Private Sub Form_Load()
    'SetControlCaption Me, "frmLogin"
    
    Me.Top = (frmSystem.Height - Me.Height - 1000) / 2
    Me.Left = (frmSystem.Width - Me.Width) / 2
    
    If strTaxIdString = "" Then
        cmdClose.Visible = True
        cmdBack.Visible = False
        cmdDelete.Enabled = True
    Else
        cmdClose.Visible = False
        cmdBack.Visible = True
        cmdDelete.Enabled = False
    End If
    'Reset tax id to system caption
    'frmSystem.lblUserInfo.caption = Mid$(frmSystem.lblUserInfo.caption, 1, _
        InStr(1, frmSystem.lblUserInfo.caption, ":") + 1)
            
    
    SetValueToList
    
    If cboTaxIdString.ListIndex = -1 Then
        cmdDelete.Enabled = False
    End If
End Sub

'******************************
'Description: SetValueToList procedure list subfolders in
'             the datafiles folder and add names to list.
'******************************
Private Sub SetValueToList()
    On Error GoTo ErrHandle
    
    Dim fso As New FileSystemObject
    Dim fldList As Folders
    Dim fldSubFolder As Folder
    
    'Get subfolders in DataFiles folder
    Set fldList = fso.GetFolder(GetAbsolutePath("..\DataFiles\")).SubFolders
    
    'Add name of subfolders to list
    For Each fldSubFolder In fldList
        If IsValidTaxId(fldSubFolder.Name) Then _
        cboTaxIdString.AddItem fldSubFolder.Name
    Next
    If cboTaxIdString.listcount > 0 Then _
        cboTaxIdString.ListIndex = 0
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetValueToList", Err.Number, Err.Description
End Sub

'******************************
'Description: ExistItem function check whether
'             selected item is availble in list.
'******************************
Private Function ExistItem(ByVal strValue As String, Optional ByRef lIndex As Long) As Boolean
    Dim lCtrl As Long
    
    For lCtrl = 0 To cboTaxIdString.listcount - 1
        If strValue = cboTaxIdString.List(lCtrl) Then
            ExistItem = True
            lIndex = lCtrl
            Exit Function
        End If
    Next lCtrl
End Function

Private Sub Form_Resize()
     SetFormCaption Me, imgCaption, lblCaption
     imgCaption.Height = 200
End Sub

'******************************
'Description: SetHeaderInfo procedure set taxid string to Header
'******************************
Private Sub SetHeaderInfo(ByVal strTaxID As String)
    On Error GoTo ErrHandle
    Dim xmlDom As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode, xmlNodeList As MSXML.IXMLDOMNodeList
    Dim intCtrl As Integer, strDataFileName As String
    
    ' Load data header to DOM
    xmlDom.Load GetAbsolutePath("..\InterfaceTemplates\xml\Header_01.xml")
    
    ' Get Cell nodes
    Set xmlNodeList = xmlDom.getElementsByTagName("Cell")
    
    'Set tax Id string
    For intCtrl = 1 To Len(strTaxID)
        Set xmlNode = xmlNodeList(intCtrl - 1)
        SetAttribute xmlNode, "Value", Mid$(strTaxID, intCtrl, 1)
    Next intCtrl
    
    'Save to specified data folder
    strDataFileName = "..\DataFiles\" & cboTaxIdString.Text & "\Header_01.xml"
    xmlDom.save GetAbsolutePath(strDataFileName)
    
    Set xmlDom = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetHeaderInfo", Err.Number, Err.Description
End Sub

Private Sub prepareHeaderInfo(ByVal headerXMLlNodeList As MSXML.IXMLDOMNodeList)
    On Error GoTo ErrHandle
    Dim xmlDom As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode, xmlNodeList As MSXML.IXMLDOMNodeList
    Dim headerXMLNode As MSXML.IXMLDOMNode
    Dim i As Integer
    Dim strValue As String
    Dim intCtrl As Integer, strDataFileName As String
    Dim clsConverter As New clsUnicodeTCVNConverter
    
    ' Load data header to DOM
    xmlDom.Load GetAbsolutePath("..\InterfaceTemplates\xml\Header_01.xml")
    
    ' Get Cell nodes
    Set xmlNodeList = xmlDom.getElementsByTagName("Cell")
    
    For i = 0 To xmlNodeList.length - 1 ' Header cu co 25 chi tieu (Chua co CQT va Chi cuc thue)
        ' Lay tung Item cua Header cu
        Set headerXMLNode = headerXMLlNodeList(i)
        strValue = Trim(GetAttribute(headerXMLNode, "Value"))
        'If strValue = " " Then strValue = ""
        'strValue = clsConverter.Convert(strValue, UNICODE, TCVN)
        
        ' Set vao header moi
        Set xmlNode = xmlNodeList(i)
        SetAttribute xmlNode, "Value", strValue
        
    Next
        
    'Save to specified data folder
    strDataFileName = "..\DataFiles\" & cboTaxIdString.Text & "\Header_01.xml"
    xmlDom.save GetAbsolutePath(strDataFileName)
    
    Set xmlDom = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set headerXMLNode = Nothing
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetHeaderInfo", Err.Number, Err.Description
End Sub

'******************************
'Description: DeleteFolder procedure delete a specified folder
'******************************
Private Sub DeleteFolder(ByVal strFolderName As String)
    On Error GoTo ErrHandle
    
    Dim fso As New FileSystemObject
    
    If fso.FolderExists(strFolderName) Then
        fso.DeleteFolder strFolderName, True
    End If
    
    Set fso = Nothing
    Exit Sub
    
ErrHandle:
    Set fso = Nothing
    SaveErrorLog Me.Name, "DeleteFolder", Err.Number, Err.Description
End Sub

'******************************
'Description: SetFirstUse procedure call common interface form
'******************************
Private Sub SetFirstUse()
    On Error GoTo ErrHandle
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode, xmlNodeRoot As MSXML.IXMLDOMNode
    Dim strId As String
'    xmlDocument.Load App.path & "\Menu.xml"
'    Set xmlNode = xmlDocument.getElementsByTagName("Root").Item(0)
'
'    If Trim(xmlNode.Attributes.getNamedItem("FirstTimeRunID").nodeValue) <> "" Then
'        frmTreeviewMenu.ProcessMenuAction Trim(xmlNode.Attributes.getNamedItem("FirstTimeRunID").nodeValue)
'    End If
'    Set xmlDocument = Nothing
'    Set xmlNode = Nothing
    
    xmlDocument.Load App.path & "\Menu.xml"
    Set xmlNodeRoot = xmlDocument.getElementsByTagName("Root").Item(0)
    
    'Get default Id loaded in the first use.
    strId = Trim(xmlNodeRoot.Attributes.getNamedItem("FirstTimeRunID").nodeValue)
    If strId = "" Then _
        Exit Sub
        
    For Each xmlNode In xmlNodeRoot.childNodes
        If strId = xmlNode.Attributes.getNamedItem("ID").nodeValue Then
            TAX_Utilities_v2.NodeMenu = xmlNode
            ReDim arrActiveForm(1)
            arrActiveForm(1).id = strId
            arrActiveForm(1).showed = False
            frmInterfaces.Show
            Exit For
        End If
    Next
    
    Set xmlDocument = Nothing
    Set xmlNode = Nothing
    Set xmlNodeRoot = Nothing
    
    Exit Sub
ErrHandle:
    Set xmlDocument = Nothing
    Set xmlNode = Nothing
    Set xmlNodeRoot = Nothing
    SaveErrorLog Me.Name, "SetFirstUse", Err.Number, Err.Description
End Sub

Private Sub UpdateHeader(xmlDomData As MSXML.DOMDocument, xmlDomTemplate As MSXML.DOMDocument)
    On Error GoTo ErrHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim lCtrl As Long, strDataHeaderFileName As String
    Dim fso As New FileSystemObject
    
    For lCtrl = 1 To xmlDomData.getElementsByTagName("Cell").length
        SetAttribute xmlDomTemplate.getElementsByTagName("Cell")(lCtrl - 1), "Value", _
            GetAttribute(xmlDomData.getElementsByTagName("Cell")(lCtrl - 1), "Value")
    Next
    Set xmlDomData = xmlDomTemplate
    
    strDataHeaderFileName = GetAbsolutePath("..\DataFiles\" & cboTaxIdString.Text) & "\Header_01.xml"
    If fso.FileExists(strDataHeaderFileName) Then
        fso.GetFile(strDataHeaderFileName).Attributes = Normal
    End If
    
    xmlDomData.save strDataHeaderFileName
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "UpdateHeader", Err.Number, Err.Description
End Sub

