VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessageBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************
' Company           : CMC Software
' Center Name       : FIS (Financial Insurance Solution)
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Class Modules
' Form, Module
'   or Class name   : clsMessageBox
' Descriptions      : Class to process message box
' Start date        : 10/08/2007 (dd/mm/yyyy)
' Finish date       :
' Coder             : hlnam
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************

Option Explicit

Public Enum MsgBoxIcon
    miInformation = 0
    miQuestion = 1
    miWarning = 2
    miCriticalError = 3
End Enum

Public Enum MsgBoxStyle
    msAbortRetryIgnore = 0
    msOKCancel = 1
    msOKOnly = 2
    msOKCancelRetry = 3
    msYesNo = 4
    msYesNoCancel = 5
End Enum

Public Enum MsgBoxResult
    mrAbort = 0
    mrCancel = 1
    mrIgnore = 2
    mrNo = 3
    mrOK = 4
    mrRetry = 5
    mrYes = 6
End Enum

'*******************************************************
'Description    : DisplayMessage function dispay un message dialog
'Author         : hlnam
'Date           : 10/08/2007
'Paramter       :
'   pMsgID    : id of message in the Message.xml
'   pMsgStyle : style of dialog
'   pIcon     : icon of dialog
'   pTitle    : title of dialog
'Return         :True if it is a number, false if it is not a number
'*******************************************************
Public Function DisplayMessage(pMsgID As String, Optional pMsgStyle As MsgBoxStyle, Optional pIcon As MsgBoxIcon, Optional pTitle As String = "Message", Optional pDefaultButton As MsgBoxResult) As MsgBoxResult
    Dim frmMessage As New frmDisplayMessage
    DisplayMessage = frmMessage.DisplayMessage(pMsgID, pMsgStyle, pIcon, pTitle, pDefaultButton)
End Function

'*******************************************************
'Description    : SetControlName function set the caption from
'   init.xml file to the control on frmForm.
'Author         : hlnam
'Date           : 10/08/2007
'Input          : frmForm (Form has controls which are set caption)
'OutPut         :
'Return         : True if it is successful, false if it has any err.
'*******************************************************
Public Sub SetControlCaption(frmForm As Object, Optional formName As String = vbNullString)
    On Error GoTo ErrHandle
    Dim lControl As Control
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeButton As MSXML.IXMLDOMNode
    
    If formName = vbNullString Then formName = frmForm.Name
    
    For Each xmlNode In xmlNodeCaption.childNodes
        If UCase(xmlNode.baseName) = UCase(formName) Then
            frmForm.Caption = xmlNode.Attributes.getNamedItem("Caption").nodeValue
            For Each lControl In frmForm.Controls
                For Each xmlNodeButton In xmlNode.childNodes
                    If UCase(lControl.Name) = UCase(xmlNodeButton.Attributes.getNamedItem("Name").nodeValue) Then
                        lControl.Caption = xmlNodeButton.Attributes.getNamedItem("Caption").nodeValue
                        Exit For
                    End If
                Next
            Next
            Exit For
        End If
    Next
    
    Exit Sub

ErrHandle:
End Sub

'****************************************************
'Description    : SetFormCaption procedure set caption for form
'Author         : hlnam
'Modify by      :
'Date           : 10/08/2007
'Input          : frmForm: (Form has controls which are set caption)
'  bkGround   : Image control
'  lblCaption : Lable control FM20.DLL
'Output:
'Return:

'****************************************************
Public Sub SetFormCaption(frmForm As Object, bkGround As Object, lblCaption As Object)

    'set style for background of caption
    bkGround.Picture = LoadPicture(GetAbsolutePath("..\Pictures\caption.bmp"))
    bkGround.Move 0, 0, frmForm.Width, 320
    bkGround.Stretch = True
    
    'set style for lable of caption
    lblCaption.Top = 30
    lblCaption.Left = 30
    lblCaption.Width = bkGround.Width
    lblCaption.Height = bkGround.Height
    lblCaption.BackStyle = 0
    lblCaption.FontSize = 9
    lblCaption.FontBold = True
    lblCaption.ForeColor = vbWhite

End Sub

'****************************************************
'Description:GetAbsolutePath function get the absolute path of specified path.
'Author     : hlnam
'Date       : 10/08/2007
'Return     : Absolute path.
'****************************************************
Public Function GetAbsolutePath(Optional ByVal strPath As String) As String
    Dim intLoc1 As Integer, intLoc2 As Integer, intCtrl As Integer
    
    On Error GoTo ErrHandle
    
    If InStr(1, strPath, ":\") <> 0 Then
        GetAbsolutePath = strPath
        Exit Function
    End If
    
    GetAbsolutePath = Mid$(App.Path, 1, InStrRev(App.Path, "\")) & "Project"
    
    'If strPath is null string, return project path
    If strPath = vbNullString Then
        Exit Function
    End If
    
    Do
        intLoc1 = InStr(1, strPath, "..\")
        If intLoc1 <> 0 Then
            strPath = Mid$(strPath, intLoc1 + 3)
            intLoc2 = InStrRev(GetAbsolutePath, "\")
            If intLoc2 <> 0 Then _
            GetAbsolutePath = Mid$(GetAbsolutePath, 1, intLoc2 - 1)
        End If
    Loop Until intLoc1 = 0
        
    If Mid$(strPath, 1, 1) = "\" Or strPath = "" Then
        'param contain "\"
        GetAbsolutePath = GetAbsolutePath & strPath
    Else
        'param not contain "\"
        GetAbsolutePath = GetAbsolutePath & "\" & strPath
    End If
    
    Exit Function
    
ErrHandle:
    Dim lFileNum As Long
    
    lFileNum = FreeFile
    Open App.Path & "\Log.log" For Output As #lFileNum
    Print #lFileNum, "GetAbsolutePath: " & Err.Number & " : " & Err.Description
    Print #lFileNum, "App path: " & App.Path
    Close #lFileNum
    
End Function

