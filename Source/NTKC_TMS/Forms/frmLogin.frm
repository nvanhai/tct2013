VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   3885
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
         TabIndex        =   2
         Top             =   300
         Width           =   2175
         VariousPropertyBits=   746604571
         Size            =   "3836;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblUsername 
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   360
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
      Height          =   375
      Left            =   2550
      TabIndex        =   6
      Top             =   1650
      Width           =   1305
      Caption         =   "Exit"
      Size            =   "2302;661"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1140
      TabIndex        =   5
      Top             =   1650
      Width           =   1305
      Caption         =   "Login"
      Size            =   "2302;661"
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
' Project           : Du an ho tro ke khai thue
' Package           : Interface
' Form, Module
'   or Class name   : frmTreeviewMenu
' Descriptions      : Report sh
' Start date        :
' Finish date       :
' Coder             :
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
    Unload frmSystem
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrorHandle
   
    If Len(txtUsername.Text) = 0 Then
        DisplayMessage "0056", msOKOnly, miInformation
        txtUsername.SetFocus
        Exit Sub
    End If
    
    Dim IsValid As Integer
    IsValid = IsValidUserESB()
    Select Case IsValid
        Case 1
            If DisplayMessage("0130", msYesNo, miQuestion) = mrYes Then
                txtPassword.SetFocus
                Exit Sub
            Else
                Unload Me
                Unload frmSystem
                Exit Sub
            End If
        Case 0
            If DisplayMessage("0040", msYesNo, miQuestion) = mrYes Then
                txtPassword.SetFocus
                Exit Sub
            Else
                Unload Me
                Unload frmSystem
                Exit Sub
            End If
    End Select
    
    'Set user name to system caption
    frmSystem.lblUser.caption = Mid$(frmSystem.lblUser.caption, 1, _
        InStr(1, frmSystem.lblUser.caption, ":") + 1) & _
        strUserName
    '********************************
    ' Date: 27/04/06
    ' Check version of application
    If Not CheckVersion Then
        Unload Me
        Unload frmSystem
        Exit Sub
    End If
    '********************************
    ' set trang thai active cua PIT
    isPITActive = checkActivePIT
    TAX_Utilities_Srv_New.isCheckPIT = isPITActive
    ' end
    
    Unload Me
    frmTreeviewMenu.Show
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "cmdOK_Click", Err.Number, Err.Description
End Sub



Private Sub Form_Activate()
    txtUsername.SetFocus
End Sub

Private Sub Form_Load()
    SetControlCaption Me, "frmLogin"
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub
'****************************************************
'Description:IsValidUserESB function check if user and password are valid
'Author:nshung
'Modify by:
'Date:08/20/2013
'Input:
'Output:
'Return:
'****************************************************
Private Function IsValidUserESB() As Integer
On Error GoTo ErrorHandle

    Dim xmlESBReturn As New MSXML.DOMDocument
    Dim strESBReturn As String
    Dim sStatus As String
    
    IsValidUserESB = 2
    Exit Function
    
    strESBReturn = getInfoUserFromESB
    
    'Chuan hoa file xml ket qua - lay duoc tu ESB
    strESBReturn = ChangeTagASSCII(strESBReturn, False)
    
    xmlESBReturn.loadXML strESBReturn

    If (Not xmlESBReturn Is Nothing) Then
        sStatus = xmlESBReturn.getElementsByTagName("Status")(0).Text
        strCurrentVersion = xmlESBReturn.getElementsByTagName("NTKversion")(0).Text
        strUserName = xmlESBReturn.getElementsByTagName("UserName")(0).Text
        Select Case sStatus
            Case "01"  ' Thanh cong
                IsValidUserESB = 2  'Message login thanh cong
                Exit Function
            Case "02" 'Loi xac thuc NSD
                IsValidUserESB = 0   'Message loi use, pass
                Exit Function
            Case "03" 'Cac loi khac cua he thong TMS
                IsValidUserESB = 1  'Message loi khong dang nhap duoc
                Exit Function
            Case Else
                IsValidUserESB = 1
                Exit Function
        End Select
    Else
        IsValidUserESB = 1
    End If
    
ErrorHandle:
    Me.MousePointer = vbDefault
    frmSystem.MousePointer = vbDefault
    SaveErrorLog Me.Name, "IsValidUserESB", Err.Number, Err.Description
End Function
Private Function ChangeTagASSCII(ByVal strTemp As String, ByVal IsTagToASSCII As Boolean) As String
    If (strTemp <> "") Then
        If IsTagToASSCII Then
            strTemp = Strings.Replace$(strTemp, "<", "&lt;", 1, Len(strTemp), vbTextCompare)
            strTemp = Strings.Replace$(strTemp, ">", "&gt;", 1, Len(strTemp), vbTextCompare)
        Else
            strTemp = Strings.Replace$(strTemp, "&lt;", "<", 1, Len(strTemp), vbTextCompare)
            strTemp = Strings.Replace$(strTemp, "&gt;", ">", 1, Len(strTemp), vbTextCompare)
        End If
        ChangeTagASSCII = strTemp
    End If
End Function
Private Function getInfoUserFromESB() As String
On Error GoTo ErrorHandle
    'Load file template xml --> param gui cho ESB
    Dim paXmlDoc As New MSXML.DOMDocument
    paXmlDoc.Load GetAbsolutePath("..\InterfaceTemplates\xml\paramNsdInESB.xml")
    
    'Get value config
    Dim cfigXml As New MSXML.DOMDocument
    'cfigXml.Load GetAbsolutePath("..\Project\ConfigWithESB.xml")
    Set cfigXml = LoadConfig()


    Dim paNode As MSXML.IXMLDOMNode
    Dim cfigNode As MSXML.IXMLDOMNode
    Dim CloneNode As MSXML.IXMLDOMNode
    Dim paNodeChild As MSXML.IXMLDOMNode
    Dim sTranCode As String
    Dim sTaxOffice As String
    Dim sUrlWs As String
    Dim soapAct As String
    Dim fldName As String
    Dim fldValue As String
    Dim xmlRequest As String
    
    sUrlWs = cfigXml.getElementsByTagName("WsUrl")(0).Text
    soapAct = cfigXml.getElementsByTagName("SoapAction")(0).Text
    xmlRequest = cfigXml.getElementsByTagName("XmlRequest")(0).firstChild.xml & cfigXml.getElementsByTagName("XmlRequest")(0).lastChild.xml
    sTranCode = cfigXml.getElementsByTagName("TRAN_CODE")(0).Text
    sTaxOffice = cfigXml.getElementsByTagName("TaxOffcice")(0).Text
    fldName = cfigXml.getElementsByTagName("ParamName")(0).Text
    
    'Set value config to template xml param
    paXmlDoc.getElementsByTagName("TRAN_CODE")(0).Text = sTranCode
    paXmlDoc.getElementsByTagName("UserName")(0).Text = txtUsername.Text
    paXmlDoc.getElementsByTagName("TaxOffcice")(0).Text = sTaxOffice
    paXmlDoc.getElementsByTagName("Pass")(0).Text = txtPassword.Text
    
    fldValue = paXmlDoc.xml
    fldValue = ChangeTagASSCII(fldValue, True)
    
    'Return value from ESB
    getInfoUserFromESB = DataFromESB(sUrlWs, soapAct, xmlRequest, fldName, fldValue)
    
ErrorHandle:
        SaveErrorLog Me.Name, "IsValidUserESB", Err.Number, Err.Description
End Function

Private Function DataFromESB(sWebUrl As String, sSoapAct As String, sXmlSoap As String, sParam As String, sValue As String) As String
    Dim oWsXML As New XMLRequestNuic '' initialize a new Instance of XMLRequestNuic Class
    'Dim aDatos() As String           '' Variable for store the parameters that we need to pass to de Web service
    Dim iTotalElem As Integer        '' This is only for know how many filters o parameters we are passing to the web service
    Dim bFlag As Boolean             '' When the value is 0 (zero,false) the XML Structure is not correct, but if the value is 1 (One,True) then the structure is correct.
    Dim iCant As Integer             '' is a counter for replace the values into the name of parameters
    iCant = 1
    bFlag = 0
'    aDatos = Split(sValue, ",")
'    If Not IsArray(aDatos) Then
'        aDatos = Split(sValue, "-")
'        If Not IsArray(aDatos) Then
'            aDatos = Split(sValue, ".")
'            If Not IsArray(aDatos) Then
'                aDatos = Split(sValue, "+")
'                If Not IsArray(aDatos) Then
'                    SaveErrorLog Me.Name, "frmLogin", Err.Number, Err.Description
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
    'iTotalElem = UBound(aDatos)      '' We Store the MAX index to the iTotalElem variable
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' WE VALIDATING, IF THE XML STRUCTURE IS CORRECT TO MADE THE PETITION
    ''   bFlag=0 IS WRONG
    ''   bFlag=1 IT IS OK
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If InStr(sXmlSoap, "<?xml") > 0 And InStr(sXmlSoap, "<?xml") <= 6 Then
         bFlag = 1
        If InStr(sXmlSoap, "<soap:Envelope") > 0 Then
            bFlag = 1
            If InStr(sXmlSoap, "<soap:Body>") > 0 Then
                bFlag = 1
            Else
                 bFlag = 0
            End If
        Else
             bFlag = 0
        End If
    Else
        bFlag = 0
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Starting to replace the input parameters
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If bFlag Then
         
            Dim iInicio As Integer
            Dim iFinalParte1 As Integer
            Dim iInicioParte2 As Integer
            Dim iFinal As Integer
            Dim LongURL As Integer
            Dim oFuncion() As String
            Dim sFuncionNombre As String
            Dim sBuscar As String
            Dim sInputParam As String
            Dim tmpUrlSoap As String
            Dim iCont As Integer
            Dim tmpXmlSoap As String
            Dim tmpParte1 As String
            Dim tmpParte2 As String
            Dim oParametro As Variant
            
            ''WE STORED THE ORIGINAL XML STRUCTURE IN A TEMPORARY VARIABLE
            tmpXmlSoap = sXmlSoap
            iCont = 1
            Dim i As Integer
            For i = 1 To Len(sXmlSoap)
                If InStr(tmpXmlSoap, "string") Then
                    ''SET the first coincidence with the "string" Word
                    iFinalParte1 = InStr(sXmlSoap, "string")
                     ''SET the end of the first coincidence with the "string" Word
                    iInicioParte2 = InStr(sXmlSoap, "string") + 6
                    tmpParte1 = Mid(tmpXmlSoap, 1, iFinalParte1 - 1)
                    sXmlSoap = tmpParte1
                    tmpParte2 = Mid(tmpXmlSoap, iInicioParte2, Len(tmpXmlSoap))
                    sXmlSoap = tmpParte2
                    tmpXmlSoap = tmpParte1 & "@Parametro" & iCont & tmpParte2
                    sXmlSoap = tmpXmlSoap
                    i = i + 6
                    iCont = iCont + 1
                End If

            Next
            ''Asignamos el resultado al txtXmlSoap.text
            ''WE SET THE RESULT OF THE "FOR" TO THE txtXmlSoap.text CONTROL
            sXmlSoap = tmpXmlSoap
       
'        ''Replacing the "@Parametro1" with the value in the first position of the txtCriterios.text CONTROL.
'        For Each oParametro In aDatos
'            Dim Var As String
'            If InStr(sXmlSoap, "@Parametro" & iCant) > 0 Then
'                sXmlSoap = Replace(sXmlSoap, "@Parametro" & iCant, oParametro)
'            End If
'            iCant = iCant + 1
'        Next

        sXmlSoap = Replace(sXmlSoap, "@Parametro1", sValue)
        ''validating if all is ok
        If sWebUrl = "" Or sSoapAct = "" Or sXmlSoap = "" Then
            SaveErrorLog Me.Name, "frmLogin", Err.Number, Err.Description & "Kiem tra Url webservice,soap action..."
            Exit Function
        Else
            DataFromESB = oWsXML.PostWebservice(sWebUrl, sSoapAct, sXmlSoap)
        End If
    Else
         'DataFromESB = "the XML Structure is not Correct. please verify your XML structura data."
         SaveErrorLog Me.Name, "frmLogin", Err.Number, Err.Description & "the XML Structure is not Correct. please verify your XML structura data."
         Exit Function
    End If
End Function

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
    
    
'    connect to database BMT
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "BMT", "LOGIN_USER", "LOGIN_USER"
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
'        clsDAO.Connect
'        Me.MousePointer = vbDefault
'        frmSystem.MousePointer = vbDefault
'    End If
'
'    'set key trong BMT, call prc_get_key
'    cmd.ActiveConnection = clsDAO.Connection
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "{call BMT_PCK_BMHT.prc_get_key()}"
'    cmd.Execute
'    Set cmd = Nothing
'
'    'create slq query check username and password
'    userid = clsConvert.Convert(txtUsername.Text, UNICODE, TCVN)
'    password = clsConvert.Convert(txtPassword.Text, UNICODE, TCVN)
'    strSQL = "SELECT nvl(BMT_PCK_BMHT.fnc_check_login('" & _
'                UCase(userid) & "','" & password & "'),-1)  result FROM dual"
'
'    'check username and password
'    Set rec = clsDAO.Execute(strSQL)
'
'    If rec.Fields(0).Value = 0 Then
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
'        '*******************************
'        'Modify date: 12/12/2005
'        If rec.Fields.Count > 0 Then
'            IsValidUser = 2
'            'get User ID
'            strUserID = rec.Fields(0).Value
'            'get User name
'            strUserName = clsConvert.Convert(rec.Fields(1).Value, TCVN, UNICODE)
'            '***********************************
'            ' get cqt id (Chi cuc thue)
'            strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_CQT'"
'            Set rec = clsDAO.Execute(strSQL)
'            ' neu chi cuc thue khong duoc dang ky su dung QLT va dung NTKC thi lay den Cuc thue
'            If rec Is Nothing Then
'                'get cqt id (Cuc thue)
'                strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_TINH'"
'                Set rec = clsDAO.Execute(strSQL)
'            End If
'            'get cqt id
'            strTaxOfficeId = clsConvert.Convert(rec.Fields(0).Value, TCVN, UNICODE)
'            If Len(Trim(strTaxOfficeId)) = 3 Then
'                ' ghep them 2 so 0 vao dang sau la lay duoc ma cuc thue
'                strTaxOfficeId = strTaxOfficeId & "00"
'            End If
'
'        Else
'            IsValidUser = 1
'        End If
'        '*******************************
'    ElseIf rec.Fields(0).Value = -1 Then
'        IsValidUser = 0
'    Else
'        strSQL = "SELECT ten_nsd, mo_ta FROM bmt_nsd WHERE ten_nsd='" & userid & "' " & _
'        "AND MA_NSD IN (SELECT MA_NSD FROM bmt_nsd_nhom " & _
'        "WHERE MA_NHOM IN (SELECT MA_NHOM FROM BMT_NHOM_CHUC_NANG " & _
'        "WHERE MA_CHUC_NANG IN (SELECT MA_CHUC_NANG FROM bmt_chuc_nang " & _
'        "WHERE ma_ud = 'HTKK')))"
'        Set rec = clsDAO.Execute(strSQL)
'
'        If rec.Fields.Count > 0 Then
'            IsValidUser = 2
'            'get User ID
'            strUserID = rec.Fields(0).Value
'            'get User name
'            strUserName = clsConvert.Convert(rec.Fields(1).Value, TCVN, UNICODE)
'            ' get cqt id (Chi cuc thue)
'            strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_CQT'"
'            Set rec = clsDAO.Execute(strSQL)
'            ' neu chi cuc thue khong duoc dang ky su dung QLT va dung NTKC thi lay den Cuc thue
'            If rec Is Nothing Then
'                'get cqt id (Cuc thue)
'                strSQL = "SELECT gia_tri  FROM bmt_tham_so WHERE ten='MA_TINH'"
'                Set rec = clsDAO.Execute(strSQL)
'            End If
'            'get cqt id
'            strTaxOfficeId = clsConvert.Convert(rec.Fields(0).Value, TCVN, UNICODE)
'            If Len(Trim(strTaxOfficeId)) = 3 Then
'                ' ghep them 2 so 0 vao dang sau la lay duoc ma cuc thue
'                strTaxOfficeId = strTaxOfficeId & "00"
'            End If
'        End If
'
'        IsValidUser = 2
'
'    End If
'    rec.Close
'    Set rec = Nothing
    IsValidUser = 2
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
    
    'set key trong BMT, call prc_get_key
'    cmd.ActiveConnection = clsDAO.Connection
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BMT_PCK_BMHT.prc_get_key"
   
    
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
    
'    strDBUserName = cmd.Parameters("P_USER_NAME").Value
'    strDBPassword = cmd.Parameters("P_PASSWORD").Value
'
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
    Dim rsObj As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    CheckVersion = True
    Exit Function
    
'    strSQL = "SELECT rv_low_value phien_ban " & _
'           "From cg_ref_codes " & _
'           "WHERE (rv_domain = 'HTKK_ABOUT.VERSION')"
'    'connect to database BMT
'    If clsDAO.Connected Then
'        Set rsObj = clsDAO.Execute(strSQL)
'        If rsObj.Fields(0).Value = "" Then
'            'Can not found table or not exist value
'            DisplayMessage "0075", msOKOnly, miCriticalError
'            Exit Function
'        ElseIf CInt(Replace(rsObj.Fields(0).Value, ".", "")) > _
'               CInt(Replace(APP_VERSION, ".", "")) Then
'            'Versions is differed
'            DisplayMessage "0076", msOKOnly, miCriticalError
'            Exit Function
'        ElseIf CInt(Replace(rsObj.Fields(0).Value, ".", "")) < _
'               CInt(Replace(APP_VERSION, ".", "")) Then
'            DisplayMessage "0075", msOKOnly, miCriticalError
'            Exit Function
'        End If
'    Else
'        DisplayMessage "0063", msOKOnly, miCriticalError
'        Exit Function
'    End If

    ' Check version cua ung dung voi phien ban cua service tra ve
        
       If strCurrentVersion = "" Then
            'Can not found table or not exist value
            DisplayMessage "0075", msOKOnly, miCriticalError
            Exit Function
        ElseIf CInt(Replace(strCurrentVersion, ".", "")) > _
               CInt(Replace(APP_VERSION, ".", "")) Then
            'Versions is differed
            DisplayMessage "0076", msOKOnly, miCriticalError
            Exit Function
        ElseIf CInt(Replace(strCurrentVersion, ".", "")) < _
               CInt(Replace(APP_VERSION, ".", "")) Then
            DisplayMessage "0075", msOKOnly, miCriticalError
            Exit Function
        End If

    
    CheckVersion = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "CheckVersion", Err.Number, Err.Description
End Function

' Kiem tra activ PIT
Private Function checkActivePIT() As Boolean
    Dim rsObj As ADODB.Recordset
    Dim strSQL As String
    Dim resultPIT As Boolean
    On Error GoTo ErrHandle
    resultPIT = False
    strSQL = "SELECT rv_low_value " & _
           "From cg_ref_codes " & _
           "WHERE (rv_domain = 'NTK.PIT_ACTIVE')"
    'connect to database QLT
    If clsDAO.Connected Then
        Set rsObj = clsDAO.Execute(strSQL)
        If Not rsObj Is Nothing Then
            If rsObj.Fields.Count > 0 Then
                If rsObj.Fields(0).Value = "1" Then
                    resultPIT = True
                Else
                    resultPIT = False
                End If
            End If
        End If
    End If
    checkActivePIT = resultPIT
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "checkActivePIT", Err.Number, Err.Description
End Function
