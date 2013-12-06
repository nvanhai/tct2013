VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmInterfaces 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7905
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   11535
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmInterfaces"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   8880
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   6750
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   11535
      Begin MSCommLib.MSComm MSComm1 
         Left            =   1050
         Top             =   1770
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin FPUSpreadADO.fpSpread fpSpread1 
         Height          =   6600
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   11475
         _Version        =   458752
         _ExtentX        =   20241
         _ExtentY        =   11642
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterfaces.frx":0000
      End
      Begin MSForms.Label lblExit 
         Height          =   945
         Left            =   2760
         TabIndex        =   11
         Top             =   2820
         Visible         =   0   'False
         Width           =   7095
         Size            =   "12515;1667"
         FontName        =   "Tahoma"
         FontHeight      =   255
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblConnecting 
         Height          =   945
         Left            =   3480
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   7095
         Size            =   "12515;1667"
         FontName        =   "Tahoma"
         FontHeight      =   255
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblLoading 
         Height          =   945
         Left            =   3480
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   7095
         Size            =   "12515;1667"
         FontName        =   "Tahoma"
         FontHeight      =   255
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   6990
      Width           =   11535
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   30
         Width           =   1335
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   375
         Left            =   7440
         TabIndex        =   18
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "X„a"
         Size            =   "2293;661"
         Accelerator     =   88
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdViewNow 
         Height          =   375
         Left            =   6090
         TabIndex        =   17
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Xem tÍ khai"
         Size            =   "2302;661"
         Accelerator     =   72
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblBarcode 
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
         VariousPropertyBits=   8388627
         Size            =   "2143;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblWarning 
         Height          =   255
         Left            =   9120
         TabIndex        =   14
         Top             =   150
         Visible         =   0   'False
         Width           =   2325
         ForeColor       =   255
         VariousPropertyBits=   8388627
         Size            =   "4101;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVersion 
         Height          =   255
         Left            =   8610
         TabIndex        =   13
         Top             =   150
         Width           =   405
         VariousPropertyBits=   8388627
         Size            =   "714;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblLabelVersion 
         Height          =   255
         Left            =   4380
         TabIndex        =   12
         Top             =   150
         Width           =   4125
         VariousPropertyBits=   8388627
         Size            =   "7276;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblFile 
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
         VariousPropertyBits=   8388627
         Size            =   "2143;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFilePath 
         Height          =   255
         Left            =   1530
         TabIndex        =   9
         Top             =   150
         Width           =   1785
         VariousPropertyBits=   8388627
         Size            =   "3149;450"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   8790
         TabIndex        =   0
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   25
         Caption         =   "Ghi lπi"
         Size            =   "2293;661"
         Accelerator     =   71
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdExit 
         Height          =   375
         Left            =   10140
         TabIndex        =   1
         Top             =   390
         Width           =   1305
         Caption         =   "Tho∏t"
         Size            =   "2293;661"
         Accelerator     =   84
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSForms.Label lblCaption 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3975
      ForeColor       =   -2147483634
      Size            =   "7011;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgCaption 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mCommPort = 1
Const mBaudRate = 9600
Const mParity = "N"
Const mDataBits = 8
Const mStopBits = 1
Const mHandshaking = 1

Private xmlDocumentInit()       As MSXML.DOMDocument
Private arrStrElements()        As String               ' array of barcode string or file name string
Private mHeaderSheet            As Integer              ' Save value of Header sheet (last sheet)
Private blnReceiveByBarcode     As Boolean                    ' Check whether form is loaded
Private objTaxBusiness          As Object               ' private business object (cls001, cls002, cls003, ...)
Private strTaxReportInfo        As String               ' Info about current tax report

Private mOnLoad                 As Boolean
Private blnOnLoadEvent          As Boolean
'Private strMaSoTep              As String
Private strNgayNhanToKhai       As String
Private strMaPhongXuLy          As String
Private blnSaveSuccess          As Boolean
Private rsPXL                   As ADODB.Recordset      ' Luu danh sach cac phong ban
Private strTaxReportVersion     As String

Private arrBCBuffer() As String
Private arrBCNew() As String
Private verToKhai As Byte       ' Luu cac kieu ma vach cho cac version ke khai khac nhau
Private maxBarCode As Long       ' Su dung trong truong hop to khai 04/TNCN

Private checkSoCT As Integer
Private isSheetTk As Boolean
Private checkTT As Integer

Private strMaPhongQuanLy As String
Private strTenPhongQuanLy As String

Private isTonTaiAC As Boolean ' Su dung de kiem tra xem bao cao ac co phai thay the hay khong?
Private isTKTonTai As Boolean

Private strMaSoThue, strMaDaiLyThue As String
Private isToKhaiCT As Boolean

Private isTKDA30 As Boolean  ' kiem tra QLT da co tk theo mau cu chua

Private isTKLanPS As Boolean
Private ngayPS As String

Private isToKhaiPsDaNhanTN As Boolean  ' Kiem tra cac to khai phat sinh da nhan trong ngay
' xu ly cho to khai 08, 08A/TNCN
Private isTKThang As Boolean
Private TuNgay As String
Private DenNgay As String

' NSHUNG bo xung phan giao tiep voi truc ESB
'Lay thong tin NNT tu ESB
Private xmlResultNNT As MSXML.DOMDocument
'Lay thong tin NNT tu ESB
Private xmlResultDLT As MSXML.DOMDocument
' Ket thuc NSHUNG bo xung


'****************************
'Description: StartBarcodeReader procedure start barcode listener on com 1 port
'Author:TuyenDS
'Date:
'Input:
'OutPut:
'Return:
'****************************
Private Sub StartBarcodeReader()
    Dim strSetting As String
On Error GoTo ErrHandle
    strSetting = mBaudRate & "," & mParity & "," & mDataBits & "," & mStopBits
    With MSComm1
        If .PortOpen = False Then
            .Handshaking = mHandshaking
            .CommPort = mCommPort
            .Settings = strSetting                          ' 9600 baud, no parity, 8 data, and 1 stop bit.
            .InputLen = 0                                   ' Read entire buffer
            .RThreshold = 1                                 ' Call **_OnComm for each character
            .InputMode = comInputModeBinary
            On Error GoTo PortOpenedErr ' Port in use
            .PortOpen = True      ' Opens the port
        End If
    End With
    Exit Sub
PortOpenedErr:
    DisplayMessage "0061", msOKOnly, miCriticalError
    Unload Me
    frmTreeviewMenu.Show
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "StartBarcodeReader", Err.Number, Err.Description
    Err.Raise Err.Number
End Sub

'****************************
'Description: StopBarcodeReader procedure stop barcode listener on com 1 port
'Author:TuyenDS
'Date:
'Input:
'OutPut:
'Return:
'****************************
Private Sub StopBarcodeReader()
On Error GoTo ErrHandle
    With MSComm1
        If .PortOpen = True Then .PortOpen = False
    End With
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "StopBarcodeReader", Err.Number, Err.Description
End Sub

'****************************
'Description: cmdClear_Click procedure clear current data on
'             the screen and go to next tax report.
'Author:ThanhDX
'Date:23/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub cmdClear_Click()

On Error GoTo ErrHandle
    If Not TAX_Utilities_Srv_New.Data(0) Is Nothing Then
        If MessageBox("0050", msYesNo, miQuestion) = mrYes Then
            If Not objTaxBusiness Is Nothing Then
                objTaxBusiness.Prepared4 dNgayDauKy
                'Get Params
                objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
            End If
            StartReceiveForm
        End If
    End If
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdClear_Click", Err.Number, Err.Description
End Sub

'****************************
'Description: cmdCommand2_Click procedure.
'Author:nshung
'Date:22/07/2013
'Input:
'OutPut:Standard XML
'Return:
'****************************
Private Sub cmdCommand2_Click()
    Dim xmlMapCT     As New MSXML.DOMDocument
    Dim xmlTK        As New MSXML.DOMDocument
    Dim xmlPL        As New MSXML.DOMDocument
    Dim xmlMapPL     As New MSXML.DOMDocument
    Dim xmlNodeTK    As MSXML.IXMLDOMNode
    Dim xmlNodeMapCT As MSXML.IXMLDOMNode

    Dim cSheet       As Integer, oSheet As Integer
    Dim strFileName  As String
    Dim MaTK         As String
    Dim nodeVal      As MSXML.IXMLDOMNode
    Dim blnFinish    As Boolean
    Dim sRow         As Integer
    
    Dim sKyLapBo As String
    Dim sNgayNopTK As String
    
    On Error GoTo ErrHandle
    
    CallFinish
    
    blnFinish = CheckValidData
    
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepared4 dNgayDauKy
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    
    If blnFinish = False Then
        Exit Sub
    End If
        
    MaTK = GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(0), "DataFile")

    If InStr(MaTK, "11") > 0 Then
        MaTK = Replace$(MaTK, "11", "")
    ElseIf InStr(MaTK, "10") > 0 Then
        MaTK = Replace$(MaTK, "10", "")
    End If
    
    'Chuan xml khi ket xuat cho cac to TNCN thang, quy
    'vi du 01A_TNCN -> 01_TNCN, 01B_TNCN -> 01_TNCN
    If InStr(MaTK, "TNCN") > 0 Then
        If Val(Left$(MaTK, 2)) < 6 Then
            MaTK = Replace$(Replace$(MaTK, "B", ""), "A", "")
     
        End If
    End If

    '    With CommonDialog1
    '        .CancelError = True
    '        .InitDir = GetAbsolutePath("..")
    '        .Filter = "XML file (*.xml)|*.xml"
    '        .FilterIndex = 1
    '        .DialogTitle = "File xml export to " & .InitDir
    '        .FileName = getFileName
    '        .ShowSave
    '
    '        If Right$(.FileName, 4) <> ".xml" Then
    '            strFileName = .FileName & ".xml"
    '        Else
    '            strFileName = .FileName
    '        End If
    '    End With

    strFileName = "test.xml" 'getFileName
        
    xmlTK.Load GetAbsolutePath("..\InterfaceTemplates\xml\" & MaTK & "_xml.xml")
    xmlMapCT.Load GetAbsolutePath("..\Ini\" & MaTK & "_xml.xml")

    With fpSpread1
        Dim cellid         As String
        Dim cellArray()    As String
        Dim nodeValIndex   As Integer
        Dim cellRange      As Integer
        Dim GroupCellRange As Integer

        .Sheet = 1

        ' Set value cho to khai
        For Each xmlNodeMapCT In xmlMapCT.lastChild.childNodes
            Dim xmlCellNode   As MSXML.IXMLDOMNode
            Dim xmlCellTKNode As MSXML.IXMLDOMNode
            Dim currentGroup  As String
            Dim nodePL        As MSXML.IXMLDOMNode
            Dim Blank         As Boolean
            Dim ID            As Integer
            Dim CloneNode     As New MSXML.DOMDocument
            
            'Set gia tri cho group dong
            If UCase(xmlNodeMapCT.nodeName) = "DYNAMIC" Then
                CloneNode.loadXML xmlNodeMapCT.xml
                ID = 1
                currentGroup = GetAttribute(xmlNodeMapCT, "GroupName")

                If GetAttribute(xmlNodeMapCT, "GroupCellRange") = vbNullString Then
                    GroupCellRange = 1
                Else
                    GroupCellRange = Val(GetAttribute(xmlNodeMapCT, "GroupCellRange"))
                End If

                Blank = True

                If xmlTK.getElementsByTagName(currentGroup)(0).hasChildNodes Then
                    xmlTK.getElementsByTagName(currentGroup)(0).removeChild xmlTK.getElementsByTagName(currentGroup)(0).firstChild

                End If

                Do
                    Blank = True
                    SetCloneNode CloneNode, xmlNodeMapCT, Blank, cellRange, sRow
                    .Col = .ColLetterToNumber("B")
                    .Row = sRow

                    If Blank = True Or .Text = "aa" Or .Text = "bb" Or .Text = "cc" Or .Text = "dd" Or .Text = "ee" Or .Text = "ff" Or .Text = "gg" Or .Text = "hh" Then

                        Exit Do
                    End If

                    SetAttribute CloneNode.firstChild.firstChild, "id", CStr(ID)
                    Set nodePL = xmlTK.getElementsByTagName(currentGroup)(0)
                    nodePL.appendChild CloneNode.firstChild.firstChild.CloneNode(True)
                    ID = ID + 1

                    cellRange = cellRange + GroupCellRange
                Loop

                cellRange = cellRange - GroupCellRange

            Else
                Dim xmlChildNode As MSXML.IXMLDOMNode
                currentGroup = GetAttribute(xmlNodeMapCT, "GroupName")

                For Each xmlCellNode In xmlNodeMapCT.childNodes

                    If xmlCellNode.hasChildNodes Then
                        cellid = xmlCellNode.Text
                    Else
                        cellid = ""
                    End If

                    cellArray = Split(cellid, "_")

                    If currentGroup = vbNullString Or currentGroup = "" Then
                        Set xmlCellTKNode = xmlTK.getElementsByTagName(xmlCellNode.nodeName)(0)
                    Else

                        For Each xmlChildNode In xmlTK.getElementsByTagName(xmlCellNode.nodeName)

                            If xmlChildNode.parentNode.nodeName = currentGroup Then
                                Set xmlCellTKNode = xmlChildNode
                                Exit For
                            End If

                        Next

                    End If

                    If InStr(cellid, "_") = 0 Then
                        xmlCellTKNode.Text = cellid
                    ElseIf Val(cellArray(1)) = 0 Then
                        xmlCellTKNode.Text = cellid

                    Else
                        .Col = .ColLetterToNumber(cellArray(0))
                        .Row = Val(cellArray(1)) + cellRange

                        If .CellType = CellTypeNumber Then
                            xmlCellTKNode.Text = .Value
                        Else
                            xmlCellTKNode.Text = .Text
                        End If
                    End If

                Next

            End If

        Next

        'Set value KyLapBo, NgayNopTK
        cellid = GetAttribute(xmlMapCT.lastChild, "ky_lap_bo")
        cellArray = Split(cellid, "_")
        .Col = .ColLetterToNumber(cellArray(0))
        .Row = Val(cellArray(1))
        
        sKyLapBo = .Text
        
        cellid = GetAttribute(xmlMapCT.lastChild, "ngay_nop_tk")
        cellArray = Split(cellid, "_")
        .Col = .ColLetterToNumber(cellArray(0))
        .Row = Val(cellArray(1))
        
        sNgayNopTK = .Text

        'Set value cho phu luc
        For nodeValIndex = 1 To TAX_Utilities_Srv_New.NodeValidity.childNodes.length
            Set nodeVal = TAX_Utilities_Srv_New.NodeValidity.childNodes(nodeValIndex)

            If GetAttribute(nodeVal, "Active") = "1" Then
                Dim currentRow As Integer
                Dim xmlSection As MSXML.IXMLDOMNode
        
                MaTK = GetAttribute(nodeVal, "DataFile")

                If InStr(MaTK, "11") > 0 Then
                    MaTK = Replace$(MaTK, "11", "")
                ElseIf InStr(MaTK, "10") > 0 Then
                    MaTK = Replace$(MaTK, "10", "")
                End If
                If InStr(MaTK, "KHBS") > 0 Then
                    MaTK = "KHBS"
                End If
                xmlPL.Load GetAbsolutePath("..\InterfaceTemplates\xml\" & MaTK & "_xml.xml")

                xmlMapPL.Load GetAbsolutePath("..\ini\" & MaTK & "_xml.xml")

                cellRange = 0
                .Sheet = nodeValIndex + 1

                For Each xmlSection In xmlMapPL.lastChild.childNodes

                    If UCase(xmlSection.nodeName) = "DYNAMIC" Then
                        CloneNode.loadXML xmlSection.xml
                        ID = 1
                        currentGroup = GetAttribute(xmlSection, "GroupName")

                        If GetAttribute(xmlSection, "GroupCellRange") = vbNullString Then
                            GroupCellRange = 1
                        Else
                            GroupCellRange = Val(GetAttribute(xmlSection, "GroupCellRange"))
                        End If

                        Blank = True

                        If xmlPL.getElementsByTagName(currentGroup)(0).hasChildNodes Then
                            xmlPL.getElementsByTagName(currentGroup)(0).removeChild xmlPL.getElementsByTagName(currentGroup)(0).firstChild
                        End If

                        Do
                            Blank = True
                            SetCloneNode CloneNode, xmlSection, Blank, cellRange, sRow
                            
                            .Col = .ColLetterToNumber("B")
                            .Row = sRow

                            If Blank = True Or .Text = "aa" Or .Text = "bb" Or .Text = "cc" Or .Text = "dd" Or .Text = "ee" Or .Text = "ff" Or .Text = "gg" Or .Text = "hh" Then

                                Exit Do
                            End If

                            SetAttribute CloneNode.firstChild.firstChild, "id", CStr(ID)
                            Set nodePL = xmlPL.getElementsByTagName(currentGroup)(0)
                            nodePL.appendChild CloneNode.firstChild.firstChild.CloneNode(True)
                            ID = ID + 1
                            cellRange = cellRange + GroupCellRange
                        Loop

                        cellRange = cellRange - GroupCellRange
                    Else
                        Dim xmlChildNodePL As MSXML.IXMLDOMNode
                        currentGroup = GetAttribute(xmlNodeMapCT, "GroupName")

                        For Each xmlCellNode In xmlSection.childNodes

                            If xmlCellNode.hasChildNodes Then
                                cellid = xmlCellNode.Text
                            Else
                                cellid = ""
                            End If

                            cellArray = Split(cellid, "_")

                            If currentGroup = vbNullString Or currentGroup = "" Then
                                Set xmlCellTKNode = xmlPL.getElementsByTagName(xmlCellNode.nodeName)(0)
                            Else

                                For Each xmlChildNodePL In xmlTK.getElementsByTagName(xmlCellNode.nodeName)

                                    If xmlChildNodePL.parentNode.nodeName = currentGroup Then
                                        Set xmlCellTKNode = xmlChildNodePL
                                        Exit For
                                    End If

                                Next

                            End If

                            If InStr(cellid, "_") = 0 Then
                                xmlCellTKNode.Text = cellid
                            ElseIf Val(cellArray(1)) = 0 Then
                                xmlCellTKNode.Text = cellid
                            Else
                                .Col = .ColLetterToNumber(cellArray(0))
                                .Row = Val(cellArray(1)) + cellRange

                                If .CellType = CellTypeNumber Then
                                    xmlCellTKNode.Text = .Value
                                Else
                                    xmlCellTKNode.Text = .Text
                                End If
                            End If

                        Next

                    End If

                Next

                xmlTK.getElementsByTagName("PLuc")(0).appendChild xmlPL.lastChild
            End If

        Next

    End With    'Save temp
    SetValueToKhaiHeader xmlTK

    Dim sFileName As String
    sFileName = "c:\TempXML\" & strFileName
    Dim xmlDocSave As New MSXML.DOMDocument
    Set xmlDocSave = AppendXMLStandard(xmlTK, sKyLapBo, sNgayNopTK)
    xmlDocSave.save sFileName
    
    ' Push MQ
    'PushDataToESB xmlDocSave.xml
    ' End push
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdExportXML_Click", Err.Number, Err.Description

End Sub

Private Sub SetCloneNode(ByRef CloneNode As MSXML.DOMDocument, _
                         ByVal nodes As MSXML.IXMLDOMNode, _
                         ByRef Blank As Boolean, _
                         ByVal cellRange As Integer, _
                         ByRef Row As Integer)
    Dim cellid      As String
    Dim cellArray() As String
    Dim cNode       As MSXML.IXMLDOMNode
    Dim dNode       As MSXML.IXMLDOMNode

    With fpSpread1

        For Each cNode In nodes.childNodes

            If cNode.hasChildNodes Then
                If cNode.firstChild.hasChildNodes Then
                    SetCloneNode CloneNode, cNode, Blank, cellRange, Row
                Else
                    cellid = cNode.Text
                    cellArray = Split(cellid, "_")
                    
                    If InStr(cellid, "_") = 0 Then
                        CloneNode.getElementsByTagName(cNode.nodeName)(0).Text = cellid
                    ElseIf Val(cellArray(1)) = 0 Then
                        CloneNode.getElementsByTagName(cNode.nodeName)(0).Text = cellid

                    Else
                        .Col = .ColLetterToNumber(cellArray(0))
                        .Row = Val(cellArray(1)) + cellRange
                        
                        For Each dNode In CloneNode.getElementsByTagName(cNode.nodeName)

                            If dNode.parentNode.nodeName = cNode.parentNode.nodeName Then
                        
                                If .CellType = CellTypeNumber Then
                                
                                    dNode.Text = .Value
                                ElseIf .CellType = CellTypeCheckBox Then

                                    If LCase$(.Text) = "x" Then
                                        dNode.Text = "1"
                                    ElseIf .Text = "" Then
                                        dNode.Text = "0"
                                        Else
                                        dNode.Text = .Text
                                    End If

                                Else
                                    dNode.Text = .Text
                            
                                End If
                            End If

                        Next

                    End If

                    If .Text <> "" And .Text <> vbNullString Then
                        If .CellType = CellTypeNumber Then
                            If .Text <> "0" Then
                                Blank = False
                    
                            End If

                        ElseIf .CellType = CellTypeDate Then

                            If .Text <> "../../...." Then
                                Blank = False
                        
                            End If

                        Else
                    
                            Blank = False
                        End If
                    End If
                    
                    Row = .Row
                    
                End If
           
            End If

        Next

    End With

End Sub

' Set gia tri mac dinh cho to khai xml
Private Sub SetValueToKhaiHeader(ByVal xmlTK As MSXML.DOMDocument)
    Dim vlue As Variant
    On Error GoTo ErrHandle
    
    'Set value from config, webservices ESB
    Dim xmlConfig As New MSXML.DOMDocument
    Set xmlConfig = LoadConfig()

    xmlTK.getElementsByTagName("pbanDVu")(0).Text = APP_VERSION

    xmlTK.getElementsByTagName("maCQTNoiNop")(0).Text = strMaCoQuanThue 'xmlConfig.getElementsByTagName("maCQTNoiNop")(0).Text
    xmlTK.getElementsByTagName("tenCQTNoiNop")(0).Text = strTenCoQuanThue 'xmlConfig.getElementsByTagName("tenCQTNoiNop")(0).Text
    xmlTK.getElementsByTagName("ngayLapTKhai")(0).Text = Format(Date, "dd-mmm-yyyy HH:mm:ss")
    
    If (xmlResultNNT.hasChildNodes And (InStr(xmlResultNNT.xml, "fault_code") <= 0)) Then
        xmlTK.getElementsByTagName("maHuyenNNT")(0).Text = xmlResultNNT.getElementsByTagName("MaQuanHuyen")(0).Text
        xmlTK.getElementsByTagName("maTinhNNT")(0).Text = xmlResultNNT.getElementsByTagName("MaTinh")(0).Text
        
        'xmlTK.getElementsByTagName("tenNNT")(0).Text = "test"
        xmlTK.getElementsByTagName("tenNNT")(0).Text = xmlResultNNT.getElementsByTagName("TenNNT")(0).Text
        ' xmlTK.getElementsByTagName("dchiNNT")(0).Text = "test"
        xmlTK.getElementsByTagName("dchiNNT")(0).Text = xmlResultNNT.getElementsByTagName("DiaChi")(0).Text
        xmlTK.getElementsByTagName("dthoaiNNT")(0).Text = xmlResultNNT.getElementsByTagName("DienThoai")(0).Text
        xmlTK.getElementsByTagName("faxNNT")(0).Text = xmlResultNNT.getElementsByTagName("Fax")(0).Text
        xmlTK.getElementsByTagName("emailNNT")(0).Text = xmlResultNNT.getElementsByTagName("Email")(0).Text
        'xmlTK.getElementsByTagName("mst")(0).Text = xmlResultNNT.getElementsByTagName("MaSoThue")(0).Text
        
        xmlTK.getElementsByTagName("tenHuyenNNT")(0).Text = xmlResultNNT.getElementsByTagName("TenQuanHuyen")(0).Text
        xmlTK.getElementsByTagName("tenTinhNNT")(0).Text = xmlResultNNT.getElementsByTagName("TenTinh")(0).Text
    End If
    
    xmlTK.getElementsByTagName("mst")(0).Text = strMaNNT
    
    If (xmlResultDLT.hasChildNodes And (InStr(xmlResultDLT.xml, "fault_code") <= 0)) Then
        ' xmlTK.getElementsByTagName("tenDLyThue")(0).Text = "test"
        xmlTK.getElementsByTagName("tenDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("TenNNT")(0).Text
        'xmlTK.getElementsByTagName("dchiDLyThue")(0).Text = "test"
        xmlTK.getElementsByTagName("dchiDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("DiaChi")(0).Text
        xmlTK.getElementsByTagName("dthoaiDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("DienThoai")(0).Text
        xmlTK.getElementsByTagName("faxDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("Fax")(0).Text
        xmlTK.getElementsByTagName("emailDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("Email")(0).Text
        xmlTK.getElementsByTagName("soHDongDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("SoHopDong")(0).Text
        xmlTK.getElementsByTagName("ngayKyHDDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("NgayHopDong")(0).Text
        xmlTK.getElementsByTagName("tenTinhDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("TenTinh")(0).Text
        xmlTK.getElementsByTagName("tenHuyenDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("TenQuanHuyen")(0).Text
        xmlTK.getElementsByTagName("maHuyenDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("MaQuanHuyen")(0).Text
        xmlTK.getElementsByTagName("maTinhDLyThue")(0).Text = xmlResultDLT.getElementsByTagName("MaTinh")(0).Text
    End If
    
    xmlTK.getElementsByTagName("mstDLyThue")(0).Text = strMaDLT
    xmlTK.getElementsByTagName("pbanTKhaiXML")(0).Text = "1.0"
    xmlTK.getElementsByTagName("maDVu")(0).Text = GetAttribute(GetMessageCellById("0133"), "Msg")
    xmlTK.getElementsByTagName("tenDVu")(0).Text = GetAttribute(GetMessageCellById("0134"), "Msg")
    xmlTK.getElementsByTagName("ttinNhaCCapDVu")(0).Text = ""
    
    vlue = xmlTK.getElementsByTagName("soLan")(0).Text
    
    If Val(vlue) > 0 Then
        xmlTK.getElementsByTagName("loaiTKhai")(0).Text = GetAttribute(GetMessageCellById("0132"), "Msg")
        xmlTK.getElementsByTagName("soLan")(0).Text = Val(vlue)
    Else
        xmlTK.getElementsByTagName("soLan")(0).Text = ""
        xmlTK.getElementsByTagName("loaiTKhai")(0).Text = GetAttribute(GetMessageCellById("0131"), "Msg")
    End If
    
    'To 03/TBAC
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "67" Then
        xmlTK.getElementsByTagName("kyKKhai")(0).Text = ""
        xmlTK.getElementsByTagName("kyKKhaiTuNgay")(0).Text = ""
        xmlTK.getElementsByTagName("kyKKhaiDenNgay")(0).Text = ""
        xmlTK.getElementsByTagName("kieuKy")(0).Text = ""

        'Xu ly rieng cho truong hop to khai 01_TAIN_DK
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "92" Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "98" Then

        If xmlTK.getElementsByTagName("ct03").length > 0 Then
            If xmlTK.getElementsByTagName("ct03")(0).Text = "1" Then
                xmlTK.getElementsByTagName("kyKKhai")(0).Text = GetKyKeKhai(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID"))
                xmlTK.getElementsByTagName("kyKKhaiTuNgay")(0).Text = Format$(dNgayDauKy, "dd/MM/yyyy")
                xmlTK.getElementsByTagName("kyKKhaiDenNgay")(0).Text = Format$(dNgayCuoiKy, "dd/MM/yyyy")
                xmlTK.getElementsByTagName("kieuKy")(0).Text = strKieuKy
            Else
                fpSpread1.Col = fpSpread1.ColLetterToNumber("D")
                fpSpread1.Row = 39
                xmlTK.getElementsByTagName("kyKKhai")(0).Text = fpSpread1.Text
                xmlTK.getElementsByTagName("kieuKy")(0).Text = "D"
            End If
        End If

    Else
        xmlTK.getElementsByTagName("kyKKhai")(0).Text = GetKyKeKhai(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID"))

        If strKieuKy <> "D" Then
            xmlTK.getElementsByTagName("kyKKhaiTuNgay")(0).Text = Format$(dNgayDauKy, "dd/MM/yyyy")
            xmlTK.getElementsByTagName("kyKKhaiDenNgay")(0).Text = Format$(dNgayCuoiKy, "dd/MM/yyyy")
           
        End If

        xmlTK.getElementsByTagName("kieuKy")(0).Text = strKieuKy
  
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetValueToKhaiHeader", Err.Number, Err.Description
End Sub

'Lay ky ke khai
Private Function GetKyKeKhai(ByVal ID_TK As String) As String
    Dim KYKKHAI As String
    On Error GoTo ErrHandle

    If isTKLanPS = True Then
        KYKKHAI = ngayPS
        strKieuKy = "D"
    ElseIf ID_TK = "01" Or ID_TK = "02" Or ID_TK = "04" Or ID_TK = "71" Or ID_TK = "36" Or ID_TK = "68" Then

        If LoaiKyKK = True Then
            KYKKHAI = TAX_Utilities_Srv_New.ThreeMonths & "/" & TAX_Utilities_Srv_New.Year
            strKieuKy = "Q"
        Else
            KYKKHAI = TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year
            strKieuKy = "M"
        End If
            
    Else

        If (Trim(TAX_Utilities_Srv_New.Month) <> vbNullString Or Trim(TAX_Utilities_Srv_New.Month) <> "") And (Trim(TAX_Utilities_Srv_New.ThreeMonths) = vbNullString Or Trim(TAX_Utilities_Srv_New.ThreeMonths) = "") Then
            KYKKHAI = TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year
            strKieuKy = "M"
        ElseIf (Trim(TAX_Utilities_Srv_New.Month) = vbNullString Or Trim(TAX_Utilities_Srv_New.Month) = "") And (Trim(TAX_Utilities_Srv_New.ThreeMonths) <> vbNullString Or Trim(TAX_Utilities_Srv_New.ThreeMonths) <> "") Then
            KYKKHAI = TAX_Utilities_Srv_New.ThreeMonths & "/" & TAX_Utilities_Srv_New.Year
            strKieuKy = "Q"
        Else
            KYKKHAI = TAX_Utilities_Srv_New.Year
            strKieuKy = "Y"
        End If

    End If
 
    GetKyKeKhai = KYKKHAI

    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetKyKeKhai", Err.Number, Err.Description

End Function
Private Function getFileName() As String
    Dim strDataFileName As String
    Dim lSheet As Integer
    
    On Error GoTo ErrHandle
    lSheet = 0
    If strKHBS = "TKBS" Then
        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "0" Then
            strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & ".xml"
        Else

            If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") <> "1" Then
                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "95" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "71" Then

                    If strQuy = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                    End If

                Else
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
                End If

            ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then

                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "75" Then

                    ' To khai 08/TNCN co to khai tu thang va to khai quy
                    If strQuy = "TK_TU_THANG" Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & Replace(TAX_Utilities_Srv_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_Srv_New.LastDay, "/", "") & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                    End If

                Else
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                End If

            ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") <> "1" Then

                'Data file contain Day from and to.
                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "82" Then
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & Replace(TAX_Utilities_Srv_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_Srv_New.LastDay, "/", "") & ".xml"
                Else
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Year & "_" & Replace(TAX_Utilities_Srv_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_Srv_New.LastDay, "/", "") & ".xml"
                End If

            ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
                'Data file contain Day.
                strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Day & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
            Else
                'Data file not contain Day from and to.
                strDataFileName = TAX_Utilities_Srv_New.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Year & ".xml"
                '*********************************
            End If
        End If

    Else

        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "0" Then
            strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & ".xml"
        Else

            If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") <> "1" Then
                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "04" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "95" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "71" Then

                    If strQuy = "TK_THANG" Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
                    ElseIf strQuy = "TK_QUY" Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                    End If

                Else
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
                End If

            ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then

                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "75" Then

                    ' To khai 08/TNCN co to khai tu thang va to khai quy
                    If strQuy = "TK_TU_THANG" Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & Replace(TAX_Utilities_Srv_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_Srv_New.LastDay, "/", "") & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                    End If

                ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "73" Then

                    ' To khai 02/TNDN
                    If isTKLanPS = True Then
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & ngayPS & ".xml"
                    Else
                        strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                    End If

                Else
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_0" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                End If

            ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") <> "1" Then

                'Data file contain Day from and to.
                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "82" Then
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & Replace(TAX_Utilities_Srv_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_Srv_New.LastDay, "/", "") & ".xml"
                Else
                    strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Year & "_" & Replace(TAX_Utilities_Srv_New.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_Srv_New.LastDay, "/", "") & ".xml"
                End If

            ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
                'Data file contain Day.
                strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Day & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
            Else
                'Data file not contain Day from and to.
                strDataFileName = TAX_Utilities_Srv_New.DataFolder & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_XML" & "_" & TAX_Utilities_Srv_New.Year & ".xml"
                '*********************************
            End If
        End If
    End If

    getFileName = strDataFileName
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetFileName", Err.Number, Err.Description

End Function


'****************************
'Description: cmdExit_Click procedure.
'Author:ThanhDX
'Date:23/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub cmdExit_Click()
    Dim intReturn As Integer
    Dim blnExit As Boolean
    
On Error GoTo ErrHandle
    
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepared4 dNgayDauKy
        ' Get Params
         objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    
    If Not blnReceiveByBarcode Then 'Receive data from file
        If UBound(arrStrElements) > 0 Then '
            If MessageBox("0055", msYesNo, miQuestion) = mrNo Then _
                Exit Sub
            blnExit = True
        End If
    End If

    If Not TAX_Utilities_Srv_New.Data(0) Is Nothing Then
        Select Case MessageBox("0052", msYesNoCancel, miQuestion)
            Case 1 ' Cancel
                Exit Sub
            Case 3 'No
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
            Case 6 'Yes
                cmdSave_Click
                Unload Me
                frmTreeviewMenu.Show
                Exit Sub
        End Select
    End If
    
    If blnExit Then
        Unload Me
        frmTreeviewMenu.Show
        Exit Sub
    End If
    
    If blnReceiveByBarcode Then
        If MessageBox("0051", msYesNo, miQuestion) = mrYes Then
            Unload Me
            frmTreeviewMenu.Show
            Exit Sub
        End If
    Else
        Unload Me
        frmTreeviewMenu.Show
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdExit_Click", Err.Number, Err.Description
End Sub

Private Sub ExecuteSave()
    Dim xmlMapCT     As New MSXML.DOMDocument
    Dim xmlTK        As New MSXML.DOMDocument
    Dim xmlPL        As New MSXML.DOMDocument
    Dim xmlMapPL     As New MSXML.DOMDocument
    Dim xmlNodeTK    As MSXML.IXMLDOMNode
    Dim xmlNodeMapCT As MSXML.IXMLDOMNode

    Dim cSheet       As Integer, oSheet As Integer
    Dim strFileName  As String
    Dim MaTK         As String
    Dim nodeVal      As MSXML.IXMLDOMNode
    Dim blnFinish    As Boolean
    Dim Level        As String
    Dim sRow         As Integer
    
    Dim sKyLapBo     As String
    Dim sNgayNopTK   As String
    
    On Error GoTo ErrHandle
    
    CallFinish
    
    blnFinish = CheckValidData
    
    If blnFinish = False Then
        Exit Sub
    End If
        
    MaTK = GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(0), "DataFile")

    If InStr(MaTK, "11") > 0 Then
        MaTK = Replace$(MaTK, "11", "")
    ElseIf InStr(MaTK, "10") > 0 Then
        MaTK = Replace$(MaTK, "10", "")
    ElseIf InStr(MaTK, "12") > 0 Then
        MaTK = Replace$(MaTK, "12", "")
    ElseIf InStr(MaTK, "13") > 0 Then
        MaTK = Replace$(MaTK, "13", "")
    End If
    
    '    With CommonDialog1
    '        .CancelError = True
    '        .InitDir = GetAbsolutePath("..")
    '        .Filter = "XML file (*.xml)|*.xml"
    '        .FilterIndex = 1
    '        .DialogTitle = "File xml export to " & .InitDir
    '        .FileName = getFileName
    '        .ShowSave
    '
    '        If Right$(.FileName, 4) <> ".xml" Then
    '            strFileName = .FileName & ".xml"
    '        Else
    '            strFileName = .FileName
    '        End If
    '    End With

    strFileName = "ToKhai.xml" 'getFileName
        
    xmlTK.Load GetAbsolutePath("..\InterfaceTemplates\xml\" & MaTK & "_xml.xml")
    xmlMapCT.Load GetAbsolutePath("..\Ini\" & MaTK & "_xml.xml")
   
    With fpSpread1
        Dim cellid         As String
        Dim cellArray()    As String
        Dim nodeValIndex   As Integer
        Dim cellRange      As Integer
        Dim GroupCellRange As Integer

        .Sheet = 1

        ' Set value cho to khai
        For Each xmlNodeMapCT In xmlMapCT.lastChild.childNodes
            Dim xmlCellNode   As MSXML.IXMLDOMNode
            Dim xmlCellTKNode As MSXML.IXMLDOMNode
            Dim currentGroup  As String
            Dim nodePL        As MSXML.IXMLDOMNode
            Dim Blank         As Boolean
            Dim ID            As Integer
            Dim CloneNode     As New MSXML.DOMDocument
            
            'Set gia tri cho group dong
            If UCase(xmlNodeMapCT.nodeName) = "DYNAMIC" Then
                ID = 1
                currentGroup = GetAttribute(xmlNodeMapCT, "GroupName")
                Level = GetAttribute(xmlNodeMapCT, "Level")

                CloneNode.loadXML xmlNodeMapCT.firstChild.xml

                If GetAttribute(xmlNodeMapCT, "GroupCellRange") = vbNullString Then
                    GroupCellRange = 1
                Else
                    GroupCellRange = Val(GetAttribute(xmlNodeMapCT, "GroupCellRange"))
                End If

                Blank = True

                If xmlTK.getElementsByTagName(currentGroup)(0).hasChildNodes Then
                    If Level = "2" Then
                        xmlTK.getElementsByTagName(currentGroup)(0).firstChild.removeChild xmlTK.getElementsByTagName(currentGroup)(0).firstChild.firstChild

                    Else
                        xmlTK.getElementsByTagName(currentGroup)(0).removeChild xmlTK.getElementsByTagName(currentGroup)(0).firstChild

                    End If

                End If

                Do
                    Blank = True
                    SetCloneNode CloneNode, xmlNodeMapCT, Blank, cellRange, sRow
                    .Col = .ColLetterToNumber("B")
                    .Row = sRow

                    If Blank = True Or .Text = "aa" Or .Text = "bb" Or .Text = "cc" Or .Text = "dd" Or .Text = "ee" Or .Text = "ff" Then
                        If ID > 1 Then
                            cellRange = cellRange - GroupCellRange
                        End If

                        Exit Do
                    End If

                    SetAttribute CloneNode.firstChild, "id", CStr(ID)

                    If Level = "2" Then
                        xmlTK.getElementsByTagName(currentGroup)(0).firstChild.appendChild CloneNode.firstChild.CloneNode(True)
                    Else
                        xmlTK.getElementsByTagName(currentGroup)(0).appendChild CloneNode.firstChild.CloneNode(True)
                    End If

                    ID = ID + 1

                    cellRange = cellRange + GroupCellRange
                Loop
                
            Else
                Dim xmlChildNode As MSXML.IXMLDOMNode
                currentGroup = GetAttribute(xmlNodeMapCT, "GroupName")

                For Each xmlCellNode In xmlNodeMapCT.childNodes

                    If xmlCellNode.hasChildNodes Then
                        cellid = xmlCellNode.Text
                    Else
                        cellid = ""
                    End If

                    cellArray = Split(cellid, "_")

                    If currentGroup = vbNullString Or currentGroup = "" Then
                        Set xmlCellTKNode = xmlTK.getElementsByTagName(xmlCellNode.nodeName)(0)
                    Else

                        For Each xmlChildNode In xmlTK.getElementsByTagName(xmlCellNode.nodeName)

                            If xmlChildNode.parentNode.nodeName = currentGroup Then
                                Set xmlCellTKNode = xmlChildNode
                                Exit For
                            End If

                        Next

                    End If

                    If UBound(cellArray) <> 1 Or Len(cellid) > 5 Then
                        
                        xmlCellTKNode.Text = cellid
                    Else
                        .Col = .ColLetterToNumber(cellArray(0))
                        .Row = Val(cellArray(1)) + cellRange

                        If .CellType = CellTypeNumber Then
                            xmlCellTKNode.Text = .Value
                        ElseIf .CellType = CellTypeCheckBox Then

                            If LCase$(.Text) = "x" Then
                                xmlCellTKNode.Text = "1"
                            ElseIf .Text = "" Then
                                xmlCellTKNode.Text = "0"
                            Else
                                xmlCellTKNode.Text = .Text
                            End If

                        Else
                            xmlCellTKNode.Text = .Text
                        End If
                    End If

                Next

            End If

        Next
        
        'Set gia tri header cho to khai
        SetValueToKhaiHeader xmlTK

        'Set value KyLapBo, NgayNopTK
        cellid = GetAttribute(xmlMapCT.lastChild, "ky_lap_bo")

        If cellid <> vbNullString Then
            cellArray = Split(cellid, "_")
            .Col = .ColLetterToNumber(cellArray(0))
            .Row = Val(cellArray(1))
            sKyLapBo = .Text
        End If

        cellid = GetAttribute(xmlMapCT.lastChild, "ngay_nop_tk")

        If cellid <> vbNullString Then

            cellArray = Split(cellid, "_")
            .Col = .ColLetterToNumber(cellArray(0))
            .Row = Val(cellArray(1))
            sNgayNopTK = .Text
        End If

        'Set value cho phu luc
        For nodeValIndex = 1 To TAX_Utilities_Srv_New.NodeValidity.childNodes.length
            Set nodeVal = TAX_Utilities_Srv_New.NodeValidity.childNodes(nodeValIndex)

            If GetAttribute(nodeVal, "Active") = "1" Then
                Dim currentRow As Integer
                Dim xmlSection As MSXML.IXMLDOMNode
        
                MaTK = GetAttribute(nodeVal, "DataFile")

                If InStr(MaTK, "11") > 0 Then
                    MaTK = Replace$(MaTK, "11", "")
                ElseIf InStr(MaTK, "10") > 0 Then
                    MaTK = Replace$(MaTK, "10", "")
                ElseIf InStr(MaTK, "12") > 0 Then
                    MaTK = Replace$(MaTK, "12", "")
                ElseIf InStr(MaTK, "13") > 0 Then
                    MaTK = Replace$(MaTK, "13", "")
                End If
                
                If InStr(MaTK, "KHBS") > 0 Then
                    MaTK = "KHBS"
                End If

                xmlPL.Load GetAbsolutePath("..\InterfaceTemplates\xml\" & MaTK & "_xml.xml")

                xmlMapPL.Load GetAbsolutePath("..\ini\" & MaTK & "_xml.xml")

                If xmlPL.hasChildNodes = True And xmlMapPL.hasChildNodes = True Then
                    cellRange = 0
                    .Sheet = nodeValIndex + 1

                    For Each xmlSection In xmlMapPL.lastChild.childNodes

                        If UCase(xmlSection.nodeName) = "DYNAMIC" Then
                            ID = 1
                            currentGroup = GetAttribute(xmlSection, "GroupName")
                            Level = GetAttribute(xmlSection, "Level")

                            CloneNode.loadXML xmlSection.firstChild.xml

                            If GetAttribute(xmlSection, "GroupCellRange") = vbNullString Then
                                GroupCellRange = 1
                            Else
                                GroupCellRange = Val(GetAttribute(xmlSection, "GroupCellRange"))
                            End If

                            Blank = True

                            If xmlPL.getElementsByTagName(currentGroup)(0).hasChildNodes Then
                                xmlPL.getElementsByTagName(currentGroup)(0).removeChild xmlPL.getElementsByTagName(currentGroup)(0).firstChild
                            End If

                            Do
                                Blank = True
                                SetCloneNode CloneNode, xmlSection, Blank, cellRange, sRow
                            
                                .Col = .ColLetterToNumber("B")
                                .Row = sRow

                                If Blank = True Or .Text = "aa" Or .Text = "bb" Or .Text = "cc" Or .Text = "dd" Or .Text = "ee" Or .Text = "ff" Then
                                    If ID > 1 Then
                                        cellRange = cellRange - GroupCellRange
                                    End If

                                    Exit Do
                                End If

                                SetAttribute CloneNode.firstChild, "id", CStr(ID)

                                If Level = "2" Then
                                    xmlPL.getElementsByTagName(currentGroup)(0).firstChild.appendChild CloneNode.firstChild.CloneNode(True)
                                Else
                                    xmlPL.getElementsByTagName(currentGroup)(0).appendChild CloneNode.firstChild.CloneNode(True)
                                End If

                                ID = ID + 1
                                cellRange = cellRange + GroupCellRange
                            Loop
                        
                        Else
                            Dim xmlChildNodePL As MSXML.IXMLDOMNode
                            currentGroup = GetAttribute(xmlSection, "GroupName")

                            For Each xmlCellNode In xmlSection.childNodes

                                If xmlCellNode.hasChildNodes Then
                                    cellid = xmlCellNode.Text
                                Else
                                    cellid = ""
                                End If

                                cellArray = Split(cellid, "_")

                                If currentGroup = vbNullString Or currentGroup = "" Then
                                    Set xmlCellTKNode = xmlPL.getElementsByTagName(xmlCellNode.nodeName)(0)
                                Else

                                    For Each xmlChildNodePL In xmlPL.getElementsByTagName(xmlCellNode.nodeName)

                                        If xmlChildNodePL.parentNode.nodeName = currentGroup Then
                                            Set xmlCellTKNode = xmlChildNodePL
                                            Exit For
                                        End If

                                    Next

                                End If

                                If UBound(cellArray) <> 1 Or Len(cellid) > 5 Then
                                    xmlCellTKNode.Text = cellid
                                Else
                                    .Col = .ColLetterToNumber(cellArray(0))
                                    .Row = Val(cellArray(1)) + cellRange

                                    If .CellType = CellTypeNumber Then
                                        xmlCellTKNode.Text = .Value
                                    ElseIf .CellType = CellTypeCheckBox Then

                                        If LCase$(.Text) = "x" Then
                                            xmlCellTKNode.Text = "1"
                                        ElseIf .Text = "" Then
                                            xmlCellTKNode.Text = "0"
                                        Else
                                            xmlCellTKNode.Text = .Text
                                        End If

                                    Else
                                        xmlCellTKNode.Text = .Text
                                    End If
                                End If

                            Next

                        End If

                    Next

                    xmlTK.getElementsByTagName("PLuc")(0).appendChild xmlPL.lastChild
           
                End If
            End If

        Next

    End With    'Save temp

    If (Dir("c:\TempXML\", vbDirectory) = "") Then
        MkDir "c:\TempXML\"
    End If

    Dim sFileName As String
    sFileName = "c:\TempXML\" & strFileName
    Dim xmlDocSave As New MSXML.DOMDocument
    Set xmlDocSave = AppendXMLStandard(xmlTK, sKyLapBo, sNgayNopTK)
    xmlDocSave.save sFileName

    ' Push MQ
'    If (Not PushDataToESB(xmlDocSave.xml)) Then
'        MessageBox "0137", msOKOnly, miCriticalError
'    End If

    ' End push
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Execute_save", Err.Number, Err.Description
End Sub
Private Sub cmdSave_Click()

On Error GoTo ErrHandle



    Dim strSQL As String, mResult As Integer, strSQL_HDR As String, strSQL_DTL As String
    Dim HdrID As Variant, strDate() As String, dDate As Date
    Dim rs As New ADODB.Recordset, i As Long
    Dim qBoSung As Variant
    Dim msgRs As MsgBoxResult
    Dim idToKhai As Integer
    
    Dim mTemp As Integer
    
    Dim dsTK_DLT As String
    'dntai them bien de luu ngay dau nam tai chinh va ngay cuoi nam tai chinh
    Dim dNgayDauNamTC As Date
    Dim dNgayCuoiNamTC As Date
    Dim varDate1 As String
    Dim varDate2 As String
    '***************************
    'Date:23/11/2005
    If TAX_Utilities_Srv_New.Data(0) Is Nothing Then Exit Sub
    '***************************
    
    blnSaveSuccess = False
    
    CallFinish
    
    
    '***************************
    'Date:02/01/2006
    If Not objTaxBusiness Is Nothing Then
        objTaxBusiness.Prepared4 dNgayDauKy
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    '***************************
     If CheckValidData = False Then
        MessageBox "0046", msOKOnly, miWarning
        Exit Sub
    End If
       
    
    ' Kiem tra trang thai 02 -> khong cho ghi, trang thai 03 canh bao van cho ghi
    If checkTT = 2 Then
        'MessageBox "0108", msOKOnly, miCriticalError
        mTemp = MessageBox("0108", msYesNo, miQuestion)
        If mTemp = mrNo Then
            Exit Sub
        End If
    End If
    ' end
    
    dsTK_DLT = "~1~2~3~4~5~6~11~12~46~47~48~49~15~16~50~51~36~70~71~72~73~74~75~80~81~82~77~86~87~89~42~43~17~59~41~76~90~"
    ' Kiem tra neu MDL thue khac thi canh bao
    idToKhai = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    'If IdToKhai = 1 Or IdToKhai = 2 Or IdToKhai = 4 Or IdToKhai = 11 Or IdToKhai = 12 Or IdToKhai = 46 Or IdToKhai = 47 Or IdToKhai = 48 Or IdToKhai = 49 Or IdToKhai = 15 Or IdToKhai = 16 Or IdToKhai = 50 Or IdToKhai = 51 _
    '    Or IdToKhai = 36 Or IdToKhai = 70 Or IdToKhai = 6 Or IdToKhai = 5 Then
    If InStr(1, dsTK_DLT, "~" & idToKhai & "~", vbTextCompare) > 0 Then
        If isMaDLT(strMaSoThue, strMaDaiLyThue) = False Then
            mTemp = MessageBox("0115", msYesNo, miQuestion)
            If mTemp = mrNo Then
                Exit Sub
            End If
        End If
        
        ' Kiem tra to khai BS'
        If verToKhai = 2 Then
            If isToKhaiCT = False Then
                 MessageBox "0116", msOKOnly, miWarning
                 Exit Sub
            End If
        End If
        
        ' Kiem tra xem ben QLT da co to khai theo mau cu chua
        If isTKDA30 = True Then
            MessageBox "0114", msOKOnly, miWarning
            Exit Sub
        End If
        ' end
    End If
    ' End
    ' Cac to khai PIT se khong nhan to khai co ky ke khai < thang 7 hoac quy 3
    If TAX_Utilities_Srv_New.isCheckPIT = True Then
        If idToKhai = 46 Or idToKhai = 48 Or idToKhai = 15 Or idToKhai = 50 Or idToKhai = 36 Then
            If TAX_Utilities_Srv_New.Year < 2011 Or (TAX_Utilities_Srv_New.Year = 2011 And TAX_Utilities_Srv_New.Month < 7) Then
                MessageBox "0118", msOKOnly, miWarning
                Exit Sub
            End If
        End If
        If idToKhai = 47 Or idToKhai = 49 Or idToKhai = 16 Or idToKhai = 51 Or (idToKhai = 74 And isTKThang = False) Or (idToKhai = 75 And isTKThang = False) Then
            If TAX_Utilities_Srv_New.Year < 2011 Or (TAX_Utilities_Srv_New.Year = 2011 And TAX_Utilities_Srv_New.ThreeMonths < 3) Then
                    MessageBox "0119", msOKOnly, miWarning
                Exit Sub
            End If
        End If
            
        If ((idToKhai = 74 Or idToKhai = 75) And isTKThang = True) Then
              Dim arrNgay() As String
              arrNgay = Split(TuNgay, "/")
              
          
              If Val(arrNgay(1)) < 2011 Or (Val(arrNgay(1)) = 2011 And Val(arrNgay(0)) < 7) Then
                  MessageBox "0118", msOKOnly, miWarning
                  Exit Sub
              End If
          End If
    End If
    ' end
  
    

    ' Tam thoi bat len message de thong bao truong hop la to khai Bo sung
    ' Truong hop to khai la version 1.3.0
    If verToKhai = 0 Then
        With fpSpread1
        
            .EventEnabled(EventAllEvents) = False
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("B"), 17, qBoSung
            ' qBoSung : neu ky lap bo lon hon ky ke khai thi set =0 nguoc lai set 1
            If qBoSung = 0 And (TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue <> "104") Then
'                ' Huy bo thi quay lai man hinh quet to khai
'                msgRs = MessageBox("0084", msYesNoCancel, miQuestion, 1)
'               If msgRs = mrCancel Then
'                    If Not TAX_Utilities_Srv_New.Data(0) Is Nothing Then
'                        If Not objTaxBusiness Is Nothing Then
'                            objTaxBusiness.Prepared4 dNgayDauKy
'                            objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
'                        End If
'                        StartReceiveForm
'                    End If
'                    Exit Sub
'                ' Neu  ghi Thay the thi set lai trang thai cua to khai la 1 va ghi binh thuong
'                ElseIf msgRs = mrYes Then
'                    verToKhai = 1
'                ' Neu ghi Bo sung thi phai set lai tinh trang cua to khai la 2 va phai yeu cau quet phu luc KHBS
'                ElseIf msgRs = mrNo Then
'
'                        verToKhai = 2
'
'                End If
            End If
            .EventEnabled(EventAllEvents) = True
        End With
    ElseIf verToKhai = 2 Then
        ' Kiem tra neu la to khai TNCN moi thi ko phai quet KHBS
        ' IdToKhai = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
        If (TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue <> "104") Then
            With fpSpread1
                    verToKhai = 2
            End With
        End If
        ' 04-01-2011
        ' Kiem tra neu la to khai TNCN thi set lai  trang thai bo sung thanh thay the (2->1)
        If (TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "104") Then
            verToKhai = 1
        End If
    End If
    
    ' Lay lai ID cua to khai de biet la to khai co duoc gia han thue hay khong
    idToKhai = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    ' Truong hop to khai TNCN mau 02 va 07/TNCN hien tai tu 01->05/2009 cho phep gia han thue
    ' ngoai thoi gian nay phai thong bao khong duoc gia han
    ' Do voi to khai 02, 07/TNCN thang
    If Val(idToKhai) = 15 Or Val(idToKhai) = 36 Then
        Dim varTemp
        ' Lay thong tin ve gia han nop thue TNCN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 36
            varTemp = .Value
        End With
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
        If Val(TAX_Utilities_Srv_New.Year) <> 2009 Then
            If Val(varTemp) = 1 Then
                MessageBox "0090", msOKOnly, miInformation
                Exit Sub
            End If
        End If
    End If
    
    ' Do voi to khai 02/TNCN quy
    If Val(idToKhai) = 37 Then
        ' Lay thong tin ve gia han nop thue TNCN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 36
            varTemp = .Value
        End With
        ' Kiem tra xem co thuoc ky duoc gia han thue hay khong, neu khac 2009 thi thong bao khong duoc gia han nop thue
        If Val(TAX_Utilities_Srv_New.Year) <> 2009 Then
            If Val(varTemp) = 1 Then
                MessageBox "0090", msOKOnly, miInformation
                Exit Sub
            End If
        End If
    End If
    
    ' Truong hop to khai TNDN va quyet toan hien tai 2009 cho phep gia han thue
    ' ngoai thoi gian nay phai thong bao khong duoc gia han
    ' Do voi to khai 01A, 01B/TNDN thang
    If Val(idToKhai) = 11 Or Val(idToKhai) = 12 Then
        If Val(idToKhai) = 11 Then
            With fpSpread1
                .Sheet = 1
                .Col = .ColLetterToNumber("E")
                .Row = 37
                varTemp = .Value
            End With
        ElseIf Val(idToKhai) = 12 Then
            With fpSpread1
                .Sheet = 1
                .Col = .ColLetterToNumber("E")
                .Row = 38
                varTemp = .Value
            End With
        End If
    End If
    ' Do voi to khai 05/TNDN thang
    If Val(idToKhai) = 14 Then
        ' Lay thong tin ve gia han nop thue TNDN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 19
            varTemp = .Value
        End With
    End If
    ' Do voi to khai 03/TNDN thang
    If Val(idToKhai) = 3 Then
        ' Lay thong tin ve gia han nop thue TNDN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 37
            varTemp = .Value
        End With
    End If

    ' xu ly cho 2 to khai 08, 08A/TNCN
    If idToKhai = "74" Or idToKhai = "75" Then
        If verToKhai = 0 And isTKTonTai = True Then ' Trong truong hop to khai thay the nhung ke khai ko su dung KHBS de ke khai ma su dung chuc nang ke khai goc
                mResult = MessageBox("0047", msYesNo, miQuestion)
                If mResult = mrYes Then ' Neu dong y ghi la to khai thay the thi phai dat lai trang thai = 1
                    verToKhai = 1
                End If
        End If
    End If
    
    'Push data to ESB
    ExecuteSave
    '**********End push data to ESB*****************
    
    ' Clear data
    If Not objTaxBusiness Is Nothing Then
        'Get Params
        objTaxBusiness.GetParams strNgayNhanToKhai, strMaPhongXuLy 'strMaSoTep, strNgayNhanToKhai, strMaPhongXuLy
    End If
    StartReceiveForm
    
    Set xmlResultDLT = Nothing
    Set xmlResultNNT = Nothing
    'Set xmlResultNSD = Nothing
    '***************************
    blnSaveSuccess = True
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "cmdSave_Click", Err.Number, Err.Description
    MessageBox "0049", msOKOnly, miCriticalError
    
End Sub

Sub CallFinish(Optional blFinish As Boolean)

'On Error GoTo ErrHandle
'    Dim iSheet As Integer, iActiveSheet As Integer
'    Dim lActiveCol As Long, lActiveRow As Long
'    Dim lCol As Long, lRow As Long
'    Dim i As Integer
'    With fpSpread1
'        .Visible = False
'        .ReDraw = False
'        iActiveSheet = .ActiveSheet
'        lActiveCol = .ActiveCol
'        lActiveRow = .ActiveRow
'
'        For i = 1 To .SheetCount - 1
'            .ActiveSheet = i
'            .SetActiveCell 1, 1
'        Next
'        .ActiveSheet = iActiveSheet
'        .Sheet = iActiveSheet
'        .Col = lActiveCol
'        .Row = lActiveRow
'        .SetActiveCell lActiveCol, lActiveRow
'        .ReDraw = True
'        .Visible = True
'    End With
'    Exit Sub
'ErrHandle:
'    SaveErrorLog Me.Name, "CallFinish", Err.Number, Err.Description
    
    On Error GoTo ErrorHandle
        
    Dim iSheet As Integer, iActiveSheet As Integer
    Dim lActiveCol As Long, lActiveRow As Long
    Dim lCol As Long, lRow As Long
    Dim i As Integer
    With fpSpread1
        .Visible = False
        .ReDraw = False
        .EditMode = False
        iActiveSheet = .ActiveSheet
        lActiveCol = .ActiveCol
        lActiveRow = .ActiveRow
        
        
        For i = 1 To .SheetCount
            .ActiveSheet = i
            .Sheet = .ActiveSheet
            .Row = 1
            .Col = 1
            .Lock = False
            .SetActiveCell 1, 1
            .EditMode = True
        Next

        For i = 1 To .SheetCount
            .ActiveSheet = i
            .Sheet = .ActiveSheet
            .Row = 1
            .Col = 1
            .Lock = True
            .EditMode = False
        Next
        .ActiveSheet = iActiveSheet
        .Sheet = iActiveSheet
        .Col = lActiveCol
        .Row = lActiveRow
        .EditMode = True
        .SetActiveCell lActiveCol, lActiveRow
        .ReDraw = True
        .Visible = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "CallFinish", Err.Number, Err.Description
End Sub

Private Sub cmdViewNow_Click()
    On Error GoTo ErrHandle
    Dim intBarcodeCount As Integer, intBarcodeNo As Integer, intBarcodeIncre As Integer
    Dim strPrefix As String, strBarcodeCount As String, strBarcode As String
    Dim i, j, t, counter As Integer
    Dim chkToKhai As Boolean
    
    Dim strLoaiTK As String
    ' Phien ban 1.3.1, Danh dau ma vach cua to khai Bo sung la TT (verToKhai = 2)
'    If verToKhai = 2 Then
'        MessageBox "0085", msOKOnly, miInformation
'        cmdViewNow.Enabled = False
'        Exit Sub
'    End If
    'strBarcode = TAX_Utilities_Srv_New.Convert(strBarcode, UNICODE, TCVN)
    If verToKhai = 2 Then
        strLoaiTK = "bs"
    Else
        strLoaiTK = "aa"
    End If
    
    For i = 1 To UBound(arrBCBuffer)
        If arrBCBuffer(i) <> vbNullString Then
            intBarcodeCount = intBarcodeCount + 1
        End If
    Next
    ' Khai bao lai mang BCNew luu tat ca cac phan tu khac null trong BCBuffer
    ReDim Preserve arrBCNew(intBarcodeCount)
    ' Clear mang bawcode arrStrElements hien dang co tren Ram de bat dau quet lai
    ReDim arrStrElements(0)
    ' Bat dau se la so thu tu cua chuoi ma vach = 1 sau do se tang dan len
    intBarcodeIncre = 1
    ' Ban dau dat dieu kien la khong co to khai, sau do kiem tra neu quet ma co to khai thi chkToKhai = true
    chkToKhai = False
    For j = 1 To UBound(arrBCBuffer)
        If arrBCBuffer(j) <> vbNullString Then
            counter = counter + 1
            strPrefix = Left$(arrBCBuffer(j), 36)
            strBarcodeCount = Right$(strPrefix, 6)
            strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 6)
            
            strBarcode = Mid$(arrBCBuffer(j), 37)
            intBarcodeNo = CInt(Val(Left$(strBarcodeCount, 3)))
            If intBarcodeNo = 1 Then
                strBarcodeCount = vbNullString
                strBarcodeCount = "001" & Right("000" & intBarcodeCount, 3)
                chkToKhai = True
            Else
                strBarcodeCount = vbNullString
                intBarcodeIncre = intBarcodeIncre + 1
                strBarcodeCount = Right("000" & intBarcodeIncre, 3) & Right("000" & intBarcodeCount, 3)
            End If
            arrBCNew(counter) = strLoaiTK & strPrefix & strBarcodeCount & strBarcode
        End If
    Next
    ' Neu chua quet to khai ma co yeu cau hien thi thi thong bao phai quet to khai
    If chkToKhai = False Then
        DisplayMessage "0083", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    For t = 1 To UBound(arrBCNew)
        Barcode_Scaned arrBCNew(t)
    Next
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "arrBCBuffer Error!", Err.Number, Err.Description
    
End Sub

Private Sub Command1_Click()
   Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String, str9 As String, str10 As String
   Dim str11 As String, str12 As String, str13 As String, str14 As String, str15 As String, str16 As String, str17 As String, str18 As String, str19 As String, str20 As String
   Dim str21 As String, str22 As String, str23 As String, str24 As String, str25 As String, str26 As String, str27 As String, str28 As String, str29 As String, str30 As String
   Dim str31 As String, str32 As String, str33 As String, str34 As String, str35 As String, str36 As String, str37 As String, str38 As String, str39 As String, str40 As String
   Dim str41 As String, str42 As String, str43 As String, str44 As String, str45 As String, str46 As String, str47 As String, str48 As String, str49 As String, str50 As String
   Dim str51 As String, str52 As String, str53 As String

'04/GTGT
'str2 = "aa320712222222222   08201300200300100201/0101/01/1900<S01><S></S><S>150000000~7500000~15000000</S><S>10000000~500000~1000000~01~20000000~1000000~2000000~06~30000000~1500000~3000000~03~40000000~2000000~4000000~04~50000000~2500000~5000000~05</S><S>5~5.5~6~01~7~7.5~8~06~9~9.5~10~03~11~11.5~12~04~13~13.5~14~05</S><S>15500000~812500~1700000</S><S>"
'Barcode_Scaned str2
'str2 = "aa320712222222222   082013002003002002500000~27500~60000~01~1400000~75000~160000~06~2700000~142500~300000~03~4400000~230000~480000~04~6500000~337500~700000~05</S><S>0~40625~170000</S><S>0~1375~6000~01~0~3750~16000~06~0~7125~30000~03~0~11500~48000~04~0~16875~70000~05</S><S>172500000~210625</S><S>CMC~CMCer~123456789~12/09/2013~1~~~0</S></S01>"
'Barcode_Scaned str2

'' 04/GTGT - co bo sung KHBS
'str2 = "bs320712222222222   08201300400500100301/0101/01/1900<S01><S></S><S>177656545~10053434~11111111</S><S>10000000~500000~1000000~01~20000000~1000000~2000000~06~30000000~1500000~3000000~03~67656545~4553434~111111~04~50000000~2500000~5000000~05</S><S>5~5.5~6~01~7~7.5~8~06~54.78~9.5~10~03~11~24.67~70~04~13~13.5~14~05</S><S>32276220~1705832~1297778</S>"
'Barcode_Scaned str2
'str2 = "bs320712222222222   082013004005002003<S>500000~27500~60000~01~1400000~75000~160000~06~16434000~142500~300000~03~7442220~1123332~77778~04~6500000~337500~700000~05</S><S>0~85292~129778</S><S>0~1375~6000~01~0~3750~16000~06~0~7125~30000~03~0~56167~7778~04~0~16875~70000~05</S><S>198821090~215070</S><S>CMC~CMCer~123456789~12/09/2013~~1~1~0</S></S01>"
'Barcode_Scaned str2
'str2 = "bs320712222222222   082013004005003003<SKHBS><S>Hµng ho∏, dﬁch vÙ chﬁu thu’ su t 5%~31~40625~85292~44667</S><S>Hµng ho∏, dﬁch vÙ chﬁu thu’ su t 10%~32~170000~129778~-40222</S><S>26/09/2013~6~13~ßi“u chÿnh t®ng gi∂m Æ” test ch≠¨ng tr◊nh.~4445</S></SKHBS>"
'Barcode_Scaned str2

''01/XS
'str2 = "aa320482222222222   09201300200200100101/0101/01/2010<S01><S></S><S>5000000~600000~100000</S><S>dfgdfhj~rtrt~dfcgfg~12/09/2013~1~~</S></S01>"
'Barcode_Scaned str2
'
''02/XS
'str2 = "aa320432222222222   00201200300300100201/0101/01/2009<S02><S></S><S>3~3800000~3~3800000~182000</S><S>CMCer~12/09/2013~CMC~123456789~1~</S></S02>"
'Barcode_Scaned str2
'
' 02/TAIN
''str2 = "aa317773100177415   00201300400400100201/0114/06/2006<S01><S>0102845045</S><S>010102~Kg~16535~0~0~30~10~010103~Kg~5847~8~11~0~888~010104~Kg~2222~70~11~0~300~010207~T n~3000~1000~15~0~1000~010203~T n~5"
''Barcode_Scaned str2
''str2 = "aa317773100177415   0020130040040020020000~0~0~100~300</S><S>010208~Kg~564565~56~10~0~765.987~010210~Kg~6343~0~0~49~845~010208~T n~100~50~10~0~10</S><S>outh~12/09/2013~rty~red~1~</S></S01>"
''Barcode_Scaned str2

'' 02/TAIN - co bo sung KHBS --3100177415 -- 0102845045
'str2 = "aa317773100177415   00201200600800100301/0114/06/2006<S01><S>0102845045</S><S>010104~Kg~123~48888~11~0~0~050101~T n~65342.895~543.768~7~0~654.987~010207~Kg~27646.456~0~0~876.456~888.320</S><S>010"
'Barcode_Scaned str2
'str2 = "aa317773100177415   002012006008002003203~Kg~6776.893~4888~15~0~0~010210~Kg~7646.876~0~0~68544~7777~010104~Kg~65356.897~457.987~11~0~765.934</S><S>ghghg~12/09/2013~gg~ghgh~~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa317773100177415   002012006008003003<SKHBS><S>Thu’ tµi nguy™n ph∏t sinh ph∂i nÈp trong k˙~10~661455~559778340~559116885</S><S>~~0~0~0</S><S>26/09/2013~178~49761403~NOI DUNG DINH KEM ABC~559116885</S></SKHBS>"
'Barcode_Scaned str2

''04-GTGT-BS
'str2 = "bs320713100177415   08201300400400100301/0101/01/1900<S01><S>0102845045</S><S>60000000~2000000~4000000</S><S>10000000~2000000~3000000~01~50000000~0~1000000~03</S><S>0~5~10~01~0~5~10~03</S><S>0~100000~40000"
'Barcode_Scaned str2
'str2 = "bs320713100177415   0820130040040020030</S><S>0~100000~300000~01~0~0~100000~03</S><S>0~5000~40000</S><S>0~5000~30000~01~0~0~10000~03</S><S>66000000~45000</S><S>dasdsad~~21sdasd~24/09/2013~~1~1~0</S></S01>"
'Barcode_Scaned str2
'str2 = "bs320713100177415   082013004004003003<SKHBS><S>~~0~0~0</S><S>Hµng ho∏, dﬁch vÙ chﬁu thu’ su t 5%~31~27500~5000~-22500~Hµng ho∏, dﬁch vÙ chﬁu thu’ su t 10%~32~551000~40000~-511000</S><S>24/09/2013~4~0~~-533500</S></SKHBS>"
'Barcode_Scaned str2

''04/GTGT chinh thuc
'
'str2 = "aa317710201027770   02201300100100100201/0101/01/1900<S01><S></S><S>70000000~11000000~55100000</S><S>10000000~2000000~3000000~01~50000000~0~1000000~03~1000000~2000000~50000000~02~4000000~5000000~100000~04~5000000~2000000~1000000~02</S><S>0~5~10~01~0~5~10~03~0~5~10~02~0~5~10~04~0~5~10~02</S><S>0~550000~5510"
'Barcode_Scaned str2
'str2 = "aa317710201027770   022013001001002002000</S><S>0~100000~300000~01~0~0~100000~03~0~100000~5000000~02~0~250000~10000~04~0~100000~100000~02</S><S>0~27500~551000</S><S>0~5000~30000~01~0~0~10000~03~0~5000~500000~02~0~12500~1000~04~0~5000~10000~02</S><S>136100000~578500</S><S>dasdsad~~21sdasd~24/09/2013~1~1~1~1</S></S01>"
'Barcode_Scaned str2
''04/GTGT bo sung
'
'str2 = "bs317713100177415   08201300400400100301/0101/01/1900<S01><S>0102845045</S><S>60000000~2000000~4000000</S><S>10000000~2000000~3000000~01~50000000~0~1000000~03</S><S>0~5~10~01~0~5~10~03</S><S>0~100000~40000"
'Barcode_Scaned str2
'str2 = "bs317713100177415   0820130040040020030</S><S>0~100000~300000~01~0~0~100000~03</S><S>0~5000~40000</S><S>0~5000~30000~01~0~0~10000~03</S><S>66000000~45000</S><S>dasdsad~~21sdasd~24/09/2013~~1~1~0</S></S01>"
'Barcode_Scaned str2
'str2 = "bs317713100177415   082013004004003003<SKHBS><S>~~0~0~0</S><S>Hµng ho∏, dﬁch vÙ chﬁu thu’ su t 5%~31~27500~5000~-22500~Hµng ho∏, dﬁch vÙ chﬁu thu’ su t 10%~32~551000~40000~-511000</S><S>24/09/2013~4~0~~-533500</S></SKHBS>"
'Barcode_Scaned str2
''01/KK-XS theo thang chinh thuc
'
'str2 = "aa31748040010191900808201300100100100101/0101/01/2010<S01><S>0102845045</S><S>2000000~500000~300000</S><S>sdfsf~dasdsad~21sdasd~24/09/2013~1~~</S></S01>"
'Barcode_Scaned str2
'str2 = "aa31748040010191900808201300100100100101/0101/01/2010<S01><S>0102845045</S><S>2000000~500000~300000</S><S>sdfsf~dasdsad~21sdasd~24/09/2013~1~~</S></S01>"
'Barcode_Scaned str2
''01/KK-XS  theo thang bo sung
'
'str2 = "bs31748040010191900808201300200200100101/0101/01/2010<S01><S>0102845045</S><S>6000000~800000~80000</S><S>sdfsf~dasdsad~21sdasd~24/09/2013~~1~1</S></S01>"
'Barcode_Scaned str2
'
'
'01/KK-XS theo quy chinh thuc

'str2 = "aa31749000000001700802201300100100100101/0101/01/2010<S01><S>0102845045</S><S>6000000~3000000~400000</S><S>~dasdsad~21sdasd~24/09/2013~1~~</S></S01>"
'Barcode_Scaned str2
''01/KK-XS theo quy bo sung
'
'str2 = "bs31749040010191900802201300100200100101/0101/01/2010<S01><S></S><S>6000000~4000000~800000</S><S>~dasdsad~21sdasd~24/09/2013~~1~1</S></S01>"
'Barcode_Scaned str2
''02/KK-XS chinh thuc
'
'str2 = "aa31743040010191900800201200200200100201/0101/01/2009<S02><S>0102845045</S><S>4~162449539~3~152449539~891182</S><S>abc~24/09/2013~dasdsad~21sdasd~1~</S></S02>"
'Barcode_Scaned str2
''02/KK-XS bo sung
'
'str2 = "aa317433100177415   00201200600600100301/0101/01/2009<S02><S>0102845045</S><S>4~393683773~3~383683773~891182</S><S>abc~24/09/2013~dasdsad~21sdasd~~1</S></S02>"
'Barcode_Scaned str2

'Dim xmlDoc As New MSXML.DOMDocument
'xmlDoc.Load "C:\tempxml\ToKhai.xml"
'PushDataToESB xmlDoc.xml

'TEST GIAI DOAN 2
'' To Khai QD 15 BCTC
'str2 = "aa999693100177415   00201200500500100801/0123/06/2006<S01><S>~41400~0~~3000~0~V.01~1000~0~~2000~0~V.02~3000~0~~1000~0~~2000~0~~33000~0~~3000~0~~4000~0~~5000~0~~6000~0~V.03~7000~0~~8000~0~~800~0~V.04~500~0~~300~0~~1600~0~~600~0~~700~0~V.05~200~0~~100~0~~80000~0~~15000~0~~5000~0~~4000~0~V.06~3000~0~V.07~2000~0~~1000~0~~33000~0~V.08~13000~0~~6000~"
'Barcode_Scaned str2
'str2 = "aa999693100177415   0020120050050020080~~7000~0~V.09~7000~0~~2000~0~~5000~0~V.10~11000~0~~8000~0~~3000~0~V.11~2000~0~V.12~7000~0~~1000~0~~6000~0~~17000~0~~5000~0~~7000~0~V.13~3000~0~~2000~0~~8000~0~V.14~1000~0~V.21~4000~0~~3000~0~~121400~0~~59800~0~~55000~0~V.15~7000~0~~6000~0~~5000~0~V.16~4000~0~~8000~0~V.17~9000~0~~2000~0~~1000~0~V.18~3000~0~"
'Barcode_Scaned str2
'str2 = "aa999693100177415   002012005005003008~6000~0~~4000~0~~4800~0~~400~0~V.19~700~0~~600~0~V.20~800~0~V.21~900~0~~400~0~~300~0~~500~0~~200~0~~61600~0~V.22~52600~0~~1000~0~~2000~0~~3000~0~~4000~0~~5000~0~~6000~0~~7000~0~~8000~0~~9000~0~~1600~0~~4000~0~~2000~0~~9000~0~V.23~3000~0~~6000~0~~121400~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~15/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999693100177415   002012005005004008<S01-1><S>VI.25~4000~0~~3000~0~~1000~0~VI.27~6000~0~~-5000~0~VI.26~9000~0~VI.28~7000~0~~6000~0~~5000~0~~4000~0~~-12000~0~~3000~0~~2000~0~~1000~0~~-11000~0~VI.30~4000~0~VI.30~5000~0~~-20000~0~~3000~0~~15/10/2013</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999693100177415   002012005005005008<S01-2><S>~5000~0~~4000~0~~3000~0~~2000~0~~1000~0~~6000~0~~7000~0~~28000~0~~3000~0~~5000~0~~7000~0~~3000~0~~2000~0~~1000~0~~5000~0~"
'Barcode_Scaned str2
'str2 = "aa999693100177415   002012005005006008~26000~0~~5000~0~~4000~0~~3000~0~~6000~0~~7000~0~~8000~0~~33000~0~~87000~0~~4500~0~~6000~0~VII.34~97500~0~~15/10/2013</S></S01-2>"
'Barcode_Scaned str2
'str2 = "aa999693100177415   002012005005007008<S01-3><S>~6000~0~~5000~0~~4000~0~~3000~0~~2000~0~~1000~0~~21000~0~~6000~0~~4000~0~~5000~0~~9000~0~~8000~0~~2000~0~~1000~0~~5000~0~~61000~0~~7000~0~~5000~0~~8000"
'Barcode_Scaned str2
'str2 = "aa999693100177415   002012005005008008~0~~3000~0~~1000~0~~6000~0~~4000~0~~34000~0~~9000~0~~4000~0~~2000~0~~3000~0~~5000~0~~1000~0~~24000~0~~119000~0~~6000~0~~3000~0~~128000~0~~15/10/2013</S></S01-3>"
'Barcode_Scaned str2

'' To Khai BC 36
'str2 = "aa999683100177415   032013012012002004120~99~0~H„a Æ¨n b∏n hµng (dµnh cho tÊ ch¯c, c∏ nh©n trong khu phi thu’ quan)~07KPTQ4/001~MN/23E~111~~~0000040~0000150~~~0~0~0~~0~~0~~0000040~0000150~111~0~Phi’u xu t kho ki™m vÀn chuy”n hµng h„a nÈi bÈ~03XKNB5/001~KT/34T~101~0000050~0000150~~~~~0~0~0~~0~~0~~0000050~0000150~101~0~Phi’u xu t kho gˆi b∏n hµng Æπi l˝~04HGDL6/001~BD/24T~121~~~0000080~0000200~~~0~0~0~~0~~0~~0000080~0000200~1"
'Barcode_Scaned str2
'str2 = "aa999683100177415   03201301201200100401/0101/01/2009<S01><S>~~01/07/2013~30/09/2013</S><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/001~AB/12T~11~0000000~0000010~~~0000000~0000008~9~6~1~0~1~1~1~2~0000009~0000010~2~0~H„a Æ¨n b∏n hµng~02GTTT2/001~CD/23T~81~~~0000020~0000100~0000020~0000041~22~20~0~~1~23~1~22~0000042~0000100~59~0~H„a Æ¨n xu t kh»u~06HDXK3/001~HN/13P~111~0000010~0000120~~~0000010~0000021~12~10~0~~1~11~1~14~0000022~0000"
'Barcode_Scaned str2
'str2 = "aa999683100177415   032013012012004004 pp tr˘c ti’p~02THDB3/001~PL/78T~199~0000048~0000246~~~~~0~0~0~~0~~0~~0000048~0000246~199~0~Tem vÀn t∂i Æ≠Íng bÈ theo pp tr˘c ti’p~02TEDB4/001~GH/56P~197~~~0000079~0000275~~~0~0~0~~0~~0~~0000079~0000275~197~0~V– vÀn t∂i Æ≠Íng bÈ theo pp tr˘c ti’p~02VEDB6/001~LH/28T~350~0000029~0000189~0000190~0000378~~~0~0~0~~0~~0~~0000029~0000378~350~0</S><S>Ph≠¨ng Lan~Lan H≠¨ng~15/10/2013~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999683100177415   03201301201200300421~0~Tem vÀn t∂i Æ≠Íng bÈ theo pp kh u trı~01TEDB8/001~BM/27T~59~0000069~0000127~~~~~0~0~0~~0~~0~~0000069~0000127~59~0~V– vÀn t∂i Æ≠Íng bÈ theo pp kh u trı~01VEDB9/001~AD/34P~201~~~0000069~0000269~~~0~0~0~~0~~0~~0000069~0000269~201~0~ThŒ vÀn t∂i Æ≠Íng bÈ theo pp kh u trı~01THDB2/001~UH/23T~440~0000011~0000300~0000301~0000450~~~0~0~0~~0~~0~~0000011~0000450~440~0~ThŒ vÀn t∂i Æ≠Íng bÈ theo"
'Barcode_Scaned str2

'-------------*************************------------------------

' To Khai 02/KK-TNCN theo Thang - To khai chinh thuc
'str2 = "aa999153100177415   09201300100100100101/0101/01/2010<S01><S>1000724808</S><S>10~10~10~0~50000000~0~0~20000000~0~0~10000000~0~0</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~1~~</S></S01>"
'Barcode_Scaned str2
'' To Khai 02/KK-TNCN theo Thang - To khai bo sung
'str2 = "bs999153100177415   09201300200200100101/0101/01/2010<S01><S>1000724808</S><S>100~100~50~50~25000000~0~25000000~10000000~0~10000000~5000000~0~5000000</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~~1~1</S></S01>"
'Barcode_Scaned str2
'' To Khai 02/KK-TNCN theo Quy - To khai chinh thuc
'str2 = "aa999163100177415   02201300100100100101/0101/01/2010<S01><S>1000724808</S><S>20~20~20~0~100000000~100000000~0~50000000~50000000~0~5000000~5000000~0</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~1~~</S></S01>"
'Barcode_Scaned str2
'' To Khai 02/KK-TNCN theo Quy - To khai bo sung
'str2 = "bs999163100177415   02201300200200100101/0101/01/2010<S01><S>1000724808</S><S>40~40~40~0~200000000~200000000~0~100000000~100000000~0~10000000~10000000~0</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~~1~1</S></S01>"
'Barcode_Scaned str2

'' To Khai 03/KK-TNCN theo Thang - To khai chinh thuc
'str2 = "aa999503100177415   09201300100100100201/0101/01/2010<S01><S>1000724808</S><S>10000000~500000~20000000~20000~30000000~1500000~5000000~500"
'Barcode_Scaned str2
'str2 = "aa999503100177415   092013001001002002000~500000000~20000000~100000000~50000000</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~1~~</S></S01>"
'Barcode_Scaned str2
'' To Khai 03/KK-TNCN theo Thang - To khai bo sung
'str2 = "bs999503100177415   09201300200200100201/0101/01/2010<S01><S>1000724808</S><S>5000000~250000~10000000~10000~15000000~750000~2500000~2500"
'Barcode_Scaned str2
'str2 = "bs999503100177415   09201300200200200200~250000000~10000000~50000000~25000000</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~~1~1</S></S01>"
'Barcode_Scaned str2
'' To Khai 03/KK-TNCN theo Quy - to khai chinh thuc
'str2 = "aa999513100177415   02201300100100100101/0101/01/2010<S01><S>1000724808</S><S>20000000~1000000~50000000~50000~100000000~5000000~200000000~20000000~100000000~4000000~200000000~100000000</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~1~~</S></S01>"
'Barcode_Scaned str2
'' To Khai 03/KK-TNCN theo Quy - To khai bo sung
'str2 = "bs999513100177415   02201300200200100101/0101/01/2010<S01><S>1000724808</S><S>10000000~500000~25000000~25000~50000000~2500000~100000000~10000000~50000000~2000000~100000000~50000000</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~~1~1</S></S01>"
'Barcode_Scaned str2

'' To Khai 07/KK-TNCN - To khai chinh thuc
'str2 = "bs999512300790401   02201300200200100101/0101/01/2010<S01><S>2100462770</S><S>10000000~500000~25000000~25000~50000000~2500000~100000000~10000000~50000000~2000000~100000000~50000000</S><S>Lan H≠¨ng~16/10/2013~Ph≠¨ng Anh~KTV~~1~1</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999982222222222   10201300200200100201/0114/06/2006<S01><S>6868686868</S><S>0~~~~1~0~0~1~~22/10/2013~435435</S><S>43543~543543~23667492849~43~10177021925~543543~1017701648964~45343534</S><S>dfds~dfg~fdgfd~fdgfdg</S><S>UEFW~32432~dfgfd~22/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999982222222222   102013002002002002<S01-1><S>1017701648964</S><S>2222222222~dfgfdg~100~1017701648964~</S><S>100~1017701648964</S></S01-1>"
''Barcode_Scaned str2
'str2 = "aa999922222222222   10201300100100100101/0114/06/2006<S01><S>6868686868</S><S>0~~~~1~0~0~1~~24/10/2013~</S><S>0~0~0~0~0~0~0</S><S>UEFW~32432~~24/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999922222222222   09201300000000100101/0114/06/2006<S01><S>6868686868</S><S>0~~~~0~0~1~1~~~</S><S>0~0~0~0~0~0~0</S><S>UEFW~32432~~24/10/2013</S></S01>"
'Barcode_Scaned str2
'' 02/KK-TNCN - Quy
'str2 = "aa999162300790401   03201300100100100101/0101/01/2010<S01><S>2100462770</S><S>1000~800~300~200~900000000~700000000~200000000~100000000~20000000~30000000~20000000~2000000~6000000</S><S>Nguyen Van A~21/10/2013~~~1~~</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999922222222222   09201300100100100201/0114/06/2006<S01><S>6868686868</S><S>5~~x~~0~0~1~1~~~</S><S>435~435~435~189225~43~81367~435435</S><S>UEFW~32432~~25/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999922222222222   092013001001002002<S01-1><S>81367</S><S>6868686868~dsfsf~100~81367~</S><S>81367</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999922222222222   09201300100100100201/0114/06/2006<S01><S>6868686868</S><S>5~x~~~0~0~1~1~~~</S><S>435~435~435~189225~43~81367~435435</S><S>UEFW~32432~~25/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999922222222222   092013001001002002<S01-1><S>81367</S><S>6868686868~dsfsf~100~81367~</S><S>81367</S></S01-1>"
'Barcode_Scaned str2

''01A_TNDN_DK
'str2 = "aa999982300790384   10201300200200100101/0114/06/2006<S01><S></S><S>1~~x~13/10/2013~1~0~0~1~~24/10/2013~lo 34</S><S>275.89~34100~9407849~25.78~2425343~26475~2398868~25.34</S><S>tk002~ngan hang AB</S><S>~~Mai~24/10/2013</S></S01><S01-1><S>2398868</S><S>3600247325~nha thau B~45.78~1098202~ghi chu 1~0102030405~nha thau C~28.39~681039~ghi chu 2</S><S>74.17~1779241</S></S01-1>"
'Barcode_Scaned str2

''01B_TNDN_DK
'str2 = "aa999992300790401   03201300200300100101/0114/06/2006<S01><S>2100462770</S><S>1~~</S><S>459.17~500000~229585000~47691~229537309~45.68~104852643~729851~104122792~26.34</S><S>tk001~ngan hang acb</S><S>~~hanh~24/10/2013</S></S01><S01-1><S>104122792</S><S>0010011000~nha thau A~25.45~26499251~ghi chu 1~0102030405~nha thau B~56.55~58881439~ghi chu 2</S><S>85380690</S></S01-1>"
'Barcode_Scaned str2

''01_TD_GTGT
'str2 = "aa999941400633697   09201300200200100101/0114/06/2006<S01><S></S><S>200000~12.58~2516000~10~251600~30000~221600</S><S>cu chuoi~chuoi cu~abcdef~29/10/2013~1~~</S></S01><S01_1><S>hjjkkjjkjk~0102030405~300000~10000~290000</S><S>300000~10000~290000</S></S01_1><S01_2><S>1~nhfdyjb~0010011000~CÙc Thu’ Tÿnh H-ng Y™n~14.78~100000~10900</S><S>100000</S></S01_2>"
'Barcode_Scaned str2

'Ra soat to khai
'01-GTGT thang - lan dau - chon tat ca cac phu luc
'str2 = "aa999012300790401   09201300600600101701/0114/06/2006<S01><S>2100462770</S><S>~0~89000000~6550000~5090277~4000000~65500000~2375000~30000000~23500000~1175000~12000000~1200000~"
'Barcode_Scaned str2
'str2 = "aa999012300790401   09201300600600201769500000~2375000~-2715277~0~0~6200000000~0~0~0~0~2715277~0~0~2715277</S><S>Nguy‘n S¸ HÔng~DEV1234~~30/10/2013~1~~~1701~x~02~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006003017<S01_1><S>0355~065~11/09/2013~lan~0102030405~gmao~4000000~0~ghi</S><S>02324~056~11/09/2013~su~0102030405~m›a~30000000~0~okie</S><S>01234~012~10/03/2013~Nguy‘n V®n H∂i~0102030405~Æ≠Íng~23500000~1175000~o"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006004017ka</S><S>032134~0234~09/09/2013~Nguy‘n V®n S˘~0102030405~Ng´~12000000~1200000~okie</S><S>05436~0345~08/08/2013~L™ Vi÷t C≠Íng~0102030405~Khoai~24000000~0~yes</S><S>69500000~2375000~65500000</S></S01_1>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006005017<S01_2><S>0178~09887~01/01/2013~Nguy‘n V®n A~0102030405~LÛa~12000000~0~0~</S><S>0234~0987~02/02/2013~Nguy‘n V®n B~0102030405~Ng´~23000000~5~1150000~</S><S>0678~0923~03/03/2013~Ngu"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006006017y‘n V®n C~0102030405~Khoai~54000000~10~5400000~</S><S>~~~~~~0~0~0~</S><S>02345~890~04/04/2013~Nguy‘n V®n D~0102030405~Sæn~90000000~10~9000000~</S><S>89000000~6550000</S></S01_2>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006007017<S01_3><S>Camry~Chi’c~100~100~1200000000~~Land Cruiser~Chi’c~200~200~1900000000~</S><S>Honda Lead~Chi’c~1020~1020~35000000~~Honda SH~Chi’c~500~500~500000000~</S></S01_3>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006008017<S01_4A><S>6620000~70000~1150000~5400000~69500000~65500000~94.24~5400000~5088960</S></S01_4A>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006009017<S01_4B><S>2009~936000~800000~60000~76000~1000000~300000~30~4578~1373~56~1317</S></S01_4B>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006010017<S01_5><S>05/05/2013~1200000000~10101~07/07/2013~5000000000~10101</S></S01_5>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006011017<S01_6><S>CMC Soft~0102030405~20000000000~15000000000~35000000"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006012017000~500000000~0~10101~10100</S><S>0~500000000~0~0</S></S01_6>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006013017<S01_6B><S>CMC Soft HCM~0102030405~10000000000~200000000~0~10101</S><S>0~10000000~0~0</S></S01_6B>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006014017<S01_7><S>Ha Noi~0102030405~1200000000~20000000000~2120"
'Barcode_Scaned str2
'str2 = "aa999012300790401   0920130060060150170000000~412000000~0</S><S>0~210000000~0~0</S></S01_7>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006016017<S01_1_TD><S>Ialy~10000000000~10000000000~0~Na Hang~30000000000~40000000000~10000000000</S></S01_1_TD>"
'Barcode_Scaned str2
'str2 = "aa999012300790401   092013006006017017<S01_2_TD><S>Thac Ba~10~2000000000~10101~~Hoa Binh~5~20000000000~10101~</S></S01_2_TD>"
'Barcode_Scaned str2

'01-GTGT- khong phu luc
'str2 = "aa999012300790401   08201300400400100201/0114/06/2006<S01><S>2100462770</S><S>~10000~30000~70000~2000~300000~60459~11994~3456~24567~4663~32436~7331~360459~1"
'Barcode_Scaned str2
'str2 = "aa999012300790401   0820130040040020021994~9994~7000~3087~2567~3907~2567~85~1255~0~0~6756~0</S><S>Nguy‘n S¸ HÔng~DEV1234~~30/10/2013~1~~~1701~~~0</S></S01>"
'Barcode_Scaned str2

'BC21-AC
str2 = "aa999662300790401   03201300200200100101/0101/01/2010<S01><S>05/12/2013~14~24</S><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/012~AB/12T~0000010~0000100~91~6;9;12~05~0~H„a Æ¨n xu t kh»u~06HDXK7/009~MN/23T~0000150~0000250~101~34~02~0~Phi’u xu t kho ki™m vÀn chuy”n hµng h„a nÈi bÈ~03XKNB8/005~HD/13T~0000028~0000149~122~5;11~04~0</S><S>chay hoa don~Hoµng~05/12/2013</S></S01>"
Barcode_Scaned str2

''QD 15 BCTC
'str2 = "aa999692300790433   00201200000000100801/0123/06/2006<S01><S>~73920~48000~~21000~3000~V.01~1000~1000~~20000~2000~V.02~3000~3000~~1000~1000~~2000~2000~~33000~21000~~3000~1000~~4000~2000~~5000~3000~~6000~4000~V.03~7000~5000~~8000~6000~~3000~3000~V.04~1000~1000~~2000~2000~~13920~18000~~3000~3000~~5000~4000~V.05~3450~5000~~2470~6000~~95780~88000~~19940~15000~~6920~4000~~3420~5000~V.06~5600~3000~V.07~1300~2000~~2700~1000~~41000~28000~V.08~9400~3000~~6000~10"
'Barcode_Scaned str2
'str2 = "aa999692300790433   00201200000000200800~~3400~2000~V.09~7400~7000~~2900~3000~~4500~4000~V.10~14800~11000~~6700~5000~~8100~6000~V.11~9400~7000~V.12~12000~17000~~9000~8000~~3000~9000~~11000~10000~~2000~1000~~1000~2000~V.13~3000~3000~~5000~4000~~11840~18000~V.14~5630~5000~V.21~3210~6000~~3000~7000~~169700~136000~~97680~94000~~53030~52000~V.15~1000~9000~~2000~8000~~3000~7000~V.16~4000~6000~~5000~5000~V.17~6000~4000~~7000~3000~~8000~2000~V.18~9000~1000~~34"
'Barcode_Scaned str2
'str2 = "aa999692300790433   00201200000000300850~4000~~4580~3000~~44650~42000~~3980~2000~V.19~8430~6000~~2180~4000~V.20~1560~7000~V.21~4000~5000~~5300~8000~~7500~3000~~3600~2000~~8100~5000~~72020~42000~V.22~59565~39000~~4700~3000~~2000~6000~~1000~4000~~4000~5000~~2000~3000~~6000~2000~~9000~4000~~3000~3000~~5000~2000~~6000~3000~~7000~1000~~9865~3000~~12455~3000~V.23~9455~2000~~3000~1000~~169700~136000~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~Hoa Linh~22/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999692300790433   002012000000004008<S01-1><S>VI.25~30000~20000~~1000~2000~~29000~18000~VI.27~2000~3000~~27000~15000~VI.26~1000~1000~VI.28~4000~2000~~5000~3000~~6000~4000~~2000~1000~~16000~9000~~5000~2000~~3000~1000~~2000~1000~~18000~10000~VI.30~2100~300~VI.30~3800~100~~12100~9600~~2000~1000~Hoa Linh~22/10/2013</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999692300790433   002012000000005008<S01-2><S>~1000~7000~~2000~5000~~3000~3000~~4000~2000~~5000~4000~~6000~6000~~7000~1000~~28000~28000~~1000~4000~~2000~3000~~3000~2000~~4000~8000~~5000~5000~~7000~900~~2000~2000~~"
'Barcode_Scaned str2
'str2 = "aa999692300790433   00201200000000600824000~24900~~3000~1000~~2000~2000~~5000~3000~~6000~4000~~7000~5000~~9000~6000~~32000~21000~~84000~73900~~8000~5000~~1000~2000~VII.34~93000~80900~Hoa Linh~22/10/2013</S></S01-2>"
'Barcode_Scaned str2
'str2 = "aa999692300790433   002012000000007008<S01-3><S>~12000~20000~~1000~2000~~2000~4000~~3000~5000~~4000~3000~~5000~1000~~27000~35000~~2000~4000~~7000~6000~~3000~5000~~2000~3000~~1000~9000~~8000~6000~~3000~2000~~2000~1000~~55000~71000~~1000~4000~~3000~3000~~4000~200"
'Barcode_Scaned str2
'str2 = "aa999692300790433   0020120000000080080~~7000~5000~~2000~7000~~4000~9000~~8000~8000~~29000~38000~~4000~1000~~2000~5000~~6000~9000~~9000~4000~~2000~7000~~1000~4000~~24000~30000~~108000~139000~~7000~5000~~2000~3000~~117000~147000~Hoa Linh~22/10/2013</S></S01-3>"
'Barcode_Scaned str2

'02/KK-TNCN Thang - Lan dau
'str2 = "aa999152300790401   10201300100100100101/0101/01/2010<S01><S>2100462770</S><S>100~50~2891~2376~6745~845~129~3289~2367~178~123~237~36</S><S>Hoµng~19/11/2013~Huy“n Linh~KTV~1~~</S></S01>"
'Barcode_Scaned str2
''02/KK-TNCN Quy - Lan dau
'str2 = "aa999162300790401   03201300100100100101/0101/01/2010<S01><S>2100462770</S><S>569~128~2367~1876~3981~3768~138~3278~17665~389~2345~1767~78</S><S>Hoµng~19/11/2013~Huy“n Linh~KTV~1~~</S></S01>"
'Barcode_Scaned str2


'str2 = "aa999192300790401   00201200000000101201/0114/09/2006<S01><S>~2315~400~III.01~10~10~III.05~2020~35~~2010~20~~10~15~~145~155~~20~30~~35~30~~40~45~~50~50~~45~45~III.02~20~25~~25~20~~95~155~~10~20~~30~35~~35~50~~20~50~~1205~3130~III.03.04~1022~2650~~947~2380~~25~200~~50~70~~40~40~~20~10~~20~30~III.05~80~60~~30~40~~50~20~~63~380~~30~230~~10~120~~23~30~~3520~3530~~1860~20"
'Barcode_Scaned str2
'str2 = "aa999192300790401   00201200000000201230~~820~1060~~230~50~~100~10~~20~20~III.06~20~30~~30~340~~40~50~~10~100~~110~210~~230~210~~10~20~~20~20~~1040~970~~10~20~~10~30~~120~310~~300~100~~200~300~~400~210~~1660~1500~III.07~1660~1500~~50~30~~310~210~~540~210~~200~100~~210~150~~200~300~~150~500~~3520~3530~~550~120~~10~20~~30~10~~20~10~~150~400~kjvfkw~26/11/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999192300790401   002012000000003012<S01-1><S>IV.08~10~20~~10~20~~0~0~~30~10~~-30~-10~~20~40~~350~250~~320~230~~10~20~~-370~-240~~310~420~~10~30~~300~390~IV.09~-70~150~~20~40~~-90~110~kjvfkw~26/11/2013</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999192300790401   002012000000004012<S01-2><S>~10~10~~20~20~~30~30~~120~210~~50~100~~100~100~~50~50~~380~520~~10~10~~20~10~~30~40~~50~210~~10~100~~20~30~~140~400~~20~"
'Barcode_Scaned str2
'str2 = "aa999192300790401   00201200000000501220~~310~210~~210~320~~200~300~~400~100~~120~210~~1260~1160~~1780~2080~~310~200~~100~300~~2190~2580~kjvfkw~26/11/2013</S></S01-2>"
'Barcode_Scaned str2
'str2 = "aa999192300790401   002012000000006012<S01-3><S>~10~20~~20~20~~120~210~~20~30~~120~100~~200~300~~490~680~~150~420~~100~200~~10~20~~250~300~~350~200~~300~200~~300~150~~200~190~~2150~2360~~10~20~~30~40~~210~200~~1"
'Barcode_Scaned str2
'str2 = "aa999192300790401   00201200000000701220~310~~240~200~~120~100~~120~320~~850~1190~~400~100~~200~10~~20~30~~100~200~~210~310~~230~310~~1160~960~~4160~4510~~10~10~~500~570~~4670~5090~kjvfkw~26/11/2013</S></S01-3>"
'Barcode_Scaned str2
'str2 = "aa999192300790401   002012000000008012<S01-4><S>50~0~60~110~70~60~10~0~20~10~20~20~20~0~10~50~30~20~20~0~30~50~20~20~230~0~220~150~230~270~10~0~10~20~20~30~120~0~100~10~120~230~100~0~110~120~90~10~100~0~10~120~20~110~90~0~80~0~170~10~30~0~20~20~40~60~10~0~10~10~20~50~20~0~10~10~20~10~210~0~190~170~230~130~100~0~90~80~110~70~110~0~100~90~120~60~200~0~190~180~170~160~250~0~200~150~300~100~190~0~180~170~160~150~180~0~170~160~150~100~100~0~80~60~120~20~150~0~100~50~20~10~190~0~150~100~100~50~100~0~190~150~140~20~650~0~450~500~600~150~200~0~150~200~150~50~"
'Barcode_Scaned str2
'str2 = "aa999192300790401   002012000000009012200~0~150~200~150~50~250~0~150~100~300~50~100~0~50~30~120~20~330~0~250~210~180~120~80~0~70~60~50~10~100~0~80~70~60~50~150~0~100~80~70~60~460~0~380~300~460~20~150~0~100~90~80~70~120~0~100~50~170~-50~90~0~90~80~100~-10~100~0~90~80~110~10~100~0~20~30~20~10~420~0~340~270~450~110~100~0~90~80~70~50~120~0~100~90~130~10~200~0~150~100~250~50~5~0~5~5~5~5~329~0~299~279~189~169~9~0~9~9~9~9~120~0~100~90~10~10~200~0~190~180~170~150~210~0~200~190~170~160~500~0~450~400~300~200~0~250~400~150~250~290~0~190~180~150~30~160~0~0~0~0~0~"
'Barcode_Scaned str2
'str2 = "aa999192300790401   0020120000000100120~0~1590~1269~880~399~1201~0~400~380~360~20~380~0~200~190~180~10~190~0~200~190~180~10~190~0~90~80~70~10~80~0~50~10~10~10~50~0~100~90~80~10~90~0~150~100~50~50~100~0~250~200~100~100~150~0~200~150~100~50~150~0~150~150~10~140~10~0~200~109~100~9~191~0~250~200~150~50~200~0~300~250~200~50~250~0~1120~870~520~500~750~0~100~200~100~100~10~0~20~20~20~20~20~0~100~100~50~30~20~0~150~100~100~50~150~0~150~100~50~50~100~0~200~50~100~50~250~0~200~150~50~100~100~0~200~150~50~100~100~0~740~410~250~200~330~0~200~200~100~100~100~0~250"
'Barcode_Scaned str2
'str2 = "aa999192300790401   002012000000011012~100~50~10~20~0~190~90~80~80~110~0~50~40~30~20~10~0~50~30~40~10~20~0~90~20~10~50~80~0~100~20~20~10~100~0~20~10~10~100~20~0~90~80~70~50~80~0~310~70~150~100~292~0~20~20~100~50~100~0~90~10~10~10~90~0~100~20~20~20~2~0~100~20~20~20~100~0~400~330~20~220~90~0~200~130~10~120~80~0~200~200~10~100~10~0~500~180~180~80~500~0~100~20~30~40~110~0~200~130~120~10~190~0~200~30~30~30~200~0~250~100~250~120~400~0~500~120~200~300~120~0~200~100~200~300~300~0~290~100~100~30~290~0~200~10~20~20~210~0~90~90~80~10~80~600~470~310~230~680~390~1"
'Barcode_Scaned str2
'str2 = "aa999192300790401   00201200000001201200~80~80~50~130~50~200~150~80~70~210~140~200~120~100~100~200~120~100~120~50~10~140~80~100~90~80~80~100~90~180~120~150~170~160~140~100~20~30~40~90~30~50~60~70~80~40~70~30~40~50~50~30~40~45~50~55~60~40~55~100~120~130~140~90~130~150~100~40~45~145~105~120~100~120~90~150~70~120~50~60~80~100~70~100~20~30~40~90~30~20~30~30~40~10~40~50~50~60~70~40~60~50~50~60~70~40~60~30~30~30~40~20~40~100~90~80~80~100~90~100~200~300~350~50~250~100~200~10~20~90~210~100~50~10~20~90~60~50~55~55~55~50~55~100~110~120~130~90~120~~</S></S01-4>"
'Barcode_Scaned str2

'str2 = "aa999982300790401   11201300100200100301/0114/06/2006<S01><S>2100462770</S><S>0~x~x~01/01/2013~1~0~0~1~~26/11/2013~abctvxq</S><S>12.55~1~12~12.55~1~23~-22~24.55"
'Barcode_Scaned str2
'str2 = "aa999982300790401   112013001002002003</S><S>12334~vssc~12211~ugfg~1233~asasasa ~23455~4rew4~eg455~2343343</S><S>du©n~kh123~duan~26/11/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999982300790401   112013001002003003<S01-1><S>-22</S><S>2300790384~gfsdgf~20~-4~~2300790384~efwef~20~-4~~0200471077~sdfdsf~20~-4~~1400633697~dsfew~20~-4~~2300790384~ewrewr~20~-4~</S><S>100~-20</S></S01-1>"
'Barcode_Scaned str2

'str2 = "aa999642300790401   11201300100100100101/0101/01/2009<S01><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/001~AB/12T~10~0000001~0000010~26/12/2013~test~6868686868~1321321~01/10/2013~~H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/002~AB/12T~10~0000011~0000020~26/12/2013~dfsf~6868686868~23432~01/01/2013~~H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/003~AB/12T~10~0000021~0000030~26/12/2013~dsfqwer~6868686868~4325435~01/01/2013~~H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/004~AB/12T~10~0000031~0000040~26/12/2013~12e~6868686868~435435~01/01/2013~~H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/005~AB/12T~10~0000041~0000050~26/12/2013~ewrrw~6868686868~5234234~01/01/2013~</S><S>~~dsgfdgfdgbd~26/11/2013~fsdgsdfds</S></S01>"
'Barcode_Scaned str2

''05-TNCN - Chuan 1.5
'str2 = "aa999172300790401   00201200500500100501/0101/01/2009<S05><S>2100462770</S><S>4~2~2~22200000~1200000~7000000~14000000~15200000~1200000~0~14000000~2898000~98"
'Barcode_Scaned str2
'str2 = "aa999172300790401   002012005005002005000~0~2800000~0~0~0~0~1000000~1898000~0~1~53000~11000~0~42000</S><S>Huy“n Linh~KTV~Hoµng KK~04/12/2013~1~~</S></S05>"
'Barcode_Scaned str2



End Sub

Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    If mOnLoad Then
'        mOnLoad = False
'        If UBound(arrStrElements) = 0 Then _
'            Unload Me
'    End If
'    If Not blnReceiveByBarcode And mOnLoad Then
'        mOnLoad = False
'        ShowFormReceiveFromFile
'    End If
'    Exit Sub
'ErrHandle:
'    SaveErrorLog Me.Name, "Form_Activate", Err.Number, Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error GoTo ErrorHandle
'    Dim strHelpContexID As String
'    Dim i As Integer
'    Dim lCol As Long, lRow As Long
'
'    If KeyCode = vbKeyF1 Then
'        fpSpread1.Sheet = mCurrentSheet
'        lCol = fpSpread1.ActiveCol
'        lRow = fpSpread1.ActiveRow
'        GetCellSpan fpSpread1, lCol, lRow
'        strHelpContexID = GetAttribute(xmlDocumentInit(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow)), "HelpContextID") 'Split(GetAttribute(xmlDocumentInit(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, lCol, lRow)), "HelpContexID"), "_")
'        If strHelpContexID <> vbNullString Then
'            fpSpread1.HelpContextID = CLng(strHelpContexID) 'Val(strHelpContexID(0) & strHelpContexID(1) & CStr(fpSpread1.ColLetterToNumber(strHelpContexID(2))) & strHelpContexID(3))
'        Else
'            fpSpread1.HelpContextID = 0
'        End If
'    End If
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog Me.Name, "Form_KeyDown", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrHandle
    
    fpSpread1.EventEnabled(EventAllEvents) = False
    fpSpread1.ImportExcelBook GetAbsolutePath("..\InterfaceTemplates\Template.xls"), vbNullString
    fpSpread1.EventEnabled(EventAllEvents) = True
    
    blnOnLoadEvent = True
    
    SetControlCaption Me, Me.Name
    
    frmSystem.chkSaveQuestion.Visible = True
        
    App.HelpFile = App.path & "\HTKK_CQT.chm"
    
    Me.Top = (frmSystem.ScaleHeight - frmInterfaces.Height) / 2
    If Me.Top <= 0 Then Me.Top = 50
    
    Me.Left = (frmSystem.ScaleWidth - Me.Width) \ 2
    
    If Me.Left <= 0 Then Me.Left = 0
    
    lblLoading.Top = Frame1.Top + (Frame1.Height - lblLoading.Height) / 2
    lblLoading.Left = Frame1.Left + (Frame1.Width - lblLoading.Width) / 2
    lblConnecting.Top = lblLoading.Top
    lblConnecting.Left = lblLoading.Left
    lblExit.Top = lblLoading.Top
    lblExit.Left = lblLoading.Left
    
    If blnReceiveByBarcode Then
        ShowFormReceiveFromBarcode
    Else
        lblLoading.Visible = False
        lblConnecting.Visible = True
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()

On Error GoTo ErrHandle

    'If UBound(arrStrElements) > 0 Then
        SetFormCaption Me, imgCaption, lblCaption
        Me.Refresh
    If Not blnReceiveByBarcode And blnOnLoadEvent Then
        ShowFormReceiveFromFile
        blnOnLoadEvent = False
    End If
        
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Resize", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrHandle
    If blnReceiveByBarcode Then
        StopBarcodeReader
    End If
    
    Set objTaxBusiness = Nothing
    ReDim arrStrElements(0)
    
    frmSystem.chkSaveQuestion.Visible = False
    frmSystem.chkQuetBangKe.Visible = False
    frmSystem.chkQuetBangKe.Value = False

    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Unload", Err.Number, Err.Description
End Sub

''' fpSpread1_ButtonClicked description
''' Update value for cell (checkbox cell)
''' Parameter1 pCol         : active column
''' Parameter2 pRow         : active row
''' Parameter3 ButtonDown   : left, right or center mouse button
Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    On Error GoTo ErrHandle
    
    With fpSpread1
        .Sheet = .ActiveSheet 'mCurrentSheet
        GetCellSpan fpSpread1, Col, Row
        .Col = Col
        .Row = Row
        If .CellType = CellTypeCheckBox Then
            UpdateCell Col, Row, IIf(ButtonDown = 1, "x", vbNullString)
        End If
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "fpSpread1_ButtonClicked", Err.Number, Err.Description
End Sub

Private Sub fpSpread1_Change(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo ErrHandle
    Dim lValue As String
    Dim IsUpdate As Boolean

    If mOnLoad = True Then Exit Sub
    With fpSpread1
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet 'mCurrentSheet

        .Col = Col
        .Row = Row
        If .Lock = False Then
            ' When user change value of cell, call UpdateCell function
            If .CellType = CellTypeNumber Then
                lValue = .Value
            Else
                lValue = .Text
            End If
            Select Case .CellType
                Case CellTypeCheckBox '10
                    ' Checkbox
                    IsUpdate = UpdateCell(Col, Row, IIf(Val(lValue) = 1, "x", vbNullString))
                Case Else
                    IsUpdate = UpdateCell(Col, Row, lValue)
            End Select
        End If
        .EventEnabled(EventAllEvents) = True
    End With

    Exit Sub

ErrHandle:
    SaveErrorLog Me.Name, "fpSpread1_Change", Err.Number, Err.Description
End Sub

Private Sub fpSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyDelete Then Exit Sub
'    With fpSpread1
'        .Sheet = .ActiveSheet
'        .Col = .ActiveCol
'        .Row = .ActiveRow
'        If .CellType = CellTypeNumber Then
'            .EventEnabled(EventAllEvents) = False
'            '.Text = vbNullString
'            fpSpread1_Change .Col, .Row
'            .EventEnabled(EventAllEvents) = True
'
'        End If
'
'    End With
End Sub

Private Sub MSComm1_OnComm()
    Static strTemp As String
    Dim i As Long
    Dim varBuff As Variant
    Dim lByte() As Byte
    
On Error GoTo ErrHandle
    Select Case MSComm1.CommEvent
        Case comEvReceive                                       ' Received RThreshold  of chars.
            varBuff = MSComm1.Input
            lByte = varBuff
            For i = 0 To UBound(lByte)
                If Chr$(lByte(i)) <> "" Then
                    strTemp = strTemp & Chr$(lByte(i))
                Else
                    'Debug.Print "a1: " & strTemp
                    'sua loi convert font don vi tinh to khai TTDB
                    'nvanhai sua ngay 01\07\2010
                    
                    strTemp = TAX_Utilities_Srv_New.Convert(strTemp, TCVN, UNICODE)
                    
                    Barcode_Scaned strTemp
                    'Debug.Print "a2: " & strTemp
                    strTemp = vbNullString
                End If
            Next
    End Select
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "MSComm1_OnComm", Err.Number, Err.Description
End Sub

'**********************************************
'Description: Barcode_Scaned procedure scan barcode image
'             and process barcode string to check the complete barcode
'Input:
'       strBarcode: the scanned barcode string
'Output:
'Return:
'**********************************************
Private Sub Barcode_Scaned(strBarcode As String)
    Dim intBarcodeCount As Long, intBarcodeNo As Long
    Dim strPrefix       As String, strBarcodeCount As String, strData As String
    Dim idToKhai        As String

    On Error GoTo ErrHandle

    'Check ngay client va ngay tren server
    If (strNgayHeThongSrv <> "" And strNgayHeThongSrv <> vbNullString) Then
        Dim dNgayHeThong As Date
        dNgayHeThong = CDate(strNgayHeThongSrv)
        Dim dCurrent As Date
        dCurrent = CDate(DateTime.Now)

        If (DateDiff("d", dCurrent, dNgayHeThong) <> 0) Then
            DisplayMessage "0143", msOKOnly, miCriticalError
            Exit Sub
        End If
    End If

    'End check ngay

    'Convert from TCVN to UNICODE format
    strBarcode = TrimString(strBarcode)
    'Debug.Print strBarcode
    strBarcode = Replace(strBarcode, "&", "", 1)
    'strBarcode = TAX_Utilities_Srv_New.Convert(strBarcode, TCVN, UNICODE)
    
    If Left$(strBarcode, 1) <> "0" Then

        'Version 1.2.0 and later
        ' Kiem tra neu version in to khai lon hon max_verion thi khong cho phep nhan
        'If Val(Left$(strBarcode, 3)) > Val(Replace$(APP_VERSION, ".", "")) Then
        If Val(Left$(strBarcode, 3)) > Val(Replace$(HTKK_LAST_VERSION, ".", "")) Then
            'Version tai doanh nghiep lon hon tai co quan thue
            DisplayMessage "0074", msOKOnly, miCriticalError
            Exit Sub
        ElseIf Val(Left$(strBarcode, 3)) < 200 Then ' Truong hop cac to khai thue TNCN duoc in tu phien ban 1.3.1 hieu luc trong nam 2008 thi ko cho nhan

            If Val(Mid$(strBarcode, 4, 2)) = 15 Or Val(Mid$(strBarcode, 4, 2)) = 16 Or Val(Mid$(strBarcode, 4, 2)) = 22 Or Val(Mid$(strBarcode, 4, 2)) = 23 Then
                DisplayMessage "0089", msOKOnly, miCriticalError
                Exit Sub
            End If
        End If

        strPrefix = Left$(strBarcode, 36)
        strBarcodeCount = Right$(strPrefix, 6)
        strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 6)
        
        ' nvhai xu ly nhan BCTC cho phien ban truoc 2.5.0
        ' BCTC in bang HTKK 2.1.0 in tung to rieng biet , tu phien ban 2.5.0 in gop 4 to thanh 1 bo
        ' begin
        ' 1. To khai CDKT phien ban 2.1.0 co ID la 18 (bo 15) -> chuyen thanh ID la 55
        ' 2. To khai CDKT phien ban 2.0.0 co ID la 18 (bo 15) -> chuyen thanh ID la 55
        If Left$(strPrefix, 3) = "210" Or Left$(strPrefix, 3) = "200" Or Left$(strPrefix, 3) = "131" Or Left$(strPrefix, 3) = "130" Then
            idToKhai = Mid(strPrefix, 4, 2)

            If Trim(idToKhai) = "18" Then
                strPrefix = Left$(strPrefix, 3) & "55" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If

            If Trim(idToKhai) = "19" Then
                strPrefix = Left$(strPrefix, 3) & "56" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If

            If Trim(idToKhai) = "20" Then
                strPrefix = Left$(strPrefix, 3) & "57" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If

            If Trim(idToKhai) = "21" Then
                strPrefix = Left$(strPrefix, 3) & "58" & Mid$(strPrefix, 6, Len(strPrefix) - 5)
            End If
        End If
        
        ' end
        
        ' nvhai
        ' 22/07/2010 chan to khai 01/TAIN, 02/TAIN ke khai tu thang 7 va 03/TAIN ke khai tu 2010 in bang HTKK 2.5.2 tro xuong
        If (Trim(Mid(strPrefix, 4, 2)) = "06" And Val(Mid(strPrefix, 19, 2)) > 6 And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Or (Trim(Mid(strPrefix, 4, 2)) = "09" And Val(Mid(strPrefix, 19, 2)) > 6 And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252) Then
            DisplayMessage "0102", msOKOnly, miInformation
            Exit Sub
        End If
        
        If Trim(Mid(strPrefix, 4, 2)) = "08" And Val(Mid(strPrefix, 21, 4)) > 2009 And Val(Left$(strPrefix, 3)) <= 252 Then
            DisplayMessage "0103", msOKOnly, miInformation
            Exit Sub
        End If
        
        ' end
        
        ' Bat dau
        ' To khai 04/TNCN bat dau thu thang 2 se ko nhan nua
        If Left$(strPrefix, 3) = "250" Or Left$(strPrefix, 3) = "210" Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 04AB/TNCN thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "39" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "40" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0093", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' To khai 07/TNCN phien ban 2.1.0 bat dau thu thang 2 se ko nhan nua
        If Left$(strPrefix, 3) = "210" Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 07/TNCN thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "36" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0097", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        ' Lay ID cua to khai de xem co hien thi luon hay ko (Cac to khai quyet toan TNCN)
        If Left(strPrefix, 3) = "250" Then
            ' Kiem tra neu la phien ban 250 va la cac to khai quyet toan TNCN thi phai dat lai tong so ma vach la 1
            idToKhai = Mid(strPrefix, 4, 2)

            If Trim(idToKhai) = "17" Or Trim(idToKhai) = "41" Or Trim(idToKhai) = "42" Or Trim(idToKhai) = "43" Then
                strBarcodeCount = Left(strBarcodeCount, Len(strBarcodeCount) - 1) & "1"
            End If
        End If
        
        ' to khai 06, 02_TNCN_SX, 02_TNCN_BH khong quet bang ke
        'If isIHTKK = True Then
        idToKhai = Mid(strPrefix, 4, 2)

        If Trim(idToKhai) = "59" Or Trim(idToKhai) = "43" Or Trim(idToKhai) = "42" Then
            strBarcodeCount = Left(strBarcodeCount, Len(strBarcodeCount) - 1) & "1"
        End If

        'End If
        If Trim(idToKhai) = "17" Then
            strBarcodeCount = Left(strBarcodeCount, Len(strBarcodeCount) - 1) & "2"
        End If
        
        ' Doi voi cac to khai thang quy/TNCN nay da bi thay doi ID giua version 210 va 250
        ' Dat lai cho ID cua 210 dung voi 250 de nhan vao QLT_NTK
        If Left$(strPrefix, 3) = "210" Or Left$(strPrefix, 3) = "200" Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 02/TNCN thang cua nam 2009 co ID = 15 thi phai set lai gia tri moi co ID = 53
            If Trim(idToKhai) = "15" And UBound(Split(Mid$(strBarcode, 37), "~")) <> 11 Then
                strPrefix = Left$(strPrefix, 3) & "53" & Mid(strPrefix, 6, Len(strPrefix) - 5)
            End If

            ' Neu la to khai 03/TNCN thang cua nam 2009 co ID = 16 thi phai set lai gia tri moi co ID = 54
            If Trim(idToKhai) = "16" And UBound(Split(Mid$(strBarcode, 37), "~")) <> 11 Then
                strPrefix = Left$(strPrefix, 3) & "54" & Mid(strPrefix, 6, Len(strPrefix) - 5)
            End If
        End If
        
        ' To khai 02/TNCN, 03/TNCN bat dau tu thang 2 se ko nhan theo TT84 nua
        'dhdang sua cho nhan to khai tu ban HTKK 2.5.0
        'date:21/04/2010
        'If (Left$(strPrefix, 3) = "250") Or (Left$(strPrefix, 3) = "210") Then
        If (Left$(strPrefix, 3) = "210") Then
            idToKhai = Mid(strPrefix, 4, 2)

            ' Neu la to khai 02AB/TNCN, 03AB/TNCN  thang bat dau tu thang 2/2010 se ko nhan to khai nua
            If (Trim(idToKhai) = "53" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "37" And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "54" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "38" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0094", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '07072011 TT28
        ' Khong nhan cac to khai theo mau cua
        idToKhai = Mid(strPrefix, 4, 2)

        If (Val(Left$(strPrefix, 3)) < 300) Then
            If Trim(idToKhai) = "01" Or Trim(idToKhai) = "02" Or Trim(idToKhai) = "04" Or Trim(idToKhai) = "11" Or Trim(idToKhai) = "12" Or Trim(idToKhai) = "46" Or Trim(idToKhai) = "47" Or Trim(idToKhai) = "48" Or Trim(idToKhai) = "49" Or Trim(idToKhai) = "15" Or Trim(idToKhai) = "16" Or Trim(idToKhai) = "50" Or Trim(idToKhai) = "51" Or Trim(idToKhai) = "36" Or Trim(idToKhai) = "70" Or Trim(idToKhai) = "06" Or Trim(idToKhai) = "05" Then
                DisplayMessage "0113", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '06012012 TT28
        ' Khong nhan cac to khai theo mau cu GD2
        If (Val(Left$(strPrefix, 3)) < 310) Then
            If Trim$(idToKhai) = "71" Or Trim$(idToKhai) = "72" Or Trim$(idToKhai) = "73" Or Trim$(idToKhai) = "03" Or Trim$(idToKhai) = "74" Or Trim$(idToKhai) = "75" Or Trim$(idToKhai) = "80" Or Trim$(idToKhai) = "81" Or Trim$(idToKhai) = "82" Or Trim$(idToKhai) = "17" Or Trim$(idToKhai) = "42" Or Trim$(idToKhai) = "43" Or Trim$(idToKhai) = "59" Or Trim$(idToKhai) = "76" Or Trim$(idToKhai) = "41" Or Trim$(idToKhai) = "77" Or Trim$(idToKhai) = "86" Or Trim$(idToKhai) = "87" Or Trim$(idToKhai) = "89" Then
                DisplayMessage "0126", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '17102011 khong nhan cac mau an chi in ra bang HTKK phien ban nho hon 302
        If (Val(Left$(strPrefix, 3)) < 302) Then
            If Trim(idToKhai) = "64" Or Trim(idToKhai) = "65" Or Trim(idToKhai) = "66" Or Trim(idToKhai) = "67" Or Trim(idToKhai) = "68" Then
                DisplayMessage "0122", msOKOnly, miInformation
                Exit Sub
            End If
        End If

        ' Ket thuc
        ' Khong nhan cac to khai 02/TAIN, 05/TNDN
        If Trim(idToKhai) = "08" Or Trim(idToKhai) = "24" Then
            DisplayMessage "0120", msOKOnly, miInformation
            Exit Sub
        End If

        ' end
        
        strBarcode = Mid$(strBarcode, 37)
        intBarcodeNo = CInt(Val(Left$(strBarcodeCount, 3)))
        intBarcodeCount = CInt(Val(Right$(strBarcodeCount, 3)))
        
        If intBarcodeNo = 0 Or intBarcodeCount = 0 Then
            MessageBox "0054", msOKOnly, miCriticalError
            Exit Sub
        End If
        
        If strTaxReportInfo = "" Then

            ' Scanning
            If UBound(arrStrElements()) = 0 Then
                ProgressBar1.max = intBarcodeCount
                ProgressBar1.Value = 0
                arrStrElements(0) = strPrefix
                cmdViewNow.Enabled = True
            Else

                If IsDifferent(strPrefix, arrStrElements(0)) Then

                    'Another tax report
                    If MessageBox("0035", msYesNo, miQuestion) = mrYes Then
                        ReDim arrStrElements(0)
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                    Exit Sub
                Else

                    If ProgressBar1.max <> intBarcodeCount Then
                        MessageBox "0062", msOKOnly, miCriticalError
                        Exit Sub
                    End If
                End If
            End If
            
            ReDim Preserve arrStrElements(intBarcodeCount)
            arrStrElements(intBarcodeNo) = strBarcode
            ' hlnam Edit
            ' Lay them trong truong hop ko quet het ma vach ma muon hien thi luon
            ReDim Preserve arrBCBuffer(intBarcodeCount)
            arrBCBuffer(intBarcodeNo) = strPrefix & strBarcodeCount & strBarcode

            ' hlnam End
            If IsCompleteData(strData) Then
                Dim tmp As String

                ' Check version <= 3.1.6
                If Val(Left$(strData, 3)) <= 316 Then
                    If Mid$(strData, 4, 2) = "01" Or Mid$(strData, 4, 2) = "02" Or Mid$(strData, 4, 2) = "04" Or Mid$(strData, 4, 2) = "71" Or Mid$(strData, 4, 2) = "36" Then
                        If Val(idToKhai) <> 36 Then
                            tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) - 5)
                            strData = tmp & "~0" & Right$(strData, Len(strData) - InStr(1, strData, "</S01>", vbTextCompare) + 5)
                        Else
                            strData = Left$(strData, Len(strData) - 10) & "~0" & Right$(strData, 10)
                        End If

                    ElseIf Mid$(strData, 4, 2) = "68" Then
                        tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) - 5)
                        strData = tmp & "~1" & Right$(strData, Len(strData) - InStr(1, strData, "</S01>", vbTextCompare) + 5)
                    ElseIf Mid$(strData, 4, 2) = "73" Then
                        tmp = Mid(strData, 1, InStr(1, strData, "</S02>", vbTextCompare) - 5)
                        strData = tmp & "~" & Right$(strData, Len(strData) - InStr(1, strData, "</S02>", vbTextCompare) + 5)
                    End If
                End If

                lblLoading.Visible = False
                lblConnecting.Visible = True
                frmInterfaces.Refresh

                If Not LoadForm(strData) Then
                    StartReceiveForm
                End If

                ' Free memory
                ReDim arrStrElements(0)
                ' Khai bao lai mot mang rong
                ' hlnam Edit
                ReDim arrBCBuffer(0)
                ' hlnam End
            End If
                
        Else ' Form is loaded

            If strTaxReportInfo = strPrefix Then
                MessageBox "0044", msOKOnly, miWarning
                Exit Sub
            Else

                If frmSystem.chkSaveQuestion.Value Then
                    cmdSave_Click

                    If blnSaveSuccess Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                Else

                    If MessageBox("0045", msYesNo, miQuestion) = mrYes Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If

    Else
        'Version 1.1.0 and 1.0.0
        strPrefix = Left$(strBarcode, 25)
        strBarcodeCount = Right$(strPrefix, 4)
        strPrefix = Mid(strPrefix, 1, Len(strPrefix) - 4)
        
        strBarcode = Mid$(strBarcode, 26)
        intBarcodeNo = CInt(Val(Left$(strBarcodeCount, 2)))
        intBarcodeCount = CInt(Val(Right$(strBarcodeCount, 2)))
        
        If intBarcodeNo = 0 Or intBarcodeCount = 0 Then
            MessageBox "0054", msOKOnly, miCriticalError
            Exit Sub
        End If
        
        If strTaxReportInfo = "" Then

            ' Scanning
            If UBound(arrStrElements()) = 0 Then
                ProgressBar1.max = intBarcodeCount
                ProgressBar1.Value = 0
                arrStrElements(0) = strPrefix
            Else

                If IsDifferent(strPrefix, arrStrElements(0)) Then

                    'Another tax report
                    If MessageBox("0035", msYesNo, miQuestion) = mrYes Then
                        ReDim arrStrElements(0)
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                    Exit Sub
                Else

                    If ProgressBar1.max <> intBarcodeCount Then
                        MessageBox "0062", msOKOnly, miCriticalError
                        Exit Sub
                    End If
                End If
            End If
            
            ReDim Preserve arrStrElements(intBarcodeCount)
            arrStrElements(intBarcodeNo) = strBarcode
            
            If IsCompleteData(strData) Then
                lblLoading.Visible = False
                lblConnecting.Visible = True
                frmInterfaces.Refresh

                If Not LoadForm(strData) Then
                    StartReceiveForm
                End If

                'Free memory
                ReDim arrStrElements(0)
            End If
                
        Else ' Form is loaded

            If strTaxReportInfo = strPrefix Then
                MessageBox "0044", msOKOnly, miWarning
                Exit Sub
            Else

                If frmSystem.chkSaveQuestion.Value Then
                    cmdSave_Click

                    If blnSaveSuccess Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    End If

                Else

                    If MessageBox("0045", msYesNo, miQuestion) = mrYes Then
                        StartReceiveForm
                        Barcode_Scaned (strPrefix & strBarcodeCount & strBarcode)
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Barcode_Scaned", Err.Number, Err.Description
End Sub

'
''****************************
''Description: GetCell function get cell string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetCell(pCellID As String, pFirstCell As String, pValue As String, _
'                        bolPrefix As Boolean, bolSuffix As Boolean, pOneCell As Boolean, xmlTemplateCells As MSXML.IXMLDOMNode, lCellCount As Long)
'
'On Error GoTo ErrHandle
'    If (bolPrefix = True And bolSuffix = False) Or (pOneCell = True) Then GetCell = "<Cells>" & vbCrLf
'    If bolSuffix = True And bolPrefix = True And pOneCell = False Then
'        GetCell = GetCell & "</Cells>" & vbCrLf
'        GetCell = GetCell & "<Cells>" & vbCrLf
'    End If
'    GetCell = GetCell & "<Cell CellID=""" & pCellID & """ "
'    If pFirstCell <> vbNullString Then
'        GetCell = GetCell & "FirstCell=""" & pFirstCell & """ Value=""" & pValue & """ "
'    Else
'        GetCell = GetCell & "Value=""" & pValue & """ "
'    End If
'
'    If GetAttribute(xmlTemplateCells.childNodes(lCellCount), "MCT") <> "" Then
'        GetCell = GetCell & "MCT=""" & GetAttribute(xmlTemplateCells.childNodes(lCellCount), "MCT") & """/>" & vbCrLf
'    Else
'        GetCell = GetCell & "/>" & vbCrLf
'    End If
'    If (bolSuffix = True And bolPrefix = False) Or (pOneCell = True) Then GetCell = GetCell & "</Cells>" & vbCrLf
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetCell", Err.Number, Err.Description
'End Function
'
''****************************
''Description: GetCell function get cells string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetCells(ByRef lIncreaseRow As Long, strCellsString As String, xmlTemplateCells As MSXML.IXMLDOMNode) As String
'
'On Error GoTo ErrHandle
'    Dim i As Long, lCol As Long, lRow As Long, lCellCount As Long
'    Dim strCellArr() As String
'    Dim strCellID As String, strFirstCell As String
'    Dim bolPrefix As Boolean, bolSuffix As Boolean
'    Dim blnBeginOfData As Boolean, blnEndOfData As Boolean
'    Dim bolOneCell As Boolean
'
'    strCellArr = Split(strCellsString, "~")
'
'    blnBeginOfData = True
'    blnEndOfData = False
'
'    i = 0
'    While i <= UBound(strCellArr) Or lCellCount < xmlTemplateCells.childNodes.length
'        If UBound(strCellArr) = 0 Then
'            bolOneCell = True
'        Else
'            bolOneCell = False
'        End If
'        bolPrefix = blnBeginOfData
'        bolSuffix = blnEndOfData
'
'        strFirstCell = GetAttribute(xmlTemplateCells.childNodes(lCellCount), "FirstCell")
'        If lCellCount >= xmlTemplateCells.childNodes.length Then
'            lCellCount = 0
'            lIncreaseRow = lIncreaseRow + GetDynRowCount(fpSpread1, xmlTemplateCells)
'            strFirstCell = "1"
'            bolPrefix = True
'            bolSuffix = True
'        End If
'        ParserCellID fpSpread1, GetAttribute(xmlTemplateCells.childNodes(lCellCount), "CellID"), lCol, lRow
'        lRow = lRow + lIncreaseRow
'        strCellID = GetCellID(fpSpread1, lCol, lRow)
'
'        If GetAttribute(xmlTemplateCells.childNodes(lCellCount), "Encode") <> "0" And i <= UBound(strCellArr) Then
'            GetCells = GetCells & GetCell(strCellID, strFirstCell, Replace(strCellArr(i), Chr$(20), "~"), bolPrefix, bolSuffix, bolOneCell, xmlTemplateCells, lCellCount)
'        Else
'            GetCells = GetCells & GetCell(strCellID, strFirstCell, "", bolPrefix, bolSuffix, bolOneCell, xmlTemplateCells, lCellCount)
'            i = i - 1
'        End If
'        lCellCount = lCellCount + 1
'
'        If i <= 0 Then
'            blnBeginOfData = False
'        End If
'
'        If i >= UBound(strCellArr) - 1 And lCellCount = xmlTemplateCells.childNodes.length - 1 Then
'            blnEndOfData = True
'        End If
'
'        i = i + 1
'    Wend
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetCells", Err.Number, Err.Description
'End Function
'
''****************************
''Description: GetCell function get Section string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetSection(ByRef lIncreaseRow As Long, xmlNodeDataSection As MSXML.IXMLDOMNode, xmlNodeTemplateSection As MSXML.IXMLDOMNode) As String
'    Dim xmlNodeCells As MSXML.IXMLDOMNode
'    Dim strDynamic As String, strMaxRow As String
'
'On Error GoTo ErrHandle
'    strDynamic = GetAttribute(xmlNodeTemplateSection.parentNode, "Dynamic")
'    strMaxRow = GetAttribute(xmlNodeTemplateSection.parentNode, "MaxRows")
'
'    GetSection = "<Section Dynamic=""" & strDynamic & """ MaxRows=""" & strMaxRow & """>" & vbCrLf
'
'    If Not xmlNodeDataSection Is Nothing Then
'        GetSection = GetSection & GetCells(lIncreaseRow, xmlNodeDataSection.Text, xmlNodeTemplateSection)
'    Else
'        GetSection = GetSection & GetCells(lIncreaseRow, " ", xmlNodeTemplateSection)
'    End If
'    GetSection = GetSection & "</Section>" & vbCrLf
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetSection", Err.Number, Err.Description
'End Function
'
''****************************
''Description: GetCell function get Sections string.
''Author: Tuyends
''Modify by: ThanhDX
''Date:18/11/2005
''Input:
''OutPut:
''Return:
''****************************
'Private Function GetSections(xmlNodeDataSections As MSXML.IXMLDOMNode, xmlNodeTemplateSections As MSXML.IXMLDOMNode) As String ', rsTaxInfor As ADODB.Recordset) As String
'    Dim i As Long, j As Long, lIncreaseRow As Long
'
'On Error GoTo ErrHandle
'    GetSections = "<!-- edited with XML Spy v4.1 U (http://www.xmlspy.com) by tuyends (FSS) -->" & vbCrLf
'    GetSections = GetSections & "<!DOCTYPE Sections SYSTEM ""Schema.dtd"">" & vbCrLf
'    GetSections = GetSections & "<Sections"
'
'    'Get all attr of  Section node
'    For i = 0 To xmlNodeTemplateSections.Attributes.length - 1
'        GetSections = GetSections & " " & xmlNodeTemplateSections.Attributes(i).xml
'    Next i
'    GetSections = GetSections & "> " & vbCrLf
'
'    lIncreaseRow = 0
'
'    GetSections = GetSections & xmlNodeTemplateSections.childNodes(0).xml & vbCrLf
'
'    For i = 1 To xmlNodeTemplateSections.childNodes.length - 1
'        If xmlNodeTemplateSections.childNodes(i).baseName = "Section" Then
'            GetSections = GetSections & GetSection(lIncreaseRow, xmlNodeDataSections.childNodes(i - 1).childNodes(0), xmlNodeTemplateSections.childNodes(i).childNodes(0))
'        End If
'    Next
'    GetSections = GetSections & "</Sections>" & vbCrLf
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetSections", Err.Number, Err.Description
'End Function
Private Sub GetCells(xmlSectionTemplate As MSXML.IXMLDOMNode, arrStrValue() As String)
    On Error GoTo ErrHandler
    Dim lCtrl As Long, lCtrl2 As Long
    
    'Fill data from array of data to Cell node
    While lCtrl <= UBound(arrStrValue) And Not xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2) Is Nothing
        If GetAttribute(xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2), "Receive") <> "0" Then
            SetAttribute xmlSectionTemplate.selectNodes("Cells/Cell")(lCtrl2), "Value", _
                Replace(Replace(arrStrValue(lCtrl), "1" & Chr$(20) & Chr$(20) & "1", ""), Chr$(20), "~")
        Else
            lCtrl = lCtrl - 1
        End If
        lCtrl = lCtrl + 1
        lCtrl2 = lCtrl2 + 1
    Wend
    
    Exit Sub
ErrHandler:
    SaveErrorLog Me.Name, "GetCells", Err.Number, Err.Description
End Sub

'*******************************************
'Description: GetSection procedure convert data from data string
'               to Dom data.
'Author: ThanhDX
'Date: 21/02/2006
'Input:
'   xmlSectionTemplate: Section template node
'   xmlSectionData : Section data node
'*******************************************
Private Sub GetSection(xmlSectionTemplate As MSXML.IXMLDOMNode, xmlSectionData As MSXML.IXMLDOMNode, blnValidData As Boolean)

On Error GoTo ErrHandler

    Dim lCtrl As Long
    Dim lElementsNo As Long
    Dim lDataNo As Long
    Dim arrStrValue() As String
    Dim idToKhaiCheck As Integer
    ' Khong check doi voi cac BCTC
    idToKhaiCheck = Val(TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    If (idToKhaiCheck >= 24 And idToKhaiCheck <= 35) Or (idToKhaiCheck >= 55 And idToKhaiCheck <= 58) Or (idToKhaiCheck >= 18 And idToKhaiCheck <= 21) Or idToKhaiCheck = 69 Then
        isSheetTk = False
    End If
    
    lElementsNo = GetElementsNo(xmlSectionTemplate.childNodes(0))
    'Get array of data units
    arrStrValue = Split(xmlSectionData.Text, "~")
    ' Lay ve so chi tieu cua chuoi ma vach
    lDataNo = UBound(arrStrValue)
    If lDataNo = -1 Then
        lDataNo = 0
    End If
    ' End
    
    If GetAttribute(xmlSectionTemplate, "Dynamic") = "0" Then
        ' to khai 01/GTGT cho phep nhan 2 CTMV
        If idToKhaiCheck = 1 Then
            'Static data
            ' Truong hop chuoi ma vach nhieu chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 > lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 3 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 3 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
        ElseIf idToKhaiCheck = 11 Then
             If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
        ElseIf idToKhaiCheck = 12 Then
             If ((lDataNo + 1 > lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 6)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 6)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
        ElseIf idToKhaiCheck = 3 Then
           If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 3 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
        Else
            'Static data
            ' Truong hop chuoi ma vach nhieu chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 > lElementsNo) And isSheetTk Then
            If (lDataNo + 1 > lElementsNo) And isSheetTk Then
                blnValidData = False
                checkSoCT = 1
                Exit Sub
            End If
            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
            If (lDataNo + 1 < lElementsNo) And isSheetTk Then
                blnValidData = False
                checkSoCT = 2
                Exit Sub
            End If
        
        End If
        
    Else
        ' Kiem tra neu chuoi ma vach bi thieu hoac thua chi tieu se khong cho phep nhan
        ' If ((UBound(arrStrValue) + 1) Mod lElementsNo <> 0) And isSheetTk Then
        If ((lDataNo + 1) Mod lElementsNo <> 0) And isSheetTk Then
            blnValidData = False
            checkSoCT = 3
            Exit Sub
        End If
        ' Dynamic data
        For lCtrl = 2 To IIf((UBound(arrStrValue) + 1) Mod lElementsNo = 0, _
                            (UBound(arrStrValue) + 1) / lElementsNo, (UBound(arrStrValue) + 1) \ lElementsNo + 1)
            'Insert nodes
            InsertNode xmlSectionTemplate
        Next lCtrl
    End If
    
    GetCells xmlSectionTemplate, arrStrValue
    
    blnValidData = True
    
    Exit Sub
ErrHandler:
    blnValidData = False
    SaveErrorLog Me.Name, "GetSection", Err.Number, Err.Description
End Sub

'*******************************************
'Description: GetSections procedure convert data from data string
'               to Dom data.
'Author: ThanhDX
'Date: 21/02/2006
'Input:
'   xmlSectionsTemplate: Sections template node
'   xmlSectionsData : Sections data node
'*******************************************
Private Sub GetSections(xmlSectionsTemplate As MSXML.IXMLDOMNode, xmlSectionsData As MSXML.IXMLDOMNode, blnValidData As Boolean)

On Error GoTo ErrHandler

    Dim xmlSectionNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    If xmlSectionsData.childNodes.length > xmlSectionsTemplate.childNodes.length Then
        blnValidData = False
        'DisplayMessage "0072", msOKOnly, miCriticalError
        Exit Sub
    End If
    
    For lCtrl = 1 To xmlSectionsData.childNodes.length
        GetSection xmlSectionsTemplate.childNodes(lCtrl), xmlSectionsData.childNodes(lCtrl - 1), blnValidData
        If Not blnValidData Then
            blnValidData = False
            Exit Sub
        End If
    Next
    blnValidData = True
    
    Exit Sub
ErrHandler:
    blnValidData = False
    SaveErrorLog Me.Name, "GetSections", Err.Number, Err.Description
End Sub

Private Function GetElementsNo(xmlCellsNode As MSXML.IXMLDOMNode) As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCntElementsNo As Long
    
    For Each xmlCellNode In xmlCellsNode.childNodes
        If GetAttribute(xmlCellNode, "Receive") <> "0" Then
            lCntElementsNo = lCntElementsNo + 1
        End If
    Next
    GetElementsNo = lCntElementsNo
End Function

Private Sub InsertNode(xmlSectionTemplate As MSXML.IXMLDOMNode)
    Dim xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode, xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim lRows As Long, lRow2s As Long
    Dim lRowLBound As Long, lRowUbound As Long
    Dim lRow As Long, lCol As Long
    
    Set xmlCellsNode = xmlSectionTemplate.lastChild
    lRows = GetDynRowCount(fpSpread1, xmlCellsNode, lRow2s, lRowLBound, lRowUbound)
    
    'Increase row value on each cell in Dom data
    IncreaseRowInDOM fpSpread1, xmlSectionTemplate.parentNode.parentNode, lRowUbound + 1, lRows, lRow2s
    
    Set xmlNodeNewCells = xmlCellsNode.CloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
        
'        ' Set new ID2 for node (CellID2)
'        ParserCellID fpSpread1, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
'        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
    Next
    
    ' Insert new node to DOM object
    xmlCellsNode.parentNode.appendChild xmlNodeNewCells
    
End Sub

Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To 0 Step -1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        End If
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "IncreaseRowInDOM", Err.Number, Err.Description
End Sub
'****************************
'Description: RestoreDataFile function restore
'             data files from data string.
'   Step 1: Cut data string into sheet datas
'   Step 2: Load content of sheet datas to DOM, load template to DOM
'   Step 3: Generate xml string and save it to xml file
'Author: ThanhDX
'Date:20/11/2005
'Input:
'       strBarcodeData: Data string.
'OutPut:
'Return: True if restore data file successfully
'        False if the otherwise.
'****************************
Private Function RestoreDataFile(ByVal strBarcodeData As String) As Boolean ', rsTaxInfor As ADODB.Recordset)
    Dim strDataRestore As String, strFileName As String
    Dim lIndex As Long, lCtrl As Long, arrStrData() As String
    Dim xmlData As New MSXML.DOMDocument, xmlTemplate As New MSXML.DOMDocument
    Dim fso As New FileSystemObject, tstFile As TextStream
    Dim blnValidData As Boolean
    
On Error GoTo ErrHandle
    arrStrData = GetSheetDatas(strBarcodeData)
        
    If UBound(arrStrData) < TAX_Utilities_Srv_New.NodeValidity.childNodes.length Then
        RestoreDataFile = False
        Exit Function
    End If
    
    For lIndex = 1 To UBound(arrStrData())
        ' Chi kiem tra so chi tieu tren to khai sheet 1
        If lIndex = 1 Then
            isSheetTk = True
        Else
            isSheetTk = False
        End If
        ' end
        xmlTemplate.Load GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "TemplateFolder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & ".xml"
        
        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
            strFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lIndex - 1), "DataFile") & "_00" & TAX_Utilities_Srv_New.Year & ".xml"
        End If
        
        If arrStrData(lIndex) <> vbNullString Then
            If Not xmlData.loadXML(arrStrData(lIndex)) Then
                RestoreDataFile = False
                Exit Function
            End If
            
            'Get data string and structure
            GetSections xmlTemplate.getElementsByTagName("Sections")(0), xmlData.firstChild, blnValidData
            If Not blnValidData Then
                RestoreDataFile = False
                Exit Function
            Else
                xmlTemplate.save strFileName
            End If
        End If
                        
        'Set tstFile = fso.CreateTextFile(strFileName, True, True)
        'tstFile.Write strDataRestore
        'tstFile.Close
    Next lIndex
    
    Set xmlData = Nothing
    Set xmlTemplate = Nothing
    Set fso = Nothing
    
    RestoreDataFile = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "RestoreDataFile", Err.Number, Err.Description
End Function

'****************************
'Description: SetNodeMenu procedure set value to menu node.
'Author: ThanhDX
'Date:20/11/2005
'Input:
'       strMenuId: Menu id string.
'OutPut:
'Return:
'****************************
Private Sub SetNodeMenu(strMenuId As String)
    Dim xmlMenuDom As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
On Error GoTo ErrHandle
    xmlMenuDom.Load App.path & "\Menu.xml"
    
    For Each xmlNode In xmlMenuDom.getElementsByTagName("Root")(0).childNodes
        If GetAttribute(xmlNode, "ID") = strMenuId Then
            TAX_Utilities_Srv_New.NodeMenu = xmlNode
            Exit For
        End If
    Next
    
    Set xmlNode = Nothing
    Set xmlMenuDom = Nothing
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetNodeMenu", Err.Number, Err.Description
End Sub

'****************************
'Description: SetPeriod procedure set value to month, threemonth and year property.
'Author: ThanhDX
'Date:20/11/2005
'Input:
'       strValue: Value set.
'OutPut:
'Return:
'****************************
Private Sub SetPeriod(ByVal strValue As String)

    On Error GoTo ErrHandle
    Dim strID As String
    strID = Left$(strTaxReportInfo, 2)
    
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        TAX_Utilities_Srv_New.Month = Left$(strValue, 2)

        If strID = "01" Or strID = "02" Or strID = "04" Or strID = "71" Or strID = "36" Or strID = "68" Then
            TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
        Else
            TAX_Utilities_Srv_New.ThreeMonths = ""
        End If

    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = 1 Then

        If strID = "68" Then
            TAX_Utilities_Srv_New.Month = Left$(strValue, 2)
            TAX_Utilities_Srv_New.ThreeMonths = ""
        Else
            TAX_Utilities_Srv_New.Month = ""
            TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
        End If

        TAX_Utilities_Srv_New.ThreeMonths = Left$(strValue, 2)
    End If
    
    TAX_Utilities_Srv_New.Year = Right$(strValue, 4)
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetPeriod", Err.Number, Err.Description
End Sub

'****************************
'Description: InitParameters function initialize parameters from data string
'   Step 1:
'   Step 2:
'   Step 3:
'Author:ThanhDX
'Date:25/11/2005
'Input:
'       strData: data string of tax report.
'Output:
'       rsTaxInfor: recordset contain query data.
'Return: true if initialize sucessfully
'        false if the otherwise
'****************************
'Private Function InitParameters(ByVal strData As String, ByRef rsTaxInfor As ADODB.Recordset) As Boolean
Private Function InitParameters(ByVal strData As String, arrStrHeaderData() As String) As Boolean

'ThanhDX modified
'Date: 10/04/06
    Dim strTaxID As String, strID As String
    Dim blnConnected As Boolean
    Dim strValidDate As String, strTempDate As String
    Dim rsParams As ADODB.Recordset
    Dim strPhongXuLy As String
    Dim rsTaxInfor As ADODB.Recordset
    
    ' Xu ly cho cac to khai lan phat sinh
    Dim arrCT() As String
    Dim strTemp As String
    
On Error GoTo ErrHandle
    
    TAX_Utilities_Srv_New.Month = ""
    TAX_Utilities_Srv_New.ThreeMonths = ""
    TAX_Utilities_Srv_New.Year = ""
    TAX_Utilities_Srv_New.FinanceStartDate = ""
    
    
'    If Left$(strData, 3) = "120" Then
'        lblVersion.caption = "1.2.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "130" Then
'    'Version 1.3.0
'        'Get version of application
'        lblVersion.caption = "1.3.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'
'    ElseIf Left$(strData, 3) = "131" Then
'    'Version 1.3.1
'        'Get version of application
'        lblVersion.caption = "1.3.1"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "200" Then
'    'Version 2.0.0
'        'Get version of application
'        lblVersion.caption = "2.0.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "210" Then
'    'Version 2.1.0
'        'Get version of application
'        lblVersion.caption = "2.1.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "250" Then
'    'Version 2.5.0
'        'Get version of application
'        lblVersion.caption = "2.5.0"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "251" Then
'    'Version 2.5.0
'        'Get version of application
'        lblVersion.caption = "2.5.1"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "252" Then
'    'Version 2.5.2
'    'Get version of application
'    lblVersion.caption = "2.5.2"
'    strTaxReportVersion = Left$(strData, 3)
'    strData = Mid$(strData, 4)
'    ElseIf Left$(strData, 3) = "253" Then
'        'Version 2.5.2
'        'Get version of application
'        lblVersion.caption = "2.5.3"
'        strTaxReportVersion = Left$(strData, 3)
'        strData = Mid$(strData, 4)
'    End If
    
    ' 03122010 - sua lai doan lay version cua ung dung in ma vach
    strTaxReportVersion = Left$(strData, 3)
    strData = Mid$(strData, 4)
    lblVersion.caption = Left$(strTaxReportVersion, 1) & "." & Mid$(strTaxReportVersion, 2, 1) & "." & Right$(strTaxReportVersion, 1)
    ' end doan lay version
    
    'Get info of barcode string --25 characters
    strTaxReportInfo = Left$(strData, 21)
    
    If xmlSQL.url = "" Then
        xmlSQL.Load App.path & "\SQL.xml"
    End If
    
    'Get Tax id
    strTaxID = Trim(Mid$(strTaxReportInfo, 3, 13))
    If Len(strTaxID) = 13 Then
        If Trim(Right(strTaxID, 3)) = "000" Then
            strTaxID = Mid$(strTaxID, 1, 10)
        Else
            strTaxID = Mid$(strTaxID, 1, 10) & "-" & Mid$(strTaxID, 11, 13)
        End If
    End If
    
    'Connect DB and get informations
    ' nvhai
    ' Yeu cau nhan BCTC khong can check co quan thue cua user dang nhap
    ' 09-06-2010
    ' begin
    Dim strIDBCTC As String
    strIDBCTC = Left$(strTaxReportInfo, 2)
     If (Val(strIDBCTC) = 24 Or Val(strIDBCTC) = 25 Or Val(strIDBCTC) = 26 Or Val(strIDBCTC) = 27 Or Val(strIDBCTC) = 28 Or Val(strIDBCTC) = 29 _
            Or Val(strIDBCTC) = 30 Or Val(strIDBCTC) = 31 Or Val(strIDBCTC) = 32 Or Val(strIDBCTC) = 33 Or Val(strIDBCTC) = 34 Or Val(strIDBCTC) = 35 _
            Or Val(strIDBCTC) = 55 Or Val(strIDBCTC) = 56 Or Val(strIDBCTC) = 57 Or Val(strIDBCTC) = 58 Or Val(strIDBCTC) = 18 Or Val(strIDBCTC) = 19 _
            Or Val(strIDBCTC) = 20 Or Val(strIDBCTC) = 21 Or Val(strIDBCTC) = 69) Then
        Set rsTaxInfor = GetTaxInfoBCTC(strTaxID, blnConnected)
    Else
        Set rsTaxInfor = GetTaxInfo(strTaxID, blnConnected)
    End If
    ' end
    
     'Connect DB fail
    If Not blnConnected Then _
        Exit Function
    
    If (rsTaxInfor Is Nothing Or rsTaxInfor.Fields.Count = 0) And Len(strTaxID) = 14 Then
        strTaxID = Replace(strTaxID, "-", " ")
        Set rsTaxInfor = GetTaxInfo(strTaxID, blnConnected)
    End If
    'Tax id is not exist
    If rsTaxInfor Is Nothing Or rsTaxInfor.Fields.Count = 0 Then
        InitParameters = False
        MessageBox "0041", msOKOnly, miCriticalError
        Exit Function
    End If
    
    'Tax id is closed
    
    If rsTaxInfor.Fields(0) = "01" Then
        InitParameters = False
        MessageBox "0042", msOKOnly, miCriticalError
        Exit Function
    End If
    
    'DTNT chuyen di noi khac
    ' reset trang thai kiem tra
    checkTT = 0
    If rsTaxInfor.Fields(0) = "02" Then
        InitParameters = False
        'checkTT = 1  ' Trang thai chuyen di noi khac
        MessageBox "0043", msOKOnly, miCriticalError
        Exit Function
    End If
    
    If rsTaxInfor.Fields(0) = "03" Then
    If (MessageBox("0140", msYesNo, miWarning) = mrNo) Then
        InitParameters = False
        Exit Function
    End If
    'InitParameters = False
    checkTT = 2  ' Trang thai DTNT mat tich
    'MessageBox "0087", msOKOnly, miCriticalError
    'Exit Function
    End If
    
   
    
'    'DTNT hien dang bi mat tich
'    If rsTaxInfor.Fields(0) = "03" Then
'        'InitParameters = False
'        checkTT = 2  ' Trang thai DTNT mat tich
'        'MessageBox "0087", msOKOnly, miCriticalError
'        'Exit Function
'    End If

    'DTNT hien dang tam dung kinh doanh co thoi han
    'sua theo mail cua ptly yeu cau khong kiem tra MST tam ngung kinh doanh van duoc nop
    'ngay 01.07.2010
    
    If rsTaxInfor.Fields(0) = "05" Then
        InitParameters = False
        MessageBox "0088", msOKOnly, miCriticalError
        Exit Function
    End If

    strMST = CStr(rsTaxInfor.Fields(1))
    
    If InStr(1, strData, "<S") < 35 Then
        'Ver 1.0
        ' Get NgayTaiChinh and ThangTaiChinh
        iNgayTaiChinh = 0
        iThangTaiChinh = 0
    Else
        'Ver 1.1.0 and later
        ' Get NgayTaiChinh and ThangTaiChinh
        strTempDate = Mid$(strData, 22, 5)
        iNgayTaiChinh = GetNgayTaiChinh(strTempDate)
        iThangTaiChinh = GetThangTaiChinh(strTempDate)
        TAX_Utilities_Srv_New.FinanceStartDate = strTempDate
    End If
    
    strID = Left$(strTaxReportInfo, 2)
    SetNodeMenu strID
    SetPeriod Right$(strTaxReportInfo, 6)
    TAX_Utilities_Srv_New.NodeValidity = GetValidityNode
    
    '*******************************
'Date: 13/02/2006
    'Gan gia tri tu ngay, den ngay.
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" Then
        TAX_Utilities_Srv_New.FirstDay = Mid$(strData, 37, 10)
        TAX_Utilities_Srv_New.LastDay = Mid$(strData, 47, 10)
    End If
'*******************************
'*******************************
'Date: 16/02/2006
    'Danh sach to khai can kiem tra ngay bat dau nam tai chinh
    On Error GoTo ThamSoErrHandle
    
'    Set rsParams = clsDAO.Execute("select gia_tri from rcv_thamso where ten ='LOAI_TK_TAICHINH'")
'
'    On Error GoTo ErrHandle
'    'Kiem tra ngay bat dau nam tai chinh doi voi cac loai to
'    '   khai co kiem tra ngay bat dau nam tai chinh
'    If InStr(1, "," & rsParams.Fields(0) & ",", "," & GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") & ",") <> 0 Then
'        If Not IsNull(rsTaxInfor.Fields("ngay_tchinh")) Then
'            If Mid$(rsTaxInfor("ngay_tchinh"), 1, 5) <> Mid$(strTempDate, 1, 5) Then
'                DisplayMessage "0065", msOKOnly, miCriticalError
'                Exit Function
'            End If
'            'Kiem tra ngay bat dau kinh doanh
'        Else 'Trong DB chua co gia tri ngay bat dau kinh doanh
'            DisplayMessage "0066", msOKOnly, miCriticalError
'            Exit Function
'        End If
'    End If
    
    'To khai ke khai tu ngay ... den ngay ...
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1" Then
        If Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) <> 17 Then
            If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" Then
                If Mid$(rsTaxInfor("ngay_tchinh"), 1, 5) <> Mid$(TAX_Utilities_Srv_New.FirstDay, 1, 5) Then
                   'Tu ngay phai bang ngay bat dau nam tai chinh
                   ' hoac ngay bat dau kinh doanh
                    DisplayMessage "0068", msOKOnly, miCriticalError
                    Exit Function
                End If
                ''Ky ke khai lon hon ngay bat dau kinh doanh
                'If CInt(Mid$(rsTaxInfor("ngay_kdoanh"), 7, 4)) > CInt(Mid$(TAX_Utilities_Srv_New.FirstDay, 7, 4)) Then
                '    DisplayMessage "0069", msOKOnly, miCriticalError
                '    Exit Function
                'End If
            End If
        End If
    End If
    
    'Kiem tra cach thuc tinh ky ke khai la tinh theo nam duong lich hay nam tai chinh
    On Error GoTo ThamSoErrHandle
    
'    Set rsParams = clsDAO.Execute("select gia_tri from rcv_thamso where ten ='THEO_NAM_TAICHINH'")
'    blnTinhTheoNamTaiChinh = IIf(rsParams.Fields(0) = 0 Or IsNull(rsParams.Fields(0)), False, True)
    
    ' Doi voi to khai TNDN quy 01A/TNDN va 01B/TNDN thi phai lay dung theo ky ke khai cua nam tai chinh
    ' Vi du MST co ngay bat dau Nam tai chinh la 01/04/2009 thi quy 1 se bat dau la 01/04/2009 va quy 4 se bat dau la ngay 01/01/2010
    ' Dat lai blnTinhTheoNamTaiChinh = True => Ham GetNgayDauQuy se tra lai dung ket qua
    ' ID = 11 => To khai 01A/TNDN, ID = 12 => To khai 01B/TNDN
    
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1" Then
        If Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) = 11 Or Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) = 12 Then
            blnTinhTheoNamTaiChinh = True
        End If
    End If
    
    
    On Error GoTo ErrHandle
    If Val(strIDBCTC) = 1 Or Val(strIDBCTC) = 2 Or Val(strIDBCTC) = 4 Or Val(strIDBCTC) = 71 Or Val(strIDBCTC) = 36 Then
        If Val(strIDBCTC) = 36 Then
            LoaiKyKK = LoaiToKhai(strData)
        Else
            Dim tmp As String
            tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) + 5)
            LoaiKyKK = LoaiToKhai(tmp)
        End If
    End If
    
    'Gan gia tri ngay dau ky
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
        dNgayDauKy = DateSerial(CInt(TAX_Utilities_Srv_New.Year), CInt(TAX_Utilities_Srv_New.Month), 1)
        dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)

        If Val(strIDBCTC) = 1 Or Val(strIDBCTC) = 2 Or Val(strIDBCTC) = 4 Or Val(strIDBCTC) = 71 Or Val(strIDBCTC) = 36 Then
            If LoaiKyKK = True Then
                dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
                dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
                dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
            End If

        End If

    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then

        If Val(strIDBCTC) = 68 And LoaiKyKK = False Then
            dNgayDauKy = DateSerial(CInt(TAX_Utilities_Srv_New.Year), CInt(TAX_Utilities_Srv_New.Month), 1)
            dNgayCuoiKy = DateAdd("m", 1, dNgayDauKy)
            dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
        Else
            dNgayDauKy = GetNgayDauQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
            dNgayCuoiKy = DateAdd("m", 3, dNgayDauKy)
            dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
        End If

    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
        dNgayDauKy = GetNgayDauNam(CInt(TAX_Utilities_Srv_New.Year), iThangTaiChinh, iNgayTaiChinh)
        dNgayCuoiKy = DateAdd("m", 12, dNgayDauKy)
        dNgayCuoiKy = DateAdd("d", -1, dNgayCuoiKy)
    End If
'*******************************
'*******************************
'Date: 11/01/2006
    'Check validity of start date.
    
    If InStr(1, strData, "<S") < 35 Then
        'Ver 1.0
        strTempDate = Mid$(strData, 22, 8)
        strValidDate = GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "StartDate")
        If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))) _
                , DateSerial(CInt(Mid$(strTempDate, 5, 4)), CInt(Mid$(strTempDate, 3, 2)), CInt(Mid$(strTempDate, 1, 2)))) = 0 Then
            DisplayMessage "0064", msOKOnly, miInformation
            Exit Function
        End If
    '*******************************
        'Get main content
        strData = Mid$(strData, 30)
    Else
        'Ver 1.1.0 and later
        strTempDate = Mid$(strData, 27, 10)
        
        ' Neu la to khai 02A-02B/TNCN, 03A-03B/TNCN ke khai trong phien ban 2.5.0 thi set lai ngay bat dau hieu luc la 01/01/2010
        ' Sau khi co phien ban 2.5.2 se bat rang buoc cac dieu kien ve phien ban tu 2.5.2 thi bo doan check nay di
        If (Trim(strID) = "15") Or (Trim(strID) = "16") Or (Trim(strID) = "50") Or (Trim(strID) = "51") Or (Trim(strID) = "46") Or (Trim(strID) = "47") _
        Or (Trim(strID) = "48") Or (Trim(strID) = "49") Or (Trim(strID) = "36") Then
            strTempDate = "01/01/2010"
        End If
        ' End
        
        strValidDate = GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "StartDate")
        If Not DateDiff("d", DateSerial(CInt(Mid$(strValidDate, 7, 4)), CInt(Mid$(strValidDate, 4, 2)), CInt(Mid$(strValidDate, 1, 2))) _
                , DateSerial(CInt(Mid$(strTempDate, 7, 4)), CInt(Mid$(strTempDate, 4, 2)), CInt(Mid$(strTempDate, 1, 2)))) = 0 Then
            ' Truong hop dang bi map nham ID cua cac BCTC giua version 2.5.1 ngay 17/03/1020 voi ID cua version 2.1.0
            ' Thi ko hien thi message nay ma map lai version, id cho dung voi to khai 2.5.1 roi quet lai
            ' Begin
            If Trim(strID) = "55" Or Trim(strID) = "56" Or Trim(strID) = "57" Or Trim(strID) = "58" Then
                Exit Function
            ' end
            ElseIf Trim(strID) = "70" Then
            Else
                DisplayMessage "0064", msOKOnly, miInformation
                Exit Function
            End If
        End If
        
    '*******************************
        'Get main content
        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") <> "0" Then
            If Trim(strID) = "70" Or Trim(strID) = "81" Or Trim(strID) = "91" Then
                strData = Mid$(strData, 37)
            Else
                strData = Mid$(strData, 57)
            End If
        Else
            strData = Mid$(strData, 37)
        End If
    End If
    'RestoreDataFile (strData)
    If Not RestoreDataFile(strData) Then  ', rsTaxInfor
        'So chi tieu tren ma vach nhieu hon so chi tieu tren to khai
        If checkSoCT = 1 Then
            MessageBox "0104", msOKOnly, miCriticalError
            Exit Function
        End If
        ' So chi tieu tren ma vach it hon so chi tieu tren to khai
        If checkSoCT = 2 Then
            MessageBox "0105", msOKOnly, miCriticalError
            Exit Function
        End If
        ' Kiem tra cac to khai co so dong dong (chi kiem tra duoc khac chu khong phan biet duoc truong hop thieu hoac thua)
        If checkSoCT = 3 Then
            MessageBox "0106", msOKOnly, miCriticalError
            Exit Function
        End If
        
        If blnReceiveByBarcode Then
            MessageBox "0057", msOKOnly, miCriticalError
        Else
            MessageBox "0053", msOKOnly, miCriticalError
        End If
        Exit Function
    End If
    '***********************************
    'Date: 21/05/06
    'Lay thong tin phong xu ly.
    Set rsPXL = GetPhongXuLy(strPhongXuLy, blnConnected)
    If Not blnConnected Then
        Exit Function
    End If
    
    If rsPXL Is Nothing Then
        DisplayMessage "0077", msOKOnly, miCriticalError
        Exit Function
    End If
    '***********************************
    '***********************************
    'Date 26/05/06
    
    'Gan thong tin Header vao mang
    If Not GetHeaderData(rsTaxInfor, arrStrHeaderData) Then
        DisplayMessage "0080", msOKOnly, miCriticalError
        Exit Function
    End If
    
    ' Truoc khi lay thong tin ve tep, chuyen cac ma quy uoc cho to khai cu ve ma quy uoc cho to khai moi
    ' Lay thong tin ma so tep va so thu tu to khai.
'    If Not GetThongTinTep(changeMaToKhai(strID), arrStrHeaderData) Then
'        DisplayMessage "0079", msOKOnly, miCriticalError
'        Exit Function
'    End If
    
    ' Lay so thu tu cua to khai da dua vao RCV_TKHAI_HDR
    ' So thu tu nay phai lay theo cung Nguoi nop thue, ky ke khai, va cung loai to khai
    ' An chi
    If (Val(strID) >= 64 And Val(strID) <= 68) Or Val(strID) = 91 Then
            ' An chi
                    ' 01/TBAC
        If Val(strID) = 64 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(UBound(arrCT) - 1)) <> "" Then
                ngayPS = arrCT(UBound(arrCT) - 1)
                isTKLanPS = True
            End If
        End If
            If Not getSoTTTK_AC(changeMaToKhai(strID), arrStrHeaderData, strData) Then
                DisplayMessage "0079", msOKOnly, miCriticalError
                Exit Function
            End If
    Else
        ' cac to khai binh thuong
        isTKLanPS = False
        isTKThang = False
        ngayPS = ""

        ' 02/TNDN
        If Val(strID) = 73 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(32)) <> "" Then
                ngayPS = arrCT(32)
                isTKLanPS = True
            End If
        End If
        ' 01/TTDB
        If Val(strID) = 5 Then
            strTemp = Left$(strData, InStr(1, strData, "</S></S01>") + 9)
            arrCT = Split(strTemp, "~")
            If Trim(arrCT(UBound(arrCT) - 1)) <> "" Then
                ngayPS = arrCT(UBound(arrCT) - 1)
                isTKLanPS = True
            End If
        End If

        ' Xy ly to khai 08, 08A/TNCN
        If Val(strID) = 74 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(2)) <> "" Then
                TuNgay = arrCT(2)
                DenNgay = arrCT(3)
                isTKThang = True
            End If
            
        End If
' 08A/TNCN
        If Val(strID) = 75 Then
            arrCT = Split(strData, "~")
            If Trim(arrCT(1)) <> "" Then
                TuNgay = Right$(arrCT(0), 7)
                DenNgay = arrCT(1)
                isTKThang = True
            End If
            
        End If
        
        ' To khai 01/NTNN
        If Val(strID) = 70 Then
            strTemp = Left$(strData, InStr(1, strData, "</S></S01>") + 9)
            arrCT = Split(strTemp, "~")
            If Trim(arrCT(UBound(arrCT))) <> "" And Left$(Trim(arrCT(UBound(arrCT))), 10) <> "</S></S01>" Then
                ngayPS = Left$(Trim(arrCT(UBound(arrCT))), 10)
                isTKLanPS = True
            End If
        End If
        ' To khai 03/NTNN
        If Val(strID) = 81 Then
            strTemp = Left$(strData, InStr(1, strData, "</S></S01>") + 9)
            arrCT = Split(strTemp, "~")
            If Trim(arrCT(UBound(arrCT))) <> "" And Left$(Trim(arrCT(UBound(arrCT))), 10) <> "</S></S01>" Then
                ngayPS = Left$(Trim(arrCT(UBound(arrCT))), 10)
                isTKLanPS = True
            End If
        End If
        
            
        If Not getSoTTTK(changeMaToKhai(strID), arrStrHeaderData) Then
            DisplayMessage "0079", msOKOnly, miCriticalError
            Exit Function
       End If
       
        ' 18122012
        ' to khai lan phat sinh trog ngay chi nhan 1 to khai
        If (Val(strID) = 70 Or Val(strID) = 73 Or Val(strID) = 81 Or Val(strID) = 5) And isTKLanPS = True Then
            If isToKhaiPsDaNhanTN = True Then
                DisplayMessage "0129", msOKOnly, miCriticalError
                Exit Function
            End If
        End If
            
    End If
    
    '***********************************
    
'     Kiem tra to khai ton tai theo mau cu QLT
    isTKDA30 = isDA30(strID, arrStrHeaderData)
        
        
    InitParameters = True
    Exit Function
ThamSoErrHandle:
    DisplayMessage "0078", msOKOnly, miCriticalError
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "InitParameters", Err.Number, Err.Description
End Function

Private Sub SetupSpread()
    On Error GoTo ErrHandle
    
    Dim lSheet As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .Sheet = lSheet
            .AllowCellOverflow = False
            .AllowEditOverflow = False
            .Appearance = AppearanceFlat
            .ArrowsExitEditMode = True
            '.GrayAreaBackColor = RGB(238, 238, 238)
            .GrayAreaBackColor = vbButtonFace
            
            .MaxCols = .DataColCnt - 1
            .MaxRows = .DataRowCnt - 1
            .GridShowHoriz = False
            .GridShowVert = False
            
            .EditModePermanent = True
            .EditModeReplace = True
            .ColHeadersShow = False
            .RowHeadersShow = False
            .BorderStyle = BorderStyleNone
            .EditEnterAction = EditEnterActionNext
            .ProcessTab = True
            .ScrollBarExtMode = True
            .ScrollBarTrack = ScrollBarTrackOff
            .ScrollBars = ScrollBarsBoth  'ScrollBarsVertical
            .SetActionKey ActionKeyClear, False, False, 0
            .TabStripPolicy = TabStripPolicyAsNeeded
            .TabStripFont.Name = "Tahoma"
            
            .TextTip = TextTipFloating
        
            If UCase(.SheetName) <> UCase("Header") Then
                .SheetName = GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(.Sheet - 1), "Caption")
            Else
                mHeaderSheet = .Sheet
            End If
            
            
            .SetTextTipAppearance "Tahoma", 8, False, False, RGB(255, 255, 235), &H0
            .Protect = True
        Next
        .ActiveSheet = 1
        .Sheet = 1
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "SetupSpread", Err.Number, Err.Description
End Sub

Private Sub FormatGrid()
    On Error GoTo ErrHandle
    
    Dim lSheet As Long, i As Long, j As Long
        
    With fpSpread1
        .ReDraw = False
        For lSheet = 1 To .SheetCount
            .Sheet = lSheet
            If .SheetVisible = True Then
                For i = 1 To .MaxRows
                    .Row = i
                    If .RowHeight(i) > 10 And .RowHeight(i) < 15 Then .RowHeight(i) = 14
                    For j = 1 To .MaxCols
                        .Col = j
                        
                        If .BackColor = 12632256 Then
                            'Form backcolor
                            '.BackColor = RGB(238, 238, 238)
                            .BackColor = vbButtonFace
                            Me.BackColor = .BackColor
                        End If
                        
                        If .BackColor = 9868950 Then
                            'Grid header backcolor
                            .BackColor = RGB(215, 215, 215)
                        End If
                        
                        If .BackColor = 16777164 Then
                            'Grid hight light 1 backcolor
                            .BackColor = RGB(233, 245, 254)
                        End If
                        
                        If .BackColor = 13434879 Then
                            'Grid hight light 2 backcolor
                            .BackColor = RGB(255, 255, 235)
                        End If
                        
                        If .CellType = CellTypeNumber Then
                            .TypeNumberDecimal = ","
                            .TypeNumberSeparator = "."
                            .TypeNumberNegStyle = TypeNumberNegStyle2
                        End If
                        
                        Select Case Trim(.Text)
                            Case "chk"
                                .CellType = CellTypeCheckBox
                                .TypeCheckCenter = True
                            Case "cbo"
                                .CellType = CellTypeComboBox
                                .Text = ""
                            Case "cmd"
                                .CellType = CellTypeButton
                            Case "picture"
                                .CellType = CellTypePicture
                        End Select
                    Next
                Next
            End If
        Next
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "FormatGrid", Err.Number, Err.Description
End Sub

'****************************
'Description: LoadForm function load tax report to screen
'   Step 1: Convert data string into data files.
'   Step 2: Load excel template and format grid.
'   Step 3: Load and fill data from data files to grid.
'Author:ThanhDX
'Date:25/11/2005
'Input:
'       strData: data string of tax report.
'Output:
'Return: true if load data sucessfully
'        false if the otherwise
'****************************
Private Function LoadForm(ByVal strData As String) As Boolean

On Error GoTo ErrHandle

    Dim rsHeaderData As ADODB.Recordset
    Dim arrStrHeaderData() As String
    Dim LoaiTk As String
    Dim strMST As String
    
    Dim dsTK_DLT As String
    
    Dim blnDLConnected As Boolean
    Dim strTaxDLID As String
    Dim rsTaxDLInfor As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    frmSystem.MousePointer = vbHourglass
    
    LoaiTk = Mid(strData, 4, 2)
    
    'If InitParameters(strData, rsHeaderData) = False Then
    If InitParameters(strData, arrStrHeaderData) = False Then
        ' Truong hop bi map ID sai BCTC giua phien ban 2.5.1 ban ngay 17/03/2010 voi phien ban 2.1.0
        ' Thi conver lai cho chuan ID BCTC va Init lai cac Parameter
        If Trim(LoaiTk) = "55" Then
            InitParameters "25118" & Mid(strData, 6), arrStrHeaderData
        ElseIf Trim(LoaiTk) = "56" Then
            InitParameters "25119" & Mid(strData, 6), arrStrHeaderData
        ElseIf Trim(LoaiTk) = "57" Then
            InitParameters "25120" & Mid(strData, 6), arrStrHeaderData
        ElseIf Trim(LoaiTk) = "58" Then
            InitParameters "25121" & Mid(strData, 6), arrStrHeaderData
        Else
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
            Exit Function
        End If
    End If
          
    mOnLoad = True
    fpSpread1.EventEnabled(EventAllEvents) = False
    LoadTemplate fpSpread1
    SetupSpread
    FormatGrid
    'LoadInitFiles
    
    
    ' Set cac thong tin cua DL thue
    'Get Tax id
    strMST = Trim(Mid$(strTaxReportInfo, 3, 13))
    
    If LoaiTk <> "64" And LoaiTk <> "65" And LoaiTk <> "" And LoaiTk <> "66" And LoaiTk <> "67" And LoaiTk <> "68" And LoaiTk <> "91" And LoaiTk <> "69" And LoaiTk <> "19" And LoaiTk <> "20" And LoaiTk <> "21" And LoaiTk <> "22" Then
        strTaxDLID = Mid(strData, InStr(1, strData, "<S>") + 3, InStr(1, strData, "</S>") - InStr(1, strData, "<S>") - 3)
    Else
        strTaxDLID = vbNullString
    End If
    
    
    If Len(Trim(strMST)) = 13 Then
        strMST = Left(strMST, 10) & "-" & Right(strMST, 3)
    End If
    strMaSoThue = strMST
    
    
    If Len(Trim(strTaxDLID)) = 13 Then
        strTaxDLID = Left(strTaxDLID, 10) & "-" & Right(strTaxDLID, 3)
    End If
    strMaDaiLyThue = strTaxDLID
    
    Set rsTaxDLInfor = GetTaxDLInfo(strMST, strTaxDLID, blnDLConnected)
        
    
    If Trim(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "Class")) <> vbNullString Then
        Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "Class"))
        Set objTaxBusiness.fps = fpSpread1
        'objTaxBusiness.strMaSoTep = strMaSoTep
        objTaxBusiness.strPhongXuLy = strMaPhongXuLy
        objTaxBusiness.strNgayNhanToKhai = strNgayNhanToKhai
        objTaxBusiness.strNguoiSuDung = strUserID
        
        ' set thong tin DL thue
        ' danh sach cac to khai se set thong tin dai ly thue TT28
        dsTK_DLT = "~01~02~03~04~05~06~11~12~46~47~48~49~15~16~50~51~36~70~71~72~73~74~75~80~81~82~77~86~87~89~17~42~43~59~76~41~"
'        If Trim(LoaiTk) = "01" Or Trim(LoaiTk) = "02" Or Trim(LoaiTk) = "04" Or Trim(LoaiTk) = "05" Or Trim(LoaiTk) = "06" Or Trim(LoaiTk) = "11" _
'        Or Trim(LoaiTk) = "12" Or Trim(LoaiTk) = "46" Or Trim(LoaiTk) = "47" Or Trim(LoaiTk) = "48" Or Trim(LoaiTk) = "49" Or Trim(LoaiTk) = "15" _
'        Or Trim(LoaiTk) = "16" Or Trim(LoaiTk) = "50" Or Trim(LoaiTk) = "51" Or Trim(LoaiTk) = "36" Or Trim(LoaiTk) = "70" Or Trim(LoaiTk) = "71" _
'        Or Trim(LoaiTk) = "72" Then
         If InStr(1, dsTK_DLT, "~" & Trim(LoaiTk) & "~", vbTextCompare) > 0 Then
            If Trim(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "Class")) <> vbNullString Then
                If Not (rsTaxDLInfor Is Nothing Or rsTaxDLInfor.Fields.Count = 0) Then
                    If Not objTaxBusiness Is Nothing Then
                        objTaxBusiness.strTenDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(0).Value), "", rsTaxDLInfor.Fields(0).Value), TCVN, UNICODE)
                        objTaxBusiness.strDiaChiDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(1).Value), "", rsTaxDLInfor.Fields(1).Value), TCVN, UNICODE)
                         objTaxBusiness.strDienThoaiDL = IIf(IsNull(rsTaxDLInfor.Fields(2).Value), "", rsTaxDLInfor.Fields(2).Value)
                        objTaxBusiness.strFaxDL = IIf(IsNull(rsTaxDLInfor.Fields(3).Value), "", rsTaxDLInfor.Fields(3).Value)
                        objTaxBusiness.strEmailDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(4).Value), "", rsTaxDLInfor.Fields(4).Value), TCVN, UNICODE)
                        objTaxBusiness.strSoHopDongDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(5).Value), "", rsTaxDLInfor.Fields(5).Value), TCVN, UNICODE)
                        objTaxBusiness.strNgayHopDongDL = IIf(IsNull(rsTaxDLInfor.Fields(6).Value), "", rsTaxDLInfor.Fields(6).Value)
                    End If
                End If
            End If
        End If

        
        ' set ngay bat dau nam tai chinh cho to khai 01ATNDN va 01BTNDN
        If LoaiTk = "11" Or LoaiTk = "12" Or LoaiTk = "03" Then
            objTaxBusiness.dNgayTC = dNgayDauKy
        End If
        ' end
        If Not objTaxBusiness.Prepared1 Then Exit Function
    End If
            
    SetupData fpSpread1

    If Not objTaxBusiness Is Nothing Then
        If Not objTaxBusiness.Prepared2(rsPXL) Then Exit Function
    End If
    
    ' set ma CQT
    If Not objTaxBusiness Is Nothing Then
        If Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) >= 64 And Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) <= 68 Then
            objTaxBusiness.strMaCQT = strTaxOfficeId
            ' lay ma phong quan ly
            'Get Tax id
            strMST = Trim(Mid$(Left$(strData, 21), 6, 13))
            If Len(strMST) = 13 Then
                strMST = Mid$(strMST, 1, 10) & "-" & Mid$(strMST, 11, 13)
            End If
            GetPhongQuanLy (strMST)
            objTaxBusiness.strMaPQL = strMaPhongQuanLy
            objTaxBusiness.strTenPQL = strTenPhongQuanLy
        End If
    End If
    
    ' Set Phong quan ly
    If Not objTaxBusiness Is Nothing Then
        If Val(LoaiTk) = 70 Or Val(LoaiTk) = 71 Or Val(LoaiTk) = 72 Or Val(LoaiTk) = 73 Or Val(LoaiTk) = 74 Or Val(LoaiTk) = 77 Or Val(LoaiTk) = 3 Or Val(LoaiTk) = 75 _
        Or Val(LoaiTk) = 80 Or Val(LoaiTk) = 81 Or Val(LoaiTk) = 82 Or Val(LoaiTk) = 86 Or Val(LoaiTk) = 87 Or Val(LoaiTk) = 89 Or Val(LoaiTk) = 17 Or Val(LoaiTk) = 42 Or Val(LoaiTk) = 43 _
        Or Val(LoaiTk) = 59 Or Val(LoaiTk) = 76 Or Val(LoaiTk) = 41 Then
            ' lay ma phong quan ly
            'Get Tax id
            strMST = Trim(Mid$(Left$(strData, 21), 6, 13))
            If Len(strMST) = 13 Then
                strMST = Mid$(strMST, 1, 10) & "-" & Mid$(strMST, 11, 13)
            End If
            GetPhongQuanLy (strMST)
            objTaxBusiness.strMaPQL = strMaPhongQuanLy
            objTaxBusiness.strTenPQL = strTenPhongQuanLy
        End If
    End If
    
    

    'Setup header data
    'SetupHeaderData rsHeaderData
    SetupHeaderData arrStrHeaderData

    If Not objTaxBusiness Is Nothing Then
        If Not objTaxBusiness.Prepared3 Then Exit Function
    End If
    fpSpread1.EventEnabled(EventAllEvents) = True
    cmdClear.Enabled = True
    cmdSave.Enabled = True
    cmdViewNow.Enabled = False
    fpSpread1.Visible = True
    
    lblLabelVersion.Left = 4460
    lblVersion.Left = 8650
    
    'If CLng(Replace$(strTaxReportVersion, ".", "")) < CLng(Replace$(APP_VERSION, ".", "")) Then
    If CLng(Replace$(strTaxReportVersion, ".", "")) < CLng(Replace$(HTKK_LAST_VERSION, ".", "")) Then
        lblWarning.Visible = True
    Else
        lblLabelVersion.Left = lblLabelVersion.Left + lblWarning.Width
        lblVersion.Left = lblVersion.Left + lblWarning.Width
    End If
    lblLabelVersion.Visible = True
    lblVersion.Visible = True
    
    
    If frmSystem.chkSaveQuestion = True Then
        cmdClear.SetFocus
    Else
        cmdSave.SetFocus
    End If
    ' Loai to khai la GTGT khau tru mau 01/GTGT thi cho hien check quet du lieu bang ke
    If (verToKhai = 0 And LoaiTk = "01") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Or _
                TAX_Utilities_Srv_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.Value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    ' Loai to khai la GTGT hoa hong dai ly mau 02/GTGT thi cho hien check quet du lieu bang ke
    ElseIf (verToKhai = 0 And LoaiTk = "02") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.Value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    ' Loai to khai la Quyet toan thue TNCN mau 04/TNCN thi cho hien check quet du lieu bang ke
    ElseIf (verToKhai = 0 And LoaiTk = "17") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.Value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    Else
        frmSystem.chkQuetBangKe.Value = False
        frmSystem.chkQuetBangKe.Visible = False
    End If
    
    mOnLoad = False
    LoadForm = True
    
    frmSystem.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    Exit Function
ErrHandle:
    
    SaveErrorLog Me.Name, "LoadForm", Err.Number, Err.Description
End Function

'****************************
'Description: IsCompleteData function check whether barcode data is complete
'Author:ThanhDX
'Date:19/11/2005
'Input:
'Output:
'       strData: complete data string.
'Return: true if the string is complete data
'        false if otherwise
'****************************
Private Function IsCompleteData(ByRef strData As String) As Boolean
    Dim blnReturn As Boolean
    Dim intCtrl As Integer, intCount As Integer
    Dim strTemp As String

On Error GoTo ErrHandle
    blnReturn = True
    strTemp = arrStrElements(0)
    
    '*********************************
    'Date: 10/04/06
    ' Check Version
    If Left$(strTemp, 1) = "0" Then
    'Version 1.1.0 and 1.0.0
    Else
    'Version 1.2.0 and later
        'Remove 6 character of printting session
        strTemp = Mid$(strTemp, 1, Len(strTemp) - 6)
    End If
    '*********************************
    
    For intCtrl = 1 To UBound(arrStrElements())
        If Trim(arrStrElements(intCtrl)) = vbNullString Then
            blnReturn = False
        Else
            strTemp = strTemp & arrStrElements(intCtrl)
            intCount = intCount + 1
        End If
    Next intCtrl
    
    'Get all of data string
    If blnReturn Then
        strData = strTemp
    End If
    
    If intCount > ProgressBar1.Value Then
        ProgressBar1.Value = ProgressBar1.Value + 1
        lblFilePath.caption = ProgressBar1.Value & " / " & ProgressBar1.max
    End If
    
    IsCompleteData = blnReturn
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "IsCompleteData", Err.Number, Err.Description
End Function

'****************************
'Description: IsDifferent function compare two string
'Author:ThanhDX
'Date:23/11/2005
'Input:
'       strValue1: The first string.
'       strValue2: The second string.
'Output:
'Return: 0 if two string is the same
'        1 if two string not is the same
'****************************
Private Function IsDifferent(ByVal strValue1 As String, ByVal strValue2 As String) As Integer
    Dim intReturn As Integer

On Error GoTo ErrHandle
    If strValue1 = strValue2 Then Exit Function
    
    'If Mid$(strValue1, 3) <> Mid$(strValue2, 3) Or Left$(strValue1, 2) <> Left$(strValue2, 2) Then
        intReturn = 1 'Another tax report
        'GoTo exitFunction
    'End If

'exitFunction:
    IsDifferent = intReturn
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "IsDifferent", Err.Number, Err.Description
End Function

'****************************
'Description: GetSheetDatas function divide data string into sheet datas.
'Author:ThanhDX
'Date:23/11/2005
'Input:strBarcodeData: Data string.
'Output:
'Return: array of data sheets.
'****************************
Private Function GetSheetDatas(ByVal strBarcodeData As String) As String()
    Dim arrStrData() As String ', strSheetId As String , strTemp As String
    Dim intIndex As Integer, intLoc1 As Long, intLoc2 As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    
On Error GoTo ErrHandle
    For Each xmlNode In TAX_Utilities_Srv_New.NodeValidity.childNodes
        SetAttribute xmlNode, "Active", "0"
    Next
    
    ReDim arrStrData(0)
    
    For Each xmlNode In TAX_Utilities_Srv_New.NodeValidity.childNodes
         Dim i As Integer
        intLoc1 = InStr(1, strBarcodeData, "<S" & GetAttribute(xmlNode, "ID") & ">")
        i = Len(GetAttribute(xmlNode, "ID"))
        If intLoc1 = 0 Then
            intIndex = intIndex + 1
            ReDim Preserve arrStrData(intIndex)
        Else
            intLoc2 = InStr(1, strBarcodeData, "</S" & GetAttribute(xmlNode, "ID") & ">")
            If intLoc2 > intLoc1 Then
                SetAttribute xmlNode, "Active", "1"
                intIndex = intIndex + 1
                ReDim Preserve arrStrData(intIndex)
                arrStrData(intIndex) = Mid$(strBarcodeData, intLoc1, intLoc2 + i + 3)
                strBarcodeData = Replace(strBarcodeData, arrStrData(intIndex), "")
            End If
        End If
    Next
    
    If strBarcodeData = "" Then
        If UBound(arrStrData) < TAX_Utilities_Srv_New.NodeValidity.childNodes.length Then
            ReDim Preserve arrStrData(TAX_Utilities_Srv_New.NodeValidity.childNodes.length)
        End If
    Else
        ReDim arrStrData(0)
    End If
    
    GetSheetDatas = arrStrData()
    Set xmlNode = Nothing

    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetSheetDatas", Err.Number, Err.Description
End Function

'****************************
'Description: GetTaxInfo function get data from DB
'Author:ThanhDX
'Date:23/11/2005
'Input:
'       strTaxIDString: Id string
'Output:
'Return: a Recordset contain query data.
'****************************
Private Function GetTaxInfo(ByVal strTaxIDString As String, _
                            ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL   As String
    
    On Error GoTo ErrHandle

    'Lay tu webservices cua ESB tra ve

    Dim paXmlDoc   As New MSXML.DOMDocument
    Dim sTranCode  As String
    Dim sTaxOffice As String
    Dim sUrlWs     As String
    Dim soapAct    As String
    Dim fldName    As String
    Dim fldValue   As String
    Dim xmlRequest As String
    
    
    Set xmlResultNNT = New MSXML.DOMDocument
    Dim strResultNNT As String
    
'    'Du lieu gia lap de test
'        Set xmlResultNNT = LoadXmlTemp("ResultNNTFromESB")
'        strResultNNT = "sdfsfds"
    
    If (strTaxIDString <> "" Or strTaxIDString <> vbNullString) Then
        Dim cfigXml As New MSXML.DOMDocument
        Set cfigXml = LoadConfig()

        strMaNNT = strTaxIDString

        paXmlDoc.Load GetAbsolutePath("..\InterfaceTemplates\xml\paramNntInESB.xml")
        sUrlWs = cfigXml.getElementsByTagName("WsUrlNNT")(0).Text
        soapAct = cfigXml.getElementsByTagName("SoapActionNNT")(0).Text
        xmlRequest = cfigXml.getElementsByTagName("XmlRequestNNT")(0).lastChild.xml
        sTranCode = cfigXml.getElementsByTagName("TRAN_CODE")(0).Text
        fldName = cfigXml.getElementsByTagName("ParamNameNNT")(0).Text

        'Set value config to file param NNT
        paXmlDoc.getElementsByTagName("tin_nnt")(0).Text = strTaxIDString

        paXmlDoc.getElementsByTagName("VERSION")(0).Text = cfigXml.getElementsByTagName("VERSION")(0).Text
        paXmlDoc.getElementsByTagName("SENDER_CODE")(0).Text = cfigXml.getElementsByTagName("SENDER_CODE")(0).Text
        paXmlDoc.getElementsByTagName("SENDER_NAME")(0).Text = cfigXml.getElementsByTagName("SENDER_NAME")(0).Text
        paXmlDoc.getElementsByTagName("RECEIVER_CODE")(0).Text = cfigXml.getElementsByTagName("RECEIVER_CODE")(0).Text
        paXmlDoc.getElementsByTagName("RECEIVER_NAME")(0).Text = cfigXml.getElementsByTagName("RECEIVER_NAME")(0).Text

        paXmlDoc.getElementsByTagName("ORIGINAL_CODE")(0).Text = cfigXml.getElementsByTagName("ORIGINAL_CODE")(0).Text
        paXmlDoc.getElementsByTagName("ORIGINAL_NAME")(0).Text = cfigXml.getElementsByTagName("ORIGINAL_NAME")(0).Text

        paXmlDoc.getElementsByTagName("MSG_ID")(0).Text = cfigXml.getElementsByTagName("SENDER_CODE")(0).Text & GenerateCodeByNow() '& GetGUID()
        
        
        paXmlDoc.getElementsByTagName("SEND_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")
        paXmlDoc.getElementsByTagName("ORIGINAL_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")

        fldValue = paXmlDoc.xml
        fldValue = ChangeTagASSCII(fldValue, True)

        If (Dir("c:\TempXML\", vbDirectory) = "") Then
            MkDir "c:\TempXML\"
        End If

        Dim sParamNNT As String

        sParamNNT = "c:\TempXML\" & "paramNNT.xml"
        paXmlDoc.save sParamNNT

'        'Return value from ESB
        strResultNNT = DataFromESB(sUrlWs, soapAct, xmlRequest, fldName, fldValue)

        strResultNNT = ChangeTagASSCII(strResultNNT, False)
        xmlResultNNT.loadXML strResultNNT
    Else
        Set rsReturn = Nothing
        blnSuccess = False
        MessageBox "0138", msOKOnly, miCriticalError
        Exit Function
    End If

    If (strResultNNT = "" Or strResultNNT = vbNullString Or Not xmlResultNNT.hasChildNodes) Then
        If (MessageBox("0135", msYesNo, miCriticalError) = mrNo) Then
            Set rsReturn = Nothing
            blnSuccess = False
            Exit Function
        End If

    Else
        Dim sResultNNT As String

        sResultNNT = "c:\TempXML\" & "ResultNNT.xml"
        xmlResultNNT.save sResultNNT
    
        Dim Err_des As String
        If (xmlResultNNT.getElementsByTagName("ERROR_DESC").length > 0) Then
            Err_des = xmlResultNNT.getElementsByTagName("ERROR_DESC")(0).Text
        End If
        If (Err_des <> "") Then
                MessageBox "0139", msOKOnly, miCriticalError
                Set rsReturn = Nothing
                blnSuccess = False
                Exit Function

        Else
'            If (InStr(xmlResultNNT.xml, "fault_code") > 0) Then
'                   If (MessageBox("0142", msYesNo, miCriticalError) = mrNo) Then
'                    Set rsReturn = Nothing
'                    blnSuccess = False
'                    Exit Function
'                    End If
'            End If
            
            If ((InStr(xmlResultNNT.xml, "fault_code") > 0) Or (InStr(xmlResultNNT.xml, "MaSoThue") <= 0)) Then
                If (MessageBox("0135", msYesNo, miCriticalError) = mrNo) Then
                    Set rsReturn = Nothing
                    blnSuccess = False
                    Exit Function
                End If
            End If
        End If
    End If

    rsReturn.Fields.Append "trang_thai", adChar, 2, adFldUpdatable
    rsReturn.Fields.Append "tin", adVarChar, 14, adFldUpdatable
    rsReturn.Fields.Append "ten_dtnt", adVarWChar, 100, adFldUpdatable
    rsReturn.Fields.Append "dia_chi", adVarWChar, 60, adFldUpdatable
    rsReturn.Fields.Append "dien_thoai", adVarWChar, 20, adFldUpdatable

    rsReturn.Fields.Append "fax", adVarWChar, 20, adFldUpdatable
    rsReturn.Fields.Append "mail", adVarWChar, 30, adFldUpdatable
    rsReturn.Fields.Append "ky_lapbo", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_nop", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_nhap", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_tchinh", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_kdoanh", adVarWChar, 50, adFldUpdatable

    rsReturn.Open
    rsReturn.AddNew

    If ((strResultNNT <> "" And xmlResultNNT.hasChildNodes And (InStr(xmlResultNNT.xml, "MaSoThue") > 0)) And Err_des = "") Then
        'xmlResultNNT.loadXML TAX_Utilities_Srv_New.Convert(xmlResultNNT.xml, VISCII, UNICODE)
        rsReturn!trang_thai = GetStringByLength(xmlResultNNT.getElementsByTagName("TrangThaiHoatDong")(0).Text, 2)
        rsReturn!ten_dtnt = TAX_Utilities_Srv_New.Convert(GetStringByLength(xmlResultNNT.getElementsByTagName("TenNNT")(0).Text, 100), UNICODE, TCVN)
        rsReturn!dia_chi = TAX_Utilities_Srv_New.Convert(GetStringByLength(xmlResultNNT.getElementsByTagName("DiaChi")(0).Text, 60), UNICODE, TCVN)
        
        rsReturn!Dien_thoai = GetStringByLength(xmlResultNNT.getElementsByTagName("DienThoai")(0).Text, 20)
        rsReturn!mail = GetStringByLength(xmlResultNNT.getElementsByTagName("Email")(0).Text, 30)
        rsReturn!Fax = GetStringByLength(xmlResultNNT.getElementsByTagName("Fax")(0).Text, 20)
        rsReturn!ngay_tchinh = "" 'GetStringByLength(xmlResultNNT.getElementsByTagName("START_DATE")(0).Text, 50)
        rsReturn!ngay_kdoanh = GetStringByLength(xmlResultNNT.getElementsByTagName("NgayBatDauKinhDoanh")(0).Text, 50)
    End If
    
    rsReturn!TIN = strTaxIDString
    rsReturn!ky_lapbo = IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    rsReturn!ngay_nop = IIf(DateTime.Day(DateTime.Now) < 10, "0" & DateTime.Day(DateTime.Now), CStr(DateTime.Day(DateTime.Now))) & "/" & IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    rsReturn!ngay_nhap = IIf(DateTime.Day(DateTime.Now) < 10, "0" & DateTime.Day(DateTime.Now), CStr(DateTime.Day(DateTime.Now))) & "/" & IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    
    rsReturn!ngay_kdoanh = ""
    rsReturn.Update
    
    
    Set GetTaxInfo = rsReturn

    Set rsReturn = Nothing
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetTaxInfo", Err.Number, Err.Description

    If Err.Number = -2147467259 Then MessageBox "0063", msOKOnly, miCriticalError
End Function

' Lay thong tin DL thue 05072011
Private Function GetTaxDLInfo(ByVal strTaxIDString As String, _
                              ByVal strTaxIDDLString As String, _
                              ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL   As String
    
    On Error GoTo ErrHandle

    Dim paXmlDoc   As New MSXML.DOMDocument
    Dim sTranCode  As String
    Dim sTaxOffice As String
    Dim sUrlWs     As String
    Dim soapAct    As String
    Dim fldName    As String
    Dim fldValue   As String
    Dim xmlRequest As String
    
    Set xmlResultDLT = New MSXML.DOMDocument
    Dim strResultDLT As String
    
    
'    'Du lieu gia lap de test
'    Set xmlResultDLT = LoadXmlTemp("ResultDLTFromESB")
'    strResultDLT = "sdfsfds"
    
    'Neu khong co thong tin NNT thi exit luon
    If (strTaxIDString = "" Or strTaxIDString = vbNullString) Then
        Set rsReturn = Nothing
        blnSuccess = False
        Exit Function
    End If
    

    If (strTaxIDDLString <> "" And strTaxIDDLString <> vbNullString) Then
        Dim cfigXml As New MSXML.DOMDocument
        Set cfigXml = LoadConfig()
        
        strMaDLT = strTaxIDDLString

        paXmlDoc.Load GetAbsolutePath("..\InterfaceTemplates\xml\paramDltInESB.xml")
        sUrlWs = cfigXml.getElementsByTagName("WsUrlDLT")(0).Text
        soapAct = cfigXml.getElementsByTagName("SoapActionDLT")(0).Text
        xmlRequest = cfigXml.getElementsByTagName("XmlRequestDLT")(0).lastChild.xml
        sTranCode = cfigXml.getElementsByTagName("TRAN_CODE")(0).Text
        fldName = cfigXml.getElementsByTagName("ParamNameDLT")(0).Text

        'Set value config to file param DLT
        paXmlDoc.getElementsByTagName("tin_dlt")(0).Text = strTaxIDDLString
        paXmlDoc.getElementsByTagName("tin_nnt")(0).Text = strTaxIDString

        paXmlDoc.getElementsByTagName("VERSION")(0).Text = cfigXml.getElementsByTagName("VERSION")(0).Text
        paXmlDoc.getElementsByTagName("SENDER_CODE")(0).Text = cfigXml.getElementsByTagName("SENDER_CODE")(0).Text
        paXmlDoc.getElementsByTagName("SENDER_NAME")(0).Text = cfigXml.getElementsByTagName("SENDER_NAME")(0).Text
        paXmlDoc.getElementsByTagName("RECEIVER_CODE")(0).Text = cfigXml.getElementsByTagName("RECEIVER_CODE")(0).Text
        paXmlDoc.getElementsByTagName("RECEIVER_NAME")(0).Text = cfigXml.getElementsByTagName("RECEIVER_NAME")(0).Text

        paXmlDoc.getElementsByTagName("ORIGINAL_CODE")(0).Text = cfigXml.getElementsByTagName("ORIGINAL_CODE")(0).Text
        paXmlDoc.getElementsByTagName("ORIGINAL_NAME")(0).Text = cfigXml.getElementsByTagName("ORIGINAL_NAME")(0).Text

        paXmlDoc.getElementsByTagName("MSG_ID")(0).Text = cfigXml.getElementsByTagName("SENDER_CODE")(0).Text & GenerateCodeByNow() '& GetGUID()
        paXmlDoc.getElementsByTagName("SEND_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")
        paXmlDoc.getElementsByTagName("ORIGINAL_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")

        fldValue = paXmlDoc.xml
        fldValue = ChangeTagASSCII(fldValue, True)

        If (Dir("c:\TempXML\", vbDirectory) = "") Then
            MkDir "c:\TempXML\"
        End If

        Dim sParamDLT As String

        sParamDLT = "c:\TempXML\" & "paramDLT.xml"
        paXmlDoc.save sParamDLT

        'Return value from ESB
        strResultDLT = DataFromESB(sUrlWs, soapAct, xmlRequest, fldName, fldValue)

        strResultDLT = ChangeTagASSCII(strResultDLT, False)
        xmlResultDLT.loadXML strResultDLT
    End If
    
    If strTaxIDDLString <> "" And strTaxIDDLString <> vbNullString Then
        If (strResultDLT = "" Or strResultDLT = vbNullString Or Not xmlResultDLT.hasChildNodes) Then
            If (MessageBox("0136", msYesNo, miCriticalError) = mrNo) Then
                Set rsReturn = Nothing
                blnSuccess = False
                Exit Function
            End If
    
        Else
                Dim sResultDLT As String
    
            sResultDLT = "c:\TempXML\" & "ResultDLT.xml"
            xmlResultDLT.save sResultDLT
        
            Dim Err_des As String
            If (xmlResultDLT.getElementsByTagName("ERROR_DESC").length > 0) Then
                Err_des = xmlResultDLT.getElementsByTagName("ERROR_DESC")(0).Text
            End If
            If (Err_des <> "") Then
                If (MessageBox("0139", msYesNo, miCriticalError) = mrNo) Then
                    Set rsReturn = Nothing
                    blnSuccess = False
                    Exit Function
                End If
            Else
    '            If (InStr(xmlResultDLT.xml, "NORM_NAME") <= 0) Then
    '                If (MessageBox("0136", msYesNo, miCriticalError) = mrNo) Then
    '                    Set rsReturn = Nothing
    '                    blnSuccess = False
    '                    Exit Function
    '                End If
    '            End If
    
                If (InStr(xmlResultDLT.xml, "fault_code") > 0) Then
                       If (MessageBox("0141", msYesNo, miCriticalError) = mrNo) Then
                        Set rsReturn = Nothing
                        blnSuccess = False
                        Exit Function
                        End If
                End If
                If (xmlResultDLT.getElementsByTagName("TrangThaiHoatDong").length > 0) Then
                    If (xmlResultDLT.getElementsByTagName("TrangThaiHoatDong")(0).Text = "01") Then
                            Set rsReturn = Nothing
                            blnSuccess = False
                            Exit Function
                    End If
                End If
            End If
            
        End If
    End If
    
    rsReturn.Fields.Append "repr_name", adVarWChar, 200, adFldUpdatable
    rsReturn.Fields.Append "repr_addr", adVarWChar, 200, adFldUpdatable
    rsReturn.Fields.Append "repr_tell", adVarWChar, 30, adFldUpdatable
    rsReturn.Fields.Append "repr_fax", adVarWChar, 30, adFldUpdatable
    rsReturn.Fields.Append "repr_email", adVarWChar, 60, adFldUpdatable
    rsReturn.Fields.Append "repr_cont_number", adVarWChar, 30, adFldUpdatable
    rsReturn.Fields.Append "repr_cont_date", adVarWChar, 50, adFldUpdatable
    
    rsReturn.Open
    rsReturn.AddNew
    
    If (strResultDLT <> "" And xmlResultDLT.hasChildNodes And (InStr(xmlResultDLT.xml, "MaSoThue") > 0) And Err_des = "") Then
        'xmlResultDLT.loadXML TAX_Utilities_Srv_New.Convert(xmlResultDLT.xml, VISCII, UNICODE)

        rsReturn!repr_name = TAX_Utilities_Srv_New.Convert(xmlResultDLT.getElementsByTagName("TenNNT")(0).Text, UNICODE, TCVN)
        rsReturn!repr_addr = TAX_Utilities_Srv_New.Convert(xmlResultDLT.getElementsByTagName("DiaChi")(0).Text, UNICODE, TCVN)

        rsReturn!repr_tell = xmlResultDLT.getElementsByTagName("DienThoai")(0).Text
        rsReturn!repr_fax = xmlResultDLT.getElementsByTagName("Fax")(0).Text
        rsReturn!repr_email = xmlResultDLT.getElementsByTagName("Email")(0).Text
        rsReturn!repr_cont_number = xmlResultDLT.getElementsByTagName("SoHopDong")(0).Text
        rsReturn!repr_cont_date = xmlResultDLT.getElementsByTagName("NgayHopDong")(0).Text
    End If

    rsReturn.Update
    Set GetTaxDLInfo = rsReturn

    Set rsReturn = Nothing
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetTaxDLInfo", Err.Number, Err.Description

    If Err.Number = -2147467259 Then MessageBox "0063", msOKOnly, miCriticalError
End Function


' nvhai
' Lay ve thong tin cua doi tuong NT nhung khong check CQT
' Phuc vu viec quet cac BCTC cua chi Cuc nhung quet tren Cuc
' 09-06-2010
' begin
'Input:
'       strTaxIDString: Id string
'Output:
'Return: a Recordset contain query data.
'****************************
Private Function GetTaxInfoBCTC(ByVal strTaxIDString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    Dim paXmlDoc As New MSXML.DOMDocument
    Dim sTranCode As String
    Dim sTaxOffice As String
    Dim sUrlWs As String
    Dim soapAct As String
    Dim fldName As String
    Dim fldValue As String
    Dim xmlRequest As String
    
    Set xmlResultNNT = New MSXML.DOMDocument
    Dim strResultNNT As String

   'Du lieu gia lap de test
    Set xmlResultNNT = LoadXmlTemp("ResultNNTFromESB")
    strResultNNT = "test"

    If (strTaxIDString <> "" Or strTaxIDString <> vbNullString) Then
        Dim cfigXml As New MSXML.DOMDocument
        Set cfigXml = LoadConfig()
        strMaNNT = strTaxIDString

        paXmlDoc.Load GetAbsolutePath("..\InterfaceTemplates\xml\paramNntInESB.xml")
        sUrlWs = cfigXml.getElementsByTagName("WsUrlNNT")(0).Text
        soapAct = cfigXml.getElementsByTagName("SoapActionNNT")(0).Text
        xmlRequest = cfigXml.getElementsByTagName("XmlRequestNNT")(0).lastChild.xml
        sTranCode = cfigXml.getElementsByTagName("TRAN_CODE")(0).Text
        fldName = cfigXml.getElementsByTagName("ParamNameNNT")(0).Text

        'Set value config to file param NNT
        paXmlDoc.getElementsByTagName("tin_nnt")(0).Text = strTaxIDString

        paXmlDoc.getElementsByTagName("VERSION")(0).Text = cfigXml.getElementsByTagName("VERSION")(0).Text
        paXmlDoc.getElementsByTagName("SENDER_CODE")(0).Text = cfigXml.getElementsByTagName("SENDER_CODE")(0).Text
        paXmlDoc.getElementsByTagName("SENDER_NAME")(0).Text = cfigXml.getElementsByTagName("SENDER_NAME")(0).Text
        paXmlDoc.getElementsByTagName("RECEIVER_CODE")(0).Text = cfigXml.getElementsByTagName("RECEIVER_CODE")(0).Text
        paXmlDoc.getElementsByTagName("RECEIVER_NAME")(0).Text = cfigXml.getElementsByTagName("RECEIVER_NAME")(0).Text

        paXmlDoc.getElementsByTagName("ORIGINAL_CODE")(0).Text = cfigXml.getElementsByTagName("ORIGINAL_CODE")(0).Text
        paXmlDoc.getElementsByTagName("ORIGINAL_NAME")(0).Text = cfigXml.getElementsByTagName("ORIGINAL_NAME")(0).Text

        paXmlDoc.getElementsByTagName("MSG_ID")(0).Text = cfigXml.getElementsByTagName("SENDER_CODE")(0).Text & GenerateCodeByNow() 'GetGUID()
        paXmlDoc.getElementsByTagName("SEND_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")
        paXmlDoc.getElementsByTagName("ORIGINAL_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")

        fldValue = paXmlDoc.xml
        fldValue = ChangeTagASSCII(fldValue, True)

        If (Dir("c:\TempXML\", vbDirectory) = "") Then
            MkDir "c:\TempXML\"
        End If

        Dim sParamNNT As String

        sParamNNT = "c:\TempXML\" & "paramNNT.xml"
        paXmlDoc.save sParamNNT

'        'Return value from ESB
'        strResultNNT = DataFromESB(sUrlWs, soapAct, xmlRequest, fldName, fldValue)

        strResultNNT = ChangeTagASSCII(strResultNNT, False)
        xmlResultNNT.loadXML strResultNNT
    Else
        Set rsReturn = Nothing
        blnSuccess = False
        MessageBox "0138", msOKOnly, miCriticalError
        Exit Function
    End If

    If (strResultNNT = "" Or strResultNNT = vbNullString Or Not xmlResultNNT.hasChildNodes) Then
        If (MessageBox("0135", msYesNo, miCriticalError) = mrNo) Then
            Set rsReturn = Nothing
            blnSuccess = False
            Exit Function
        End If

    Else
        Dim sResultNNT As String

        sResultNNT = "c:\TempXML\" & "ResultNNT.xml"
        xmlResultNNT.save sResultNNT
    
        Dim Err_des As String
        If (xmlResultNNT.getElementsByTagName("ERROR_DESC").length > 0) Then
            Err_des = xmlResultNNT.getElementsByTagName("ERROR_DESC")(0).Text
        End If
        
        If (Err_des <> "") Then
                MessageBox "0139", msOKOnly, miCriticalError
                Set rsReturn = Nothing
                blnSuccess = False
                Exit Function
            

        Else
'            If (InStr(xmlResultNNT.xml, "fault_code") > 0) Then
'                 If (MessageBox("0142", msYesNo, miCriticalError) = mrNo) Then
'                    Set rsReturn = Nothing
'                    blnSuccess = False
'                    Exit Function
'                End If
'            End If
            
            If ((InStr(xmlResultNNT.xml, "fault_code") > 0) Or (InStr(xmlResultNNT.xml, "MaSoThue") <= 0)) Then
                If (MessageBox("0135", msYesNo, miCriticalError) = mrNo) Then
                    Set rsReturn = Nothing
                    blnSuccess = False
                    Exit Function
                End If
            End If
        End If
    End If

    rsReturn.Fields.Append "trang_thai", adChar, 2, adFldUpdatable
    rsReturn.Fields.Append "tin", adVarChar, 14, adFldUpdatable
    rsReturn.Fields.Append "ten_dtnt", adVarWChar, 100, adFldUpdatable
    rsReturn.Fields.Append "dia_chi", adVarWChar, 60, adFldUpdatable
    rsReturn.Fields.Append "dien_thoai", adVarWChar, 20, adFldUpdatable

    rsReturn.Fields.Append "fax", adVarWChar, 20, adFldUpdatable
    rsReturn.Fields.Append "mail", adVarWChar, 30, adFldUpdatable
    rsReturn.Fields.Append "ky_lapbo", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_nop", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_nhap", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_tchinh", adVarWChar, 50, adFldUpdatable
    rsReturn.Fields.Append "ngay_kdoanh", adVarWChar, 50, adFldUpdatable

    rsReturn.Open
    rsReturn.AddNew
    
   If ((strResultNNT <> "" And xmlResultNNT.hasChildNodes And (InStr(xmlResultNNT.xml, "MaSoThue") > 0)) And Err_des = "") Then
        'xmlResultNNT.loadXML TAX_Utilities_Srv_New.Convert(xmlResultNNT.xml, VISCII, UNICODE)
        rsReturn!trang_thai = GetStringByLength(xmlResultNNT.getElementsByTagName("TrangThaiHoatDong")(0).Text, 2)
        rsReturn!ten_dtnt = TAX_Utilities_Srv_New.Convert(GetStringByLength(xmlResultNNT.getElementsByTagName("TenNNT")(0).Text, 100), UNICODE, TCVN)
        rsReturn!dia_chi = TAX_Utilities_Srv_New.Convert(GetStringByLength(xmlResultNNT.getElementsByTagName("DiaChi")(0).Text, 60), UNICODE, TCVN)
        
        rsReturn!Dien_thoai = GetStringByLength(xmlResultNNT.getElementsByTagName("DienThoai")(0).Text, 20)
        rsReturn!mail = GetStringByLength(xmlResultNNT.getElementsByTagName("Email")(0).Text, 30)
        rsReturn!Fax = GetStringByLength(xmlResultNNT.getElementsByTagName("Fax")(0).Text, 20)
        rsReturn!ngay_tchinh = "" 'GetStringByLength(xmlResultNNT.getElementsByTagName("START_DATE")(0).Text, 50)
        rsReturn!ngay_kdoanh = GetStringByLength(xmlResultNNT.getElementsByTagName("NgayBatDauKinhDoanh")(0).Text, 50)
    End If
    
    rsReturn!TIN = strTaxIDString
    rsReturn!ky_lapbo = IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    rsReturn!ngay_nop = IIf(DateTime.Day(DateTime.Now) < 10, "0" & DateTime.Day(DateTime.Now), CStr(DateTime.Day(DateTime.Now))) & "/" & IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    rsReturn!ngay_nhap = IIf(DateTime.Day(DateTime.Now) < 10, "0" & DateTime.Day(DateTime.Now), CStr(DateTime.Day(DateTime.Now))) & "/" & IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
   
    rsReturn!TIN = strTaxIDString
    rsReturn!ky_lapbo = IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    rsReturn!ngay_nop = IIf(DateTime.Day(DateTime.Now) < 10, "0" & DateTime.Day(DateTime.Now), CStr(DateTime.Day(DateTime.Now))) & "/" & IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    rsReturn!ngay_nhap = IIf(DateTime.Day(DateTime.Now) < 10, "0" & DateTime.Day(DateTime.Now), CStr(DateTime.Day(DateTime.Now))) & "/" & IIf(DateTime.Month(DateTime.Now) < 10, "0" & DateTime.Month(DateTime.Now), CStr(DateTime.Month(DateTime.Now))) & "/" & CStr(DateTime.Year(DateTime.Now))
    
    rsReturn!ngay_kdoanh = ""
    rsReturn.Update
    
    
    Set GetTaxInfoBCTC = rsReturn

    Set rsReturn = Nothing
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetTaxInfoBCTC", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

'end

Private Sub GetPhongQuanLy(ByVal strTaxIDString As String)
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    Dim strPQLString As String
On Error GoTo ErrHandle

    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'    ' Lay ma phong quan ly cua MST
'    strSQL = "select ma_phong from rcv_v_dtnt where tin = '" & Trim(strTaxIDString) & "'"
'    Set rsReturn = clsDAO.Execute(strSQL)
'    If Not (rsReturn Is Nothing) And rsReturn.Fields.Count > 0 Then
'        strPQLString = Trim(rsReturn.Fields(0).Value)
'    End If
'
'
'    ' Get SQL statement from DOM
'    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlPhongQuanLy")
'
'    '*************************************
'    'Date: 30/05/06
'    strSQL = Replace$(strSQL, "MA_PQL", strPQLString)
'    '*************************************
''    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
''    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
'
'    Set rsReturn = clsDAO.Execute(strSQL)
'    If Not (rsReturn Is Nothing) And rsReturn.Fields.Count > 0 Then
'        strMaPhongQuanLy = rsReturn.Fields(0).Value
'        strTenPhongQuanLy = rsReturn.Fields(1).Value
'    End If
'
'
'    Set rsReturn = Nothing
    
    Exit Sub
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "GetPQL", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Sub


Private Function GetPhongXuLy(ByVal strPXLString As String, ByRef blnSuccess As Boolean) As Object
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'
'    ' Get SQL statement from DOM
'    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlPhongXuLy")
'
'    '*************************************
'    'Date: 30/05/06
'    strSQL = Replace$(strSQL, "MA_CQT", strTaxOfficeId)
'    '*************************************
''    strSQL = Replace(strSQL, "strTaxOfficeId", "'" & strTaxOfficeId & "'")
''    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
'
'    Set rsReturn = clsDAO.Execute(strSQL)
    
    rsReturn.Fields.Append "ten", adVarChar, 50, adFldUpdatable
    
    rsReturn.Fields.Append "ma_phong", adVarChar, 50, adFldUpdatable
    
    rsReturn.Open
    rsReturn.AddNew
    rsReturn!ten = "ten"
    rsReturn!ma_phong = "ma_phong"
    rsReturn.Update
    Set GetPhongXuLy = rsReturn
    
    Set rsReturn = Nothing
    
    'Connect DB success
    blnSuccess = True
    
    Exit Function
ErrHandle:
    'Connect DB fail
    blnSuccess = False
    SaveErrorLog Me.Name, "GetPXL", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

'****************************
'Description: SetupHeaderData procedure get data from DB
'             and fill data to screen and DOM.
'Author:ThanhDX
'Date:23/11/2005
'Input:rsTaxInfor: A recordset contain data which get from DB
'Output:
'Return:
'****************************
'Private Sub SetupHeaderData(ByRef rsTaxInfor As ADODB.Recordset)
'    Dim lIndex As Long, lCtrl As Long
'    Dim lCol As Long, lRow As Long
'
'On Error GoTo ErrHandle
'        fpSpread1.Sheet = lCtrl + 1
'        For lIndex = 1 To TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes.length
'            If lIndex < rsTaxInfor.Fields.Count Then
'                If Not rsTaxInfor.Fields(lIndex) = vbNullString Then
'                    TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex - 1) _
'                        .Attributes.getNamedItem("Value").nodeValue = TAX_Utilities_Srv_New.Convert(rsTaxInfor.Fields(lIndex).Value, TCVN, UNICODE)
'                    ParserCellID fpSpread1, GetAttribute(TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex - 1), "CellID"), lCol, lRow
'                    fpSpread1.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsTaxInfor.Fields(lIndex).Value, TCVN, UNICODE)
'                    fpSpread1.RowHeight(lRow) = fpSpread1.MaxTextRowHeight(lRow)
'                End If
'            Else
'                Exit For
'            End If
'        Next lIndex
'    Exit Sub
'ErrHandle:
'    SaveErrorLog Me.Name, "SetupHeaderData", Err.Number, Err.Description
Private Sub SetupHeaderData(arrStrHeaderData() As String)
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    
On Error GoTo ErrHandle
        fpSpread1.Sheet = lCtrl + 1
        For lIndex = 0 To UBound(arrStrHeaderData) 'TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes.length
            'If lIndex < UBound(arrStrHeaderData) Then
                If Not arrStrHeaderData(lIndex) = vbNullString Then
                    SetAttribute TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex) _
                        , "Value", TAX_Utilities_Srv_New.Convert(arrStrHeaderData(lIndex), TCVN, UNICODE)
                    ParserCellID fpSpread1, GetAttribute(TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex), "CellID"), lCol, lRow
                    fpSpread1.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(arrStrHeaderData(lIndex), TCVN, UNICODE)
                    fpSpread1.RowHeight(lRow) = fpSpread1.MaxTextRowHeight(lRow)
                End If
            'Else
                'Exit For
            'End If
'' Thay lai ham convert unicode
'                If Not arrStrHeaderData(lIndex) = vbNullString Then
'                    SetAttribute TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex) _
'                        , "Value", arrStrHeaderData(lIndex)
'                    ParserCellID fpSpread1, GetAttribute(TAX_Utilities_Srv_New.Data(lCtrl).getElementsByTagName("Section")(0).firstChild.childNodes(lIndex), "CellID"), lCol, lRow
'                    fpSpread1.SetText lCol, lRow, arrStrHeaderData(lIndex)
'                    fpSpread1.RowHeight(lRow) = fpSpread1.MaxTextRowHeight(lRow)
'                End If
        Next lIndex
        
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetupHeaderData", Err.Number, Err.Description
End Sub

Function GenerateSQL_Details(xmlDomData As MSXML.DOMDocument, strSQL_DTL As String, vHdrID As Variant, lPos As Long) As String
    Dim xmlListSection As MSXML.IXMLDOMNodeList
    Dim xmlNodeSection As MSXML.IXMLDOMNode
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim lPosition As Long, strCondition As String
    Dim i As Long, j As Long, strLoaiDL As String
    
On Error GoTo ErrHandle
    Set xmlListSection = xmlDomData.getElementsByTagName("Section")
    For Each xmlNodeSection In xmlListSection
        If Trim(xmlNodeSection.Attributes.getNamedItem("Dynamic").nodeValue) = "1" Then
            iRowID = 0
            For i = 0 To xmlNodeSection.childNodes.length - 1
                iRowID = iRowID + 1
                For j = 0 To xmlNodeSection.childNodes(i).childNodes.length - 1
                    Set xmlAttribute = xmlDomData.createAttribute("RowID")
                    xmlAttribute.Value = iRowID
                    Set xmlNode = xmlNodeSection.childNodes(i).childNodes(j).Attributes.setNamedItem(xmlAttribute)
                    Set xmlAttribute = Nothing
                Next
            Next
        End If
    Next
        
    strLoaiDL = Trim(TAX_Utilities_Srv_New.NodeValidity.childNodes(lPos).Attributes.getNamedItem("DataFile").nodeValue)
    Set xmlList = xmlDomData.getElementsByTagName("Cell")
    If xmlList.length > 0 Then GenerateSQL_Details = "begin"
    ' Them tham so nay de tinh cu 50 dong se bat dau ghi thanh mot block du lieu
    Dim rC As Integer
    For Each xmlNode In xmlList
        If Not xmlNode.Attributes.getNamedItem("MCT") Is Nothing Then
             If Trim(xmlNode.Attributes.getNamedItem("MCT").nodeValue) <> "" Then
                rC = rC + 1
                strSQL = strSQL_DTL
                strSQL = strSQL & "'" & vHdrID & "',"
                strSQL = strSQL & "'" & strLoaiDL & "',"
                strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("MCT").nodeValue & "',"
                strSQL = strSQL & "'" & Trim(Replace(TAX_Utilities_Srv_New.Convert(xmlNode.Attributes.getNamedItem("Value").nodeValue, UNICODE, TCVN), "'", "''")) & "',"
                If Not xmlNode.Attributes.getNamedItem("RowID") Is Nothing Then
                    strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("RowID").nodeValue & "');"
                Else
                    'strSQL = strSQL & "'" & xmlNode.Attributes.getNamedItem("MCT").nodeValue & "');"
                    strSQL = strSQL & "'');"
                End If
                GenerateSQL_Details = GenerateSQL_Details & vbCrLf & strSQL
                If rC = 1 Then
                    'clsDAO.BeginTrans
                ElseIf rC = 30 Then
                    rC = 0
                    GenerateSQL_Details = GenerateSQL_Details & vbCrLf & "end;"
                    clsDAO.Execute GenerateSQL_Details
                    'clsDAO.CommitTrans
                    GenerateSQL_Details = vbNullString
                    GenerateSQL_Details = "begin"
                End If
             End If
        End If
    Next
    If Trim(GenerateSQL_Details) <> "begin" Then
        GenerateSQL_Details = GenerateSQL_Details & vbCrLf & "end;"
        clsDAO.Execute GenerateSQL_Details
        'clsDAO.CommitTrans
    End If
    Set xmlDomData = Nothing
    Set xmlList = Nothing
    Set xmlListSection = Nothing
    Exit Function

ErrHandle:
    SaveErrorLog Me.Name, "GenerateSQL_Details", Err.Number, Err.Description
    Err.Raise Err.Number
End Function

''' CheckValidData description
''' Check all formula in last sheet, if error put the notetext into cellnode
''' No parameter
''' Return True if no error checking
''' Return False if one or more error occur
Private Function CheckValidData() As Boolean
    
    Dim i As Long
    Dim strCellString As String
    
    Dim vFunction As Variant, vCell As Variant
    Dim vMsg As Variant, vWarning As Variant
    Dim vOrder As Variant, vFormulaFunc As Variant
    Dim cOrder As New Collection
    
On Error GoTo ErrHandle

    CheckValidData = True
    
    With fpSpread1
        .ReDraw = False
        If .SheetCount = 1 Then Exit Function
        .Sheet = mHeaderSheet
        
        For i = 12 To .MaxRows
            .Sheet = mHeaderSheet
            .Col = .ColLetterToNumber("B")
            .Row = i
            If .Formula <> vbNullString Then
                .Col = .Col + 1 'Column B
                strCellString = .Formula
                If Trim(strCellString) <> vbNullString Then SetCellNote strCellString, .BackColor, ""
            End If
        Next
        
        'set error note for cell
        If Not objTaxBusiness Is Nothing Then
            CheckValidData = objTaxBusiness.CheckValidData
        End If
        
        .Sheet = mHeaderSheet
        For i = 12 To .MaxRows
            .Sheet = mHeaderSheet
            .Col = 2
            .Row = i
            vFormulaFunc = .Formula
            If Trim(.Text) <> "" Then
                .GetText .ColLetterToNumber("B"), i, vFunction
                .GetText .ColLetterToNumber("E"), i, vMsg
                .GetText .ColLetterToNumber("S"), i, vWarning
                .GetText .ColLetterToNumber("T"), i, vOrder
                .Col = .Col + 1
                vCell = .Formula
                If vFormulaFunc <> vbNullString Then
                    If Val(vFunction) <> 1 Then
                        SetCellNote vCell, .BackColor, "> " & vMsg
                        If Trim(vCell) <> "" Then cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                        If UCase(Trim(vWarning)) = "Y" Then CheckValidData = False
                    End If
                Else 'Dynamic
                    If Val(vFunction) <> 1 Then
                        If Trim(vCell) <> "" Then cOrder.Add CStr(vOrder) & "[]" & CStr(vCell)
                        If UCase(Trim(vWarning)) = "Y" Then CheckValidData = False
                    End If
                End If
            End If
        Next
        
        'focus on the first error cell
        Dim min As Integer, X As Long, strCell As String
        Dim lSheet As Long, lCol As Long, lRow As Long
        
        
        If cOrder.Count > 0 Then
            min = Val(Left(cOrder(1), InStr(cOrder(1), "[]")))
            strCell = Right(cOrder(1), Len(cOrder(1)) - InStr(cOrder(1), "[]") - 1)
            For i = 2 To cOrder.Count
                X = Val(Left(cOrder(i), InStr(cOrder(i), "[]")))
                If min >= X Then
                    min = X
                    strCell = Right(cOrder(i), Len(cOrder(i)) - InStr(cOrder(i), "[]") - 1)
                End If
            Next
            'focus cell here
            getCellPosition strCell, lSheet, lCol, lRow
            .SetFocus
            .ActiveSheet = lSheet
            .SetActiveCell lCol, lRow
        End If
        
        .ReDraw = True
    End With
    Exit Function
    
ErrHandle:
    SaveErrorLog Me.Name, "CheckValidData", Err.Number, Err.Description
End Function


''' get Sheet, Col, Row from Cell Formula
'''Parameter: Cell Formula string
'''Parameter: sheet integer
'''parameter: Col integer
'''parameter: Row integer
Private Sub getCellPosition(pCellString As String, lSheet As Long, lCol As Long, lRow As Long)
        
    Dim lAnchor As Long
    Dim lSheetName As String, lCellString As String, lStringTemp As String
    Dim i As Long
    
On Error GoTo ErrHandle
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For i = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, i)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, i)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For i = 1 To .SheetCount
            .Sheet = i
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
                lSheet = i
                Exit For
            End If
        Next
    End With
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "getCellPosition", Err.Number, Err.Description
End Sub

''' SetCellNote description
''' Set CellNote for error cell
''' Parser pCellString (containt sheetname and cellID)
''' Parameter1 pCellString  : containt sheetname and cellID
''' Parameter2 pNoteText    : the string input into cellnote
Private Sub SetCellNote(ByVal pCellString As String, ByVal lNoErrColor As Long, ByVal pNoteText As String)
    
    Dim lAnchor As Long
    Dim lSheetName As String, lCellString As String, lStringTemp As String
    Dim lCol As Long, lRow As Long, i As Long
    
On Error GoTo ErrHandle
    ' Get anchor of character "!"
    lAnchor = InStr(1, pCellString, "!", vbTextCompare)
    ' Save sheet name to variable
    lSheetName = Left(pCellString, lAnchor - 1)
    ' Save cell string name to variable
    lCellString = Right(pCellString, Len(pCellString) - lAnchor)
    For i = 1 To Len(lCellString)
        If IsNumeric(Left(lCellString, i)) = False Then
            ' Aphabe charater
            lStringTemp = lStringTemp & Left(lCellString, i)
            lCellString = Right(lCellString, Len(lCellString) - 1)
        Else
            ' Numeric charater
            lRow = Val(lCellString)
            Exit For
        End If
    Next
    lCol = fpSpread1.ColLetterToNumber(lStringTemp)
    
    With fpSpread1
        For i = 1 To .SheetCount
            .Sheet = i
            If "'" & UCase(.SheetName) & "'" = UCase(lSheetName) Then
                ' Set Note text for error cell in error sheet
                .Col = lCol
                .Row = lRow
                
                If Trim(pNoteText) = "" Then
                    .CellNote = ""
                ElseIf Trim(.CellNote) = "" Then
                    .CellNote = pNoteText
                Else
                    .CellNote = .CellNote & vbCrLf & pNoteText
                End If
                If Trim(.CellNote) <> vbNullString Then
                    .BackColor = &HC0C0FF   'VB 'vbRed
                Else
                    .BackColor = lNoErrColor
                End If
                Exit For
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandle:
    SaveErrorLog Me.Name, "SetCellNote", Err.Number, Err.Description
End Sub

'****************************
'Description: StartReceiveForm initialize intput data screen.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'Output:
'Return:
'****************************
Private Sub StartReceiveForm()
    
On Error GoTo ErrHandle
    fpSpread1.Visible = False
    
    lblLabelVersion.Visible = False
    lblVersion.Visible = False
    lblWarning.Visible = False
    
    strTaxReportInfo = ""
    TAX_Utilities_Srv_New.xmlDataReDim (0)
    cmdViewNow.Enabled = False
    cmdClear.Enabled = False
    cmdSave.Enabled = False
    Set objTaxBusiness = Nothing
    frmSystem.MousePointer = vbDefault
    Me.MousePointer = vbDefault
        
    If blnReceiveByBarcode Then
        lblFile.Visible = False
        lblBarcode.Visible = True
        lblFilePath.caption = "0/0"
        lblLoading.Visible = True
        lblConnecting.Visible = False
        lblExit.Visible = False
        ProgressBar1.Value = 0
        ReDim arrStrElements(0)
        If Not blnOnLoadEvent Then cmdExit.SetFocus
        blnOnLoadEvent = False
    Else
        If UBound(arrStrElements) > 0 Then
            lblFile.Visible = True
            lblBarcode.Visible = False
            lblFilePath.caption = arrStrElements(UBound(arrStrElements))
            lblLoading.Visible = False
            lblConnecting.Visible = True
            lblExit.Visible = False
            
            LoadFormByFileName
        ElseIf Not blnOnLoadEvent Then
            Unload Me
            frmTreeviewMenu.Show
        Else
            lblLoading.Visible = False
            lblConnecting.Visible = False
            lblExit.Visible = True
        End If
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "StartReceiveForm", Err.Number, Err.Description
End Sub

'****************************
'Description: TrimString function cut not valid characters
'             at the begin of string.
'Author:ThanhDX
'Date:27/11/2005
'Input: strValue: data string.
'OutPut:
'Return: Data string cut (if it exist not valid characters)
'****************************
Private Function TrimString(ByVal strValue As String) As String
    Dim lCtrl As Long, strNumber As String
    On Error GoTo ErrHandle
    strNumber = "0123456789"
    If UCase(Left(strValue, 2)) = "AA" Then
        verToKhai = 0
    ElseIf UCase(Left(strValue, 2)) = "TT" Then
        verToKhai = 1
    ElseIf UCase(Left(strValue, 2)) = "BS" Then
        verToKhai = 2
    Else
        verToKhai = 0
    End If
    For lCtrl = 1 To Len(strValue)
        If InStr(1, strNumber, Mid$(strValue, lCtrl, 1)) <> 0 Then
            TrimString = Mid$(strValue, lCtrl)
            Exit Function
        End If
    Next lCtrl
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "TrimString", Err.Number, Err.Description
End Function

'****************************
'Description: ShowFormReceiveFromBarcode procedure initialize
'             form which is waiting barcode reader.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub ShowFormReceiveFromBarcode()
On Error GoTo ErrHandle
    StartBarcodeReader
    StartReceiveForm
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "ShowFormReceiveFromBarcode", Err.Number, Err.Description
End Sub

'****************************
'Description: ShowFormReceiveFromFile procedure initialize
'             form which get content form file.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub ShowFormReceiveFromFile()
    ProgressBar1.max = UBound(arrStrElements)
    StartReceiveForm
End Sub

'****************************
'Description: LoadFormByFileName procedure load data from file
'             then call load form.
'Author:ThanhDX
'Date:24/11/2005
'Input:
'OutPut:
'Return:
'****************************
Private Sub LoadFormByFileName()
    Dim intUbound As Integer
    Dim strData As String, strFileName As String
    
On Error GoTo ErrHandle
    intUbound = UBound(arrStrElements)
    If intUbound = 0 Then
        'Unload Me
        Exit Sub
    End If
    
    strFileName = arrStrElements(intUbound)
    ReDim Preserve arrStrElements(intUbound - 1)
    strData = GetDataFormFile(strFileName)
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If Not LoadForm(strData) Then
        'If UBound(arrStrElements) > 0 Then
            StartReceiveForm
    End If
    
    Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "LoadFormByFileName", Err.Number, Err.Description
End Sub

Public Sub SetReceiveByBarcode(ByVal blnValue As Boolean)
    blnReceiveByBarcode = blnValue
End Sub

Public Sub SetArrayElements(arrStrValue() As String)
    arrStrElements = arrStrValue
End Sub

'****************************
'Description: GetDataFromFile function get content in file by name.
'Author:ThanhDX
'Date:25/11/2005
'Input:
'   strFileName: name of file contain data
'OutPut:
'Return: Data string contained in the file
'****************************
Private Function GetDataFormFile(ByVal strFileName As String) As String
    Dim fso As New FileSystemObject
    Dim tstFile As TextStream
    
On Error GoTo ErrHandle
    Set tstFile = fso.OpenTextFile(strFileName, ForReading)
    While Not tstFile.AtEndOfStream
        GetDataFormFile = GetDataFormFile & tstFile.ReadLine
    Wend
    GetDataFormFile = TAX_Utilities_Srv_New.Convert(GetDataFormFile, TCVN, UNICODE)
    tstFile.Close
    Set fso = Nothing
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetDataFromFile", Err.Number, Err.Description
End Function

''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Private Function UpdateCell(ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String) As Boolean
    On Error GoTo ErrHandle
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fpSpread1, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_Srv_New.Data(fpSpread1.ActiveSheet - 1).nodeFromID(GetCellID(fpSpread1, pCol, pRow))
    
    If GetAttribute(xmlNodeCell, "Value") <> pValue Then
        SetAttribute xmlNodeCell, "Value", pValue
        UpdateCell = True
    End If
    
    Set xmlNodeCell = Nothing
    
    Exit Function
    
ErrHandle:
    SaveErrorLog Me.Name, "UpdateCell", Err.Number, Err.Description
End Function

'****************************
'Description: MessageBox function stop barcode reader then call message box
'             and start barcode reader (If received data method  is bacode reader)
'Author:ThanhDX
'Date:24/11/2005
'Input:
'   strMsgId: Message Id
'   intMsgStyle: Style of message
'   intMsgIcon: Type of icon message
'Output:
'Return:Action user
'****************************
Private Function MessageBox(strMsgId As String, intMsgStyle As MsgBoxStyle, intMsgIcon As MsgBoxIcon, Optional msType As Byte) As MsgBoxResult
    Dim intReturn As Integer
    
On Error GoTo ErrHandle
    If blnReceiveByBarcode Then StopBarcodeReader
    
    MessageBox = DisplayMessage(strMsgId, intMsgStyle, intMsgIcon, , msType)
    
    If blnReceiveByBarcode Then StartBarcodeReader
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "MessageBox", Err.Number, Err.Description
End Function

''' LoadInitFiles description
''' Set max len for string type cell
''' Set min value for numeric type cell
''' Set max value for numeric type cell
''' Call after load template
''' No parameter
Private Sub LoadInitFiles()
    On Error GoTo ErrorHandle
    Dim i As Long, lCol As Long, lRow As Long
    Dim xmlNodeListIni As MSXML.IXMLDOMNodeList
    Dim xmlNodeIni As MSXML.IXMLDOMNode
    
    For i = 0 To fpSpread1.SheetCount - 2
        ReDim Preserve xmlDocumentInit(i)
        Set xmlDocumentInit(i) = New MSXML.DOMDocument
        xmlDocumentInit(i).Load GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(i), "Ini"))
        Set xmlNodeListIni = xmlDocumentInit(i).getElementsByTagName("Cell")
        For Each xmlNodeIni In xmlNodeListIni
            fpSpread1.Sheet = i + 1
            ParserCellID fpSpread1, GetAttribute(xmlNodeIni, "CellID"), lCol, lRow
            fpSpread1.Col = lCol
            fpSpread1.Row = lRow
            If Val(GetAttribute(xmlNodeIni, "MaxLen")) <> 0 Then
                fpSpread1.TypeMaxEditLen = Val(GetAttribute(xmlNodeIni, "MaxLen"))
            End If
            If fpSpread1.CellType = CellTypeNumber Then
                fpSpread1.TypeNumberMin = Val(GetAttribute(xmlNodeIni, "MinValue"))
                fpSpread1.TypeNumberMax = Val(GetAttribute(xmlNodeIni, "MaxValue"))
            End If
            fpSpread1.CellTag = GetAttribute(xmlNodeIni, "HelpContexID") & fpSpread1.CellTag
        Next
    Next
    
    Set xmlNodeIni = Nothing
    Set xmlNodeListIni = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "LoadInitFiles", Err.Number, Err.Description
End Sub
'
''****************************************************
''Description: GetDynRowCount function get count of interface rows in
''             one Cells node.
''Author: ThanhDX
''Date:14/12/2006
''Input:
''       pGrid: fpSpread
''       xmlNodeCells: Cells node in dynamic section
''       lReportRows: Count of report rows in Cells node
''       lMinRow: Min row in Cells node
''       lMaxRow: Max row in Cells node
''****************************************************
'Public Function GetDynRowCount(pGrid As fpSpread, xmlNodeCells As MSXML.IXMLDOMNode, Optional ByRef lMinRow As Long, Optional lMaxRow As Long)
'    Dim xmlNodeCell As MSXML.IXMLDOMNode
'    Dim lRow As Long, lCol As Long
'
'    lMinRow = 100000
'    lMaxRow = 0
'
'    If Not xmlNodeCells Is Nothing Then
'        For Each xmlNodeCell In xmlNodeCells.childNodes
'            'Get CellID
'            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
'
'            'Get max row
'            If lRow > lMaxRow Then
'                lMaxRow = lRow
'            End If
'
'            'Get min row
'            If lRow < lMinRow Then
'                lMinRow = lRow
'            End If
'        Next
'
'        GetDynRowCount = lMaxRow - lMinRow + 1
'    End If
'End Function

'Private Function GetThongTinTep(ByVal strID As String, arrStrHeaderData() As String) As Boolean
'    Dim lngIndex As Long
'    Dim rsResult As ADODB.Recordset
'    Dim strSQL As String, strMaTkhaiQLT As String
'    Dim strPrefixMaTep As String, strMatep As String
'    Dim strSTT As String
'
'    On Error GoTo ErrHandle
'
'    lngIndex = UBound(arrStrHeaderData)
'
'    On Error GoTo ConnectErrHandle
'    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'    'Lay ma to khai theo QLT
'    strSQL = "Select hso.loai_hoso " & _
'            "From rcv_map_tkhai tkhai," & _
'            "qlt_map_hoso_tkhai hso " & _
'            "Where (tkhai.nhom_hso = hso.nhom) " & _
'            "And (tkhai.ma_tkhai_qlt = hso.loai_tkhai) " & _
'            "And (tkhai.ma_tkhai = '" & strID & "')"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strMaTkhaiQLT = rsResult.Fields(0).Value
'
'    'La^'y chuo^~i tie^`n to^' cu?a ma~ te^.p
'    strSQL = "Select To_Char(Sysdate,'RRMM')||'" & strMaTkhaiQLT & _
'            "' From Dual"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strPrefixMaTep = rsResult.Fields(0).Value
'
'    'Lay so thu tu lon nhat cua tep (hau to)
'    strSQL = "Select nvl(max(To_Number(Substr(So_Hieu_tep,8,3))),1) " & _
'            "From rcv_tkhai_hdr " & _
'            "Where So_Hieu_Tep Like '" & strPrefixMaTep & "' || '%'"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strMatep = strPrefixMaTep & "-" & rsResult.Fields(0).Value
'
'    'Lay so to khai lon nhat trong tep tim dc
'    strSQL = "Select nvl(max(so_tt_tk),0) + 1 " & _
'            "From rcv_tkhai_hdr " & _
'            "Where So_Hieu_tep = '" & strMatep & "'"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    strSTT = rsResult.Fields(0).Value
'
'    If CInt(strSTT) > 50 Or CInt(strSTT) = 1 Then
'        'Dong tep, sinh tep moi
'        'Lay ma tep tu DB QHS
'        Dim intSuffixMaTep As Integer
'        strSQL = "Select nvl(max(To_Number(Substr(So_Hieu,8,3))),0)+1 " & _
'                "From Qhs_Tep_Hoso " & _
'                "Where So_Hieu Like '" & strPrefixMaTep & "' || '%'"
'
'        Set rsResult = clsDAO.Execute(strSQL)
'        intSuffixMaTep = CInt(rsResult.Fields(0).Value)
'
'        'Kie^?m tra te^.p du?a va`o da~ co' du?~ lie^.u hay chu?a.
'        strSQL = "Select so_hoso " & _
'                "From Qhs_Tep_Hoso " & _
'                "Where (so_hieu = '" & strPrefixMaTep & "-" & (intSuffixMaTep - 1) & "')"
'
'        Set rsResult = clsDAO.Execute(strSQL)
'
'        If Not rsResult Is Nothing Then
'            If rsResult.Fields(0).Value <> "0" Or IsNull(rsResult.Fields(0)) Then
'CreateFile:
'                strMatep = strPrefixMaTep & "-" & intSuffixMaTep
'                strSTT = "1"
'
'                'Insert tep moi vao QHS
'                strSQL = "Insert Into Qhs_Tep_Hoso (So_Hieu, Dhs_Ma, " & _
'                        "Kykk_Tu_Ngay, Kykk_Den_Ngay, Ngay_Tao, So_Hoso)" & _
'                        "Values ('" & strMatep & "'," & _
'                        "'" & strMaTkhaiQLT & "'," & _
'                        "To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')," & _
'                        "To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')," & _
'                        "Sysdate,0)"
'                clsDAO.Execute (strSQL)
'                strSQL = "commit"
'                clsDAO.Execute (strSQL)
'            Else
'                strMatep = strPrefixMaTep & "-" & (intSuffixMaTep - 1)
'                strSTT = "1"
'            End If
'        Else
'            GoTo CreateFile
'        End If
'    End If
'
'    On Error GoTo ErrHandle
'
'    'Ghep ma so tep vao chuoi
'    ReDim Preserve arrStrHeaderData(lngIndex + 1)
'    arrStrHeaderData(lngIndex + 1) = strMatep
'
'    'Ghep so thu tu to khai vao chuoi
'    ReDim Preserve arrStrHeaderData(lngIndex + 2)
'    arrStrHeaderData(lngIndex + 2) = "" 'strSTT
'
'    Set rsResult = Nothing
'    GetThongTinTep = True
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "GetThongTinTep", Err.Number, Err.Description
'    Exit Function
'ConnectErrHandle:
'    SaveErrorLog Me.Name, "GetThongTinTep", Err.Number, Err.Description
'End Function

Private Function getSoTTTK(ByVal strID As String, arrStrHeaderData() As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String
    Dim strMatep As String
    Dim strSTT As Integer
    
    On Error GoTo ErrHandle
    
    lngIndex = UBound(arrStrHeaderData)
    
    On Error GoTo ConnectErrHandle
    'connect to database QLT_TNK
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'    'Lay so TT to khai trong RCV
'    If strID = "02_TNDN11" And isTKLanPS = True Then
'        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
'                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'                "And tkhai.loai_tkhai = '" & strID & "' " & _
'                " And tkhai.ngay_ps = to_date('" & ngayPS & "','dd/mm/yyyy')"
'    ElseIf (strID = "01_NTNN" Or strID = "01_TTDB11" Or strID = "03_NTNN11") And isTKLanPS = True Then
'        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
'                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'                "And tkhai.loai_tkhai = '" & strID & "' " & _
'                " And tkhai.ngay_ps = to_date('" & ngayPS & "','dd/mm/yyyy')"
'    ElseIf (strID = "08_TNCN11" Or strID = "08A_TNCN11") And isTKThang = True Then
'        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
'                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'                "And tkhai.loai_tkhai = '" & strID & "' " & _
'                "And tkhai.kykk_tu_ngay = To_Date('" & "01/" & TuNgay & "','DD/MM/RRRR')" & _
'                "And tkhai.kykk_den_ngay = To_Date('" & "01/" & DenNgay & "','DD/MM/RRRR')"
'    Else
'        strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
'                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'                "And tkhai.loai_tkhai = '" & strID & "' " & _
'                "And tkhai.kykk_tu_ngay = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
'                "And tkhai.kykk_den_ngay = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
'    End If
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    If rsResult Is Nothing Or IsNull(rsResult.Fields(0)) Then
'        strSTT = 0
'        isTKTonTai = False
'        ' Doi voi cac to khai 01_NTNN, 03_NTNN, 01_TTDB, 02_TNDN
'        If (strID = "01_NTNN" Or strID = "01_TTDB11" Or strID = "03_NTNN11" Or strID = "02_TNDN11") And isTKLanPS = True Then
'            isToKhaiPsDaNhanTN = False
'        End If
'
'    Else
'        strSTT = rsResult.Fields(0).Value + 1
'        isTKTonTai = True
'        ' Doi voi cac to khai 01_NTNN, 03_NTNN, 01_TTDB, 02_TNDN trong 1 ngay chi nhan 1 to khai
'        If (strID = "01_NTNN" Or strID = "01_TTDB11" Or strID = "03_NTNN11" Or strID = "02_TNDN11") And isTKLanPS = True Then
'            isToKhaiPsDaNhanTN = True
'        End If
'    End If
    
    ' Kiem tra to khai chinh thuc
'    If strSTT = 0 Then
'        isToKhaiCT = False
'    Else
'        isToKhaiCT = True
'    End If
    isTKTonTai = False
    isToKhaiCT = True
    isToKhaiPsDaNhanTN = False
    strSTT = 0
    'Ghep ma so tep vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 1)
    arrStrHeaderData(lngIndex + 1) = strSTT
    
    'Ghep so thu tu to khai vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 2)
    arrStrHeaderData(lngIndex + 2) = strSTT
    
    Set rsResult = Nothing
    getSoTTTK = True
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK", Err.Number, Err.Description
End Function


' Kiem tra to khai theo DA30
Private Function isDA30(ByVal strID As String, arrStrHeaderData() As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String
    
    isDA30 = False
    Exit Function
    
'    On Error GoTo ErrHandle
'    On Error GoTo ConnectErrHandle
'    'connect to database QLT_TNK
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'    strSQL = "select 1 from qlt_tkhai_hdr tkhai " & _
'            "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'            "And tkhai.DTK_MA_LOAI_TKHAI = '" & changeMaToKhaiQLT(strID) & "' " & _
'            "And tkhai.kykk_tu_ngay = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
'            "And tkhai.kykk_den_ngay = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
'            "And tkhai.YN_DA30 is null "
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    If rsResult Is Nothing Then
'        isDA30 = False
'    Else
'        isDA30 = True
'    End If
'
'    Set rsResult = Nothing
'    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "isDA30", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "isDA30", Err.Number, Err.Description




    isDA30 = False
End Function
' end

Private Function GetHeaderData(ByVal rsTaxInfor As ADODB.Recordset, arrStrHeaderData() As String) As Boolean
    Dim arrStrData() As String
    Dim lCtrl As Long
    Dim clsConvert  As New clsUnicodeConvert
    On Error GoTo ErrHandle
    
    If rsTaxInfor Is Nothing Then
        Exit Function
    End If
    
    If rsTaxInfor.Fields.Count = 0 Then
        Exit Function
    End If
    
    For lCtrl = 0 To rsTaxInfor.Fields.Count - 2
        ReDim Preserve arrStrData(lCtrl)
        If Not IsNull(rsTaxInfor.Fields(lCtrl + 1).Value) Then
            'arrStrData(lCtrl) = clsConvert.Convert(rsTaxInfor.Fields(lCtrl + 1).Value, UNICODE, TCVN)
            arrStrData(lCtrl) = rsTaxInfor.Fields(lCtrl + 1).Value
            
        End If
    Next lCtrl
           
    'Loai bo gia tri Ngay bat dau nam TC va Ngay bat dau KD
    ReDim Preserve arrStrData(UBound(arrStrData) - 1)
    
    arrStrHeaderData = arrStrData
    GetHeaderData = True
    
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "GetHeaderData", Err.Number, Err.Description
End Function

' Ham lay ve so tt quet An chi
Private Function getSoTTTK_AC(ByVal strID As String, arrStrHeaderData() As String, strData As String) As Boolean
    Dim lngIndex As Long
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String
    Dim strMatep As String
    Dim strSTT As Integer
    
    Dim arrDeltail() As String
    
    On Error GoTo ErrHandle
    
    lngIndex = UBound(arrStrHeaderData)
    
    On Error GoTo ConnectErrHandle
    'connect to database QLT_TNK
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'    ' Tach ma so thue 13 thanh ma so thue 14
'    If Len(Trim(arrStrHeaderData(0))) = 13 Then
'        arrStrHeaderData(0) = Left(Trim(arrStrHeaderData(0)), 10) & "-" & Right(Trim(arrStrHeaderData(0)), 3)
'    End If
'
'    'Lay so TT to khai trong RCV
'    If strID = "01_TBAC" Then
'        arrDeltail = Split(strData, "~")
'        If Len(Trim(arrDeltail(UBound(arrDeltail) - 3))) = 13 Then
'            arrDeltail(UBound(arrDeltail) - 3) = Left(arrDeltail(UBound(arrDeltail) - 3), 10) & "-" & Right(arrDeltail(UBound(arrDeltail) - 3), 3)
'        End If
'
'        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
'        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'        "And tkhai.LOAI_BC = '" & strID & "' " & _
'        " And tkhai.NGAY_BC=to_date('" & arrDeltail(UBound(arrDeltail) - 1) & "','dd/mm/rrrr')" & _
'        " And tkhai.TIN_DV_CQ='" & Trim(arrDeltail(UBound(arrDeltail) - 3)) & "'"
'    ElseIf strID = "03_TBAC" Then
'        arrDeltail = Split(strData, "~")
'        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
'        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'        "And tkhai.LOAI_BC = '" & strID & "' " & _
'        " And tkhai.NGAY_BC=to_date('" & Left$(arrDeltail(UBound(arrDeltail)), 10) & "','dd/mm/rrrr')"
'    ElseIf strID = "BC21_AC" Then
'        arrDeltail = Split(strData, "~")
'        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
'        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'        "And tkhai.LOAI_BC = '" & strID & "' " & _
'        " And tkhai.NGAY_BC=to_date('" & Left$(arrDeltail(UBound(arrDeltail)), 10) & "','dd/mm/rrrr')"
'    ElseIf strID = "01_AC" Then
'        arrDeltail = Split(strData, "~")
'        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
'        "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'        "And tkhai.LOAI_BC = '" & strID & "' " & _
'        "And tkhai.KYBC_TU_NGAY = to_date('" & arrDeltail(1) & "','dd/mm/rrrr')" & _
'        "And tkhai.KYBC_DEN_NGAY = to_date('" & Left$(arrDeltail(2), 10) & "','dd/mm/rrrr')"
'    Else
'        strSQL = "select max(so_tt_tk) from rcv_bcao_hdr_ac tkhai " & _
'                "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'                "And tkhai.LOAI_BC = '" & strID & "' " & _
'                "And tkhai.KYBC_TU_NGAY = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
'                "And tkhai.KYBC_DEN_NGAY = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
'    End If
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    If rsResult Is Nothing Or IsNull(rsResult.Fields(0)) Then
'        strSTT = 0
'        isTonTaiAC = False
'    Else
'        strSTT = rsResult.Fields(0).Value + 1
'        isTonTaiAC = True
'    End If
    isTonTaiAC = False
    strSTT = 0
    'Ghep ma so tep vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 1)
    arrStrHeaderData(lngIndex + 1) = strSTT
    
    'Ghep so thu tu to khai vao chuoi
    ReDim Preserve arrStrHeaderData(lngIndex + 2)
    arrStrHeaderData(lngIndex + 2) = strSTT
    
    Set rsResult = Nothing
    getSoTTTK_AC = True
    Exit Function
ErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK_AC", Err.Number, Err.Description
    Exit Function
ConnectErrHandle:
    SaveErrorLog Me.Name, "getSoTTTK_AC", Err.Number, Err.Description
End Function


' Kiem tra xem co to khai chinh thuc chua
'Private Function isToKhaiCT(ByVal strID As String, arrStrHeaderData() As String) As Boolean
'    Dim lngIndex As Long
'    Dim rsResult As ADODB.Recordset
'    Dim strSQL As String
'    Dim strMatep As String
'    Dim strSTT As Integer
'
'    On Error GoTo ErrHandle
'
'    lngIndex = UBound(arrStrHeaderData)
'
'    On Error GoTo ConnectErrHandle
'    'connect to database QLT_TNK
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'    'Lay so TT to khai trong RCV
'
'    strSQL = "select max(so_tt_tk) from rcv_tkhai_hdr tkhai " & _
'            "Where tkhai.tin = '" & arrStrHeaderData(0) & "'" & _
'            "And tkhai.loai_tkhai = '" & strID & "' " & _
'            "And tkhai.kkbs='1' " & _
'            "And tkhai.kykk_tu_ngay = To_Date('" & format$(dNgayDauKy, "DD/MM/YYYY") & "','DD/MM/RRRR')" & _
'            "And tkhai.kykk_den_ngay = To_Date('" & format$(dNgayCuoiKy, "DD/MM/YYYY") & "','DD/MM/RRRR')"
'
'    Set rsResult = clsDAO.Execute(strSQL)
'    If rsResult Is Nothing Or IsNull(rsResult.Fields(0)) Then
'        isToKhaiCT = False
'    Else
'        isToKhaiCT = True
'    End If
'
'    Set rsResult = Nothing
'    Exit Function
'ErrHandle:
'    SaveErrorLog Me.Name, "isToKhaiCT", Err.Number, Err.Description
'    Exit Function
'ConnectErrHandle:
'    SaveErrorLog Me.Name, "isToKhaiCT", Err.Number, Err.Description
'End Function

' Kiem tra mst dai ly co phai cua NNT hay khong?
' Lay thong tin DL thue 05072011
Private Function isMaDLT(ByVal strTaxIDString As String, ByVal strTaxIDDLString As String) As Boolean
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo ErrHandle

    'connect to database QLT
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'
'
'    ' Get SQL statement from DOM
'    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlMSTDL")
'    strSQL = Replace(strSQL, "strMST", "'" & strTaxIDString & "'")
'    strSQL = Replace(strSQL, "ma_dai_ly", "'" & strTaxIDDLString & "'")
'
'    Set rsReturn = clsDAO.Execute(strSQL)
'
'    If rsReturn Is Nothing Or rsReturn.Fields.Count = 0 Then
'        If Trim(strTaxIDDLString) = "" Or strTaxIDDLString = vbNullString Then
'            isMaDLT = True
'        Else
'            isMaDLT = False
'        End If
'    Else
'        isMaDLT = True
'    End If
    
'    Set rsReturn = Nothing

    'Lay ham service check ma dai ly thue
    isMaDLT = True
    Exit Function
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "isMaDLT", Err.Number, Err.Description
    If Err.Number = -2147467259 Then _
        MessageBox "0063", msOKOnly, miCriticalError
End Function

Private Function LoaiToKhai(ByVal strData As String) As Boolean
    Dim LoaiTk As String
    Dim tmp    As String
    
    On Error GoTo ErrHandle
    
    '    tmp = Mid(strData, 1, InStr(1, strData, "</S01>", vbTextCompare) + 5)
    '    tmp = Left$(tmp, Len(tmp) - 10)
    'LoaiTk = Right$(tmp, 1)
    LoaiTk = Left$(strData, Len(strData) - 10)
    LoaiTk = Right$(LoaiTk, 1)

    If LoaiTk = "1" Then
        LoaiToKhai = True
    Else
        LoaiToKhai = False
    End If
    
    Exit Function
ErrHandle:
    'Connect DB fail
    SaveErrorLog Me.Name, "LoaiToKhai", Err.Number, Err.Description

    If Err.Number = -2147467259 Then MessageBox "0063", msOKOnly, miCriticalError
End Function

Public Function AppendXMLStandard(ByVal xmlDoc As MSXML.DOMDocument, _
                                  ByVal sKyLapBo As String, _
                                  ByVal sNgayNopTK As String) As MSXML.DOMDocument
    Dim XmlDocStandard As New MSXML.DOMDocument
    XmlDocStandard.Load GetAbsolutePath("..\InterfaceTemplates\xml\TempStandard.xml")
    
    'Doc file cau hinh lay thong tin header
    Dim xmlConfig As MSXML.DOMDocument
    Set xmlConfig = LoadConfig()
    XmlDocStandard.getElementsByTagName("VERSION")(0).Text = xmlConfig.getElementsByTagName("VERSION")(0).Text
    XmlDocStandard.getElementsByTagName("SENDER_CODE")(0).Text = xmlConfig.getElementsByTagName("SENDER_CODE")(0).Text
    XmlDocStandard.getElementsByTagName("SENDER_NAME")(0).Text = xmlConfig.getElementsByTagName("SENDER_NAME")(0).Text
    XmlDocStandard.getElementsByTagName("RECEIVER_CODE")(0).Text = xmlConfig.getElementsByTagName("RECEIVER_CODE")(0).Text
    XmlDocStandard.getElementsByTagName("RECEIVER_NAME")(0).Text = xmlConfig.getElementsByTagName("RECEIVER_NAME")(0).Text
    XmlDocStandard.getElementsByTagName("TRAN_CODE")(0).Text = xmlConfig.getElementsByTagName("TRAN_CODE")(0).Text
    XmlDocStandard.getElementsByTagName("ORIGINAL_CODE")(0).Text = xmlConfig.getElementsByTagName("ORIGINAL_CODE")(0).Text
    XmlDocStandard.getElementsByTagName("ORIGINAL_NAME")(0).Text = xmlConfig.getElementsByTagName("ORIGINAL_NAME")(0).Text
    
    XmlDocStandard.getElementsByTagName("MSG_ID")(0).Text = xmlConfig.getElementsByTagName("SENDER_CODE")(0).Text & GenerateCodeByNow() 'GetGUID()
    XmlDocStandard.getElementsByTagName("SEND_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")
    XmlDocStandard.getElementsByTagName("ORIGINAL_DATE")(0).Text = Format(DateTime.Now, "dd-mmm-yyyy HH:mm:ss")
    
    XmlDocStandard.getElementsByTagName("SPARE1")(0).Text = strUserName
    XmlDocStandard.getElementsByTagName("SPARE2")(0).Text = strMaNNT
    
    ' Set value tag <add_info>
    XmlDocStandard.getElementsByTagName("ngay_nop_tk")(0).Text = sNgayNopTK
    XmlDocStandard.getElementsByTagName("ky_lap_bo")(0).Text = sKyLapBo
    XmlDocStandard.getElementsByTagName("nguon_goc_tk")(0).Text = xmlConfig.getElementsByTagName("SENDER_CODE")(0).Text
    XmlDocStandard.getElementsByTagName("nguoi_nhan_tk")(0).Text = strUserID '& "." & xmlConfig.getElementsByTagName("CODE_OFFICE")(0).Text
    XmlDocStandard.getElementsByTagName("ngay_nhan_tk")(0).Text = Format(DateTime.Now, "dd/MM/yyyy")
    XmlDocStandard.getElementsByTagName("id_tkhai")(0).Text = xmlConfig.getElementsByTagName("SENDER_CODE")(0).Text & GenerateCodeByNow()
    
    XmlDocStandard.getElementsByTagName("noi_gui")(0).Text = ""
    XmlDocStandard.getElementsByTagName("noi_nhan")(0).Text = ""
    
    'Bo sung tag <QHS> cho BCTC va AC
    'ID BCTC: 69(15_BCTC); 19(48_BCTC); 20(16_BCTC); 21(99_BCTC); 22(95_BCTC);
    'ID AC:   64(01_TBAC); 65(01_AC); 66(BC21_AC); 67(03_TBAC); 68(BC26_AC); 91(04_TBAC);
    Dim strID_BCTC, strID_QLAC As String
    strID_BCTC = xmlConfig.getElementsByTagName("BCTC")(0).Text
    strID_QLAC = xmlConfig.getElementsByTagName("QLAC")(0).Text
    
    Dim tempQHSxml As New MSXML.DOMDocument
    Dim nodeVal      As MSXML.IXMLDOMNode
    Dim nodeValIndex As Integer
    
    If (InStr(strID_BCTC, GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) > 0) Then
        '       Dim newNode As MSXML.IXMLDOMNode
        '       Set newNode = XmlDocStandard.createElement("QHS")
        '       XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO").appendChild newNode
        '       XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO/QHS").appendChild XmlDocStandard.createElement("PL_KQHDSXKD01")
        'Load template QHS
        
        tempQHSxml.Load GetAbsolutePath("..\InterfaceTemplates\xml\QHS.xml")
        XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO").appendChild tempQHSxml.lastChild.firstChild
        
        For nodeValIndex = 1 To TAX_Utilities_Srv_New.NodeValidity.childNodes.length
            Set nodeVal = TAX_Utilities_Srv_New.NodeValidity.childNodes(nodeValIndex)

            If (GetAttribute(nodeVal, "Active") = "1" And (GetAttribute(nodeVal, "ID") = "01-10" Or GetAttribute(nodeVal, "ID") = "01-1")) Then
                'kqhd01 = True
                XmlDocStandard.getElementsByTagName("PL_KQHDSXKD01")(0).Text = "X"
            End If

            If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01-11") Then
                'kqhd02 = True
                XmlDocStandard.getElementsByTagName("PL_KQHDSXKD02")(0).Text = "X"
            End If

            If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01-12") Then
                'kqhd03 = True
                XmlDocStandard.getElementsByTagName("PL_KQHDSXKD03")(0).Text = "X"
            End If

            If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01-2") Then
                'lctttt = True
                XmlDocStandard.getElementsByTagName("PL_LCTTTT")(0).Text = "X"
            End If

            If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01-3") Then
                'lcttgt = True
                XmlDocStandard.getElementsByTagName("PL_LCTTGT")(0).Text = "X"
            End If

        Next
    End If

    If (InStr(strID_QLAC, GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) > 0) Then
        '       Dim newNode As MSXML.IXMLDOMNode
        '       Set newNode = XmlDocStandard.createElement("QHS")
        '       XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO").appendChild newNode
        '       XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO/QHS").appendChild XmlDocStandard.createElement("PL_KQHDSXKD01")
        'Load template QHS
        tempQHSxml.Load GetAbsolutePath("..\InterfaceTemplates\xml\QHS.xml")
        XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO").appendChild tempQHSxml.lastChild.lastChild

        If (GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "68") Then
            For nodeValIndex = 1 To TAX_Utilities_Srv_New.NodeValidity.childNodes.length
                Set nodeVal = TAX_Utilities_Srv_New.NodeValidity.childNodes(nodeValIndex)
                If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01-1") Then
                    XmlDocStandard.getElementsByTagName("PL_BK_01AC_01")(0).Text = "X"
                End If
                If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01-2") Then
                    XmlDocStandard.getElementsByTagName("PL_BK_01AC_02")(0).Text = "X"
                End If
            Next
        End If
    End If
    
    'Ket thuc bo sung <QHS>
    
    'End <add_info>

    If (Not xmlDoc Is Nothing) Then
        'XmlDocStandard.getElementsByTagName("ROW")(0).appendChild xmlDoc.getElementsByTagName("HSoKhaiThue")(0) 'xmlDoc.childNodes(0)
        XmlDocStandard.getElementsByTagName("RETURN")(0).appendChild xmlDoc.lastChild
    End If

    Set AppendXMLStandard = XmlDocStandard
End Function
