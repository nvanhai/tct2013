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
                            xmlCellTKNode.Text = .value
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
                                    xmlCellTKNode.Text = .value
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
                                
                                    dNode.Text = .value
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
    xmlTK.getElementsByTagName("ngayLapTKhai")(0).Text = ConvertDate(strNgayHeThongSrv, False, "-") 'Format(Date, "yyyy-MM-dd HH:mm:ss")
    If (Not xmlResultNNT Is Nothing) Then
        If ((Not (xmlResultNNT Is Nothing)) And xmlResultNNT.hasChildNodes And (InStr(xmlResultNNT.xml, "fault_code") <= 0)) Then
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
    End If
    xmlTK.getElementsByTagName("mst")(0).Text = strMaNNT
    If (Not xmlResultDLT Is Nothing) Then
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
        strLoaiToKhai = GetAttribute(GetMessageCellById("0132"), "Msg")
    Else
        xmlTK.getElementsByTagName("soLan")(0).Text = ""
        xmlTK.getElementsByTagName("loaiTKhai")(0).Text = GetAttribute(GetMessageCellById("0131"), "Msg")
        strLoaiToKhai = GetAttribute(GetMessageCellById("0131"), "Msg")
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
    ElseIf ID_TK = "01" Or ID_TK = "02" Or ID_TK = "04" Or ID_TK = "71" Or ID_TK = "36" Or ID_TK = "68" Or ID_TK = "25" Then

        If LoaiKyKK = True Then
            KYKKHAI = Right$(TAX_Utilities_Srv_New.ThreeMonths, 1) & "/" & TAX_Utilities_Srv_New.Year
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
            If ID_TK = "65" Then
                KYKKHAI = Right$(TAX_Utilities_Srv_New.ThreeMonths, 1) & "/" & TAX_Utilities_Srv_New.Year
                strKieuKy = "K"
            Else
                KYKKHAI = Right$(TAX_Utilities_Srv_New.ThreeMonths, 1) & "/" & TAX_Utilities_Srv_New.Year
                strKieuKy = "Q"
            End If
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

    strFileName = MaTK & ".xml" 'getFileName
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

                    'Lay ID to khai
                    If (xmlCellNode.nodeName = "maTKhai") Then
                        strMaToKhai = cellid
                    End If
                    'Ket thuc lay ID to khai
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
                            xmlCellTKNode.Text = .value
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
                    'MaTK = "KHBS"
                    MaTK = Right$(MaTK, Len(MaTK) - 5)
                    
                    If MaTK = "01_GTGT" Then
                        xmlPL.Load GetAbsolutePath("..\InterfaceTemplates\xml\KHBS_01_GTGT_xml.xml")

                        xmlMapPL.Load GetAbsolutePath("..\InterfaceIni\KHBS_01_GTGT_xml.xml")
                    Else

                        If MaTK = "02_GTGT" Or MaTK = "03_GTGT" Or MaTK = "04_GTGT" Or MaTK = "05_GTGT" Or MaTK = "01A_TNDN" Or MaTK = "01A_TNDN" Or MaTK = "02_TNDN" Or MaTK = "01_NTNN" Or MaTK = "03_NTNN" Or MaTK = "01_TAIN" Or MaTK = "01_TTDB" Or MaTK = "01_BVMT" Or MaTK = "01_TBVMT" Then
                            xmlPL.Load GetAbsolutePath("..\InterfaceTemplates\xml\KHBS_TT156_xml.xml")

                            xmlMapPL.Load GetAbsolutePath("..\InterfaceIni\KHBS_TT156_xml.xml")
                        Else
                            xmlPL.Load GetAbsolutePath("..\InterfaceTemplates\xml\KHBS_TT28_xml.xml")

                            xmlMapPL.Load GetAbsolutePath("..\InterfaceIni\KHBS_TT28_xml.xml")
                            
                            If xmlPL.getElementsByTagName("tenTKhaiThue").length > 0 Then
                                xmlPL.getElementsByTagName("tenTKhaiThue")(0).Text = GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Caption")
                            End If

                            If xmlPL.getElementsByTagName("mauSo").length > 0 Then
                                xmlPL.getElementsByTagName("mauSo")(0).Text = MaTK
                            End If

                            If xmlPL.getElementsByTagName("kyTinhThue").length > 0 And xmlTK.getElementsByTagName("kyKKhai").length > 0 Then
                                xmlPL.getElementsByTagName("kyTinhThue")(0).Text = xmlTK.getElementsByTagName("kyKKhai")(0).Text
                            End If
                        End If
                    End If
                Else
                    xmlPL.Load GetAbsolutePath("..\InterfaceTemplates\xml\" & MaTK & "_xml.xml")

                    xmlMapPL.Load GetAbsolutePath("..\ini\" & MaTK & "_xml.xml")
                End If

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
                                        xmlCellTKNode.Text = .value
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
    If (Not PushDataToESB(xmlDocSave.xml)) Then
        MessageBox "0137", msOKOnly, miCriticalError
    End If
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
    
    dsTK_DLT = "~1~2~3~4~5~6~11~12~46~47~48~49~15~16~50~51~36~70~71~72~73~74~75~80~81~82~77~86~87~89~42~43~17~59~41~76~90~92~93~98~99~25~"
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
            
        'If ((idToKhai = 74 Or idToKhai = 75) And isTKThang = True) Then
'        If ((idToKhai = 75) And isTKThang = True) Then
'            Dim arrNgay() As String
'            arrNgay = Split(TuNgay, "/")
'
'            If Val(arrNgay(1)) < 2011 Or (Val(arrNgay(1)) = 2011 And Val(arrNgay(0)) < 7) Then
'                MessageBox "0118", msOKOnly, miWarning
'                Exit Sub
'            End If
'        End If
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
            varTemp = .value
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
            varTemp = .value
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
                varTemp = .value
            End With
        ElseIf Val(idToKhai) = 12 Then
            With fpSpread1
                .Sheet = 1
                .Col = .ColLetterToNumber("E")
                .Row = 38
                varTemp = .value
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
            varTemp = .value
        End With
    End If
    ' Do voi to khai 03/TNDN thang
    If Val(idToKhai) = 3 Then
        ' Lay thong tin ve gia han nop thue TNDN
        With fpSpread1
            .Sheet = 1
            .Col = .ColLetterToNumber("E")
            .Row = 37
            varTemp = .value
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
   
'***************** Ra soat giai doan 1 *******************
''TB01/AC:
'str2 = "aa999642300026193   12201300400400100101/0101/01/2009<S01><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT4/001~AB/12T~91~0000010~0000100~18/12/2013~Doanh nghi÷p acb~0102030405~0456~11/10/2013~~H„a Æ¨n b∏n hµng (dµnh cho tÊ ch¯c, c∏ nh©n trong khu phi thu’ quan)~07KPTQ5/003~CD/23P~126~0000025~0000150~19/12/2013~dbh~0010011000~0567~12/10/2013~~Phi’u xu t kho gˆi b∏n hµng Æπi l˝~04HGDL6/005~MN/19T~151~0000050~0000200~20/12/2013~doanh nghi÷p~3600247325~0678~11/11/2013~</S><S>~~c¨ quan thu’ quÀn ba Æ◊nh~11/12/2013~Kh∏nh V©n</S></S01>"
'Barcode_Scaned str2
'
'BC01/AC:
str2 = "aa999652300026193   02201300200200100101/0101/01/2009<S01><S>x~01/07/2013~31/12/2013</S><S>0102030405~doanh nghi÷p ab~hµ nÈi~234~13/11/2013~H„a Æ¨n xu t kh»u~06HDXK6/001~BH/12T~0000024~0000150~127~~0010011000~doanh nghi÷p 1~bæc ninh~094~14/11/2013~Phi’u xu t kho ki™m vÀn chuy”n hµng h„a nÈi bÈ~03XKNB5/003~BN/23T~0000036~0000189~154~~3600247325~MNP~hµ giang~056~15/11/2013~H„a Æ¨n b∏n hµng~02GTTT6/005~HD/34T~0000045~0000251~207~</S><S>Hπnh~25/10/2013</S></S01>"
Barcode_Scaned str2
'
''BC21/AC:
'str2 = "aa999662300026193   03201300100100100101/0101/01/2010<S01><S>25/10/2013~15~1</S><S>H„a Æ¨n xu t kh»u~06HDXK6/001~AB/23T~0000030~0000250~221~7;12;24~05~0~H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT4/003~CD/14T~0000025~0000167~143~6;14;28;30~04~0~Phi’u xu t kho ki™m vÀn chuy”n hµng h„a nÈi bÈ~03XKNB5/007~HN/16T~0000031~0000274~244~11;23;45~02~0</S><S>ch∏y h„a Æ¨n~Hπnh~25/10/2013</S></S01>"
'Barcode_Scaned str2
'
''TB03/AC:
'str2 = "aa999672300026193   03201300200200100101/0101/01/2010<S01><S>Cty ACB~ph≠¨ng ph∏p hÒy~1~25/10/2013~15</S><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT4/003~HN/12P~0000013~0000237~225~0~Phi’u xu t kho ki™m vÀn chuy”n hµng h„a nÈi bÈ~03XKNB2/008~DB/34T~0000020~0000176~157~0</S><S>H∂i Nguy™n~Hπnh~25/10/2013</S></S01>"
'Barcode_Scaned str2
'
''BC26/ AC + PL dinh kem:
'str2 = "aa999682300026193   02201300200200100301/0101/01/2009<S01><S>~X~01/04/2013~30/06/2013</S><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/001~MN/12T~131~0000020~0000150~~~0000020~0000032~13~10~1~20~1~23~1~26~0000033~0000150~118~0~H„a Æ¨n b∏n hµng (dµnhcho tÊ ch¯c, c∏ nh©n trong khu phi thu’ quan)~07KPTQ6/003~CB/23T~95~~"
'Barcode_Scaned str2
'str2 = "aa999682300026193   022013002002002003~0000030~0000124~0000030~0000038~9~6~1~32~1~34~1~33~0000039~0000124~86~0~Phi’u xu t kho gˆi b∏n hµng Æπi l˝~04HGDL7/004~HD/26T~338~0000013~0000210~0000211~0000350~0000013~0000027~15~12~1~15~1~20~1~25~0000028~0000350~323~0</S><S>H∂i Nguy™n~Hoµng Nam~14/08/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999682300026193   022013002002003003<S01_2><S>H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT3/001~MN/12T~0000033~0000150~118~H„a Æ¨n b∏n hµng (dµnh cho tÊ ch¯c, c∏ nh©n trong khu phi thu’ quan)~07KPTQ6/003~CB/23T~0000039~0000124~86~Phi’u xu t kho gˆi b∏n hµng Æπi l˝~04HGDL7/004~HD/26T~0000028~0000350~323</S></S01_2>"
'Barcode_Scaned str2


'str2 = "aa999682300790401   04201300200200100201/0101/01/2009<S01><S>X~~01/10/2013~31/12/2013</S><S>H„a Æ¨n b∏n hµng~02GTTT2/001~AA/12T~1800005~0100000~1900004~~~~~0~0~0~~0~~0~~~~0~0~H„a Æ¨n gi∏ trﬁ gia t®ng~01GTKT2/001~AA/12T~900000~0100001~1000000~~~~~0~0~0~~0~~0~~~~0~0</S><S>~Hoµng~02/01/2014~1</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999682300790401   042013002002002002<S01_1><S>~~~~~0~</S><S>~~~~~0~</S><S></S></S01_1>"
'Barcode_Scaned str2
'
''TB04/AC:
'str2 = "aa999912300026193   12201300200200100101/0101/01/2009<S01><S>ph∏t hµnh~h„a Æ¨n~01DVPH~trÙ sÎ~Æﬁa chÿ ch›nh~02DCTS~fax~hotline~03PHONE</S><S>10/11/2013~~~c¨ quan thu’ TP bæc ninh~11/12/2013~Kh∏nh V©n</S></S01>"
'Barcode_Scaned str2

'str2 = "aa999912300790401   01201400200200100101/0101/01/2009<S01><S>asdfsdf~sdfsdf~01DVPH</S><S>11/11/2013~~~asdfdsf~02/01/2014~Hoµng</S></S01>"
'Barcode_Scaned str2


''BCTC bo QD15:
'str2 = "aa999692300026210   00201200000000100801/0123/06/2006<S01><S>~73920~0~~21000~0~V.01~1000~0~~20000~0~V.02~3000~0~~1000~0~~2000~0~~33000~0~~3000~0~~4000~0~~5000~0~~6000~0~V.03~7000~0~~8000~0~~3000~0~V.04~1000~0~~2000~0~~13920~0~~3000~0~~5000~0~V.05~3450~0~~2470~0~~95780~0~~19940~0~~6920~0~~3420~0~V.06~5600~0~V.07~1300~0~~2700~0~~41000~0~V.08~9400~0~~6000~0"
'Barcode_Scaned str2
'str2 = "aa999692300026210   002012000000002008~~3400~0~V.09~7400~0~~2900~0~~4500~0~V.10~14800~0~~6700~0~~8100~0~V.11~9400~0~V.12~12000~0~~9000~0~~3000~0~~11000~0~~2000~0~~1000~0~V.13~3000~0~~5000~0~~11840~0~V.14~5630~0~V.21~3210~0~~3000~0~~169700~0~~97680~0~~53030~0~V.15~1000~0~~2000~0~~3000~0~V.16~4000~0~~5000~0~V.17~6000~0~~7000~0~~8000~0~V.18~9000~0~~3450~0~~"
'Barcode_Scaned str2
'str2 = "aa999692300026210   0020120000000030084580~0~~44650~0~~3980~0~V.19~8430~0~~2180~0~V.20~1560~0~V.21~4000~0~~5300~0~~7500~0~~3600~0~~8100~0~~72020~0~V.22~59565~0~~4700~0~~2000~0~~1000~0~~4000~0~~2000~0~~6000~0~~9000~0~~3000~0~~5000~0~~6000~0~~7000~0~~9865~0~~12455~0~V.23~9455~0~~3000~0~~169700~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~Hoa Linh~22/10/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999692300026210   002012000000004008<S01-1><S>VI.25~30000~0~~1000~0~~29000~0~VI.27~2000~0~~27000~0~VI.26~1000~0~VI.28~4000~0~~5000~0~~6000~0~~2000~0~~16000~0~~5000~0~~3000~0~~2000~0~~18000~0~VI.30~2100~0~VI.30~3800~0~~12100~0~~2000~0~Hoa Linh~22/10/2013</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999692300026210   002012000000005008<S01-2><S>~1000~0~~2000~0~~3000~0~~4000~0~~5000~0~~6000~0~~7000~0~~28000~0~~1000~0~~2000~0~~3000~0~~4000~0~~5000~0~~7000~0~~2000~0~~240"
'Barcode_Scaned str2
'str2 = "aa999692300026210   00201200000000600800~0~~3000~0~~2000~0~~5000~0~~6000~0~~7000~0~~9000~0~~32000~0~~84000~0~~8000~0~~1000~0~VII.34~93000~0~Hoa Linh~22/10/2013</S></S01-2>"
'Barcode_Scaned str2
'str2 = "aa999692300026210   002012000000007008<S01-3><S>~12000~0~~1000~0~~2000~0~~3000~0~~4000~0~~5000~0~~27000~0~~2000~0~~7000~0~~3000~0~~2000~0~~1000~0~~8000~0~~3000~0~~2000~0~~55000~0~~1000~0~~3000~0~~4000~0~~"
'Barcode_Scaned str2
'str2 = "aa999692300026210   0020120000000080087000~0~~2000~0~~4000~0~~8000~0~~29000~0~~4000~0~~2000~0~~6000~0~~9000~0~~2000~0~~1000~0~~24000~0~~108000~0~~7000~0~~2000~0~~117000~0~Hoa Linh~22/10/2013</S></S01-3>"
'Barcode_Scaned str2

''BCTC bo QD16:
'str2 = "aa999202300026210   00201200200200100801/0118/04/2007<S01><S>V.01~2121111~2111111~V.02~32222222222~12332~V.03~13648~25838~~12313~13213~~123~312~~1212~12313~V.04~3756~4467~~1332~234~~2424~4233~V.05~23423~1111~V.06~424~1251~~324~1241~V.07~100~10~V.08~128~27~~123~23~~3~3~~2~1~V.09~181374~75373~~2324~1323~~13313~31313~~31313~31313~~131313~313~~3111~11111~~257437~"
'Barcode_Scaned str2
'str2 = "aa999202300026210   0020120020020020081340646~V.10~24646~13656~~12323~12323~~12323~1333~V.11~232446~1322222~~1123~11111~~231323~1311111~V.12~345~4768~~222~222~~123~4546~V.13~1253~66243~~677~65567~~576~676~V.14~134686~567666~V.14.2~5655~555~~123425~6366~V.22.1~4464~566~V.14~575~552425~V.15~55~464~V.14.3~567~7754~~32224959462~4206065~V.16~0~2421414~V.17~0~1474~"
'Barcode_Scaned str2
'str2 = "aa999202300026210   002012002002003008~0~1232~~0~242~V.18~0~1412~V.05~0~24~V.19~0~114~V.20~0~11~V.22~0~1440342~~0~1424242~V.22.2~0~2414~V.21~0~1241~V.21~0~12445~~0~3864791~V.23~0~21479~~0~20714~~0~141~~0~141~~0~1414~~0~1414~~0~14141~~0~3463~~0~364~~0~333~~0~35~~0~33~~0~0~~0~3886270~~0~4934~~0~4445~~0~255~~0~234~~0~276~~0~196~~0~80~duan~26/11/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999202300026210   002012002002004008<S01-1><S>VI.24~10000~0~VI.25~3000~0~~7000~0~~4000~0~~1000~0~VI.26~3000~0~VI.27~111~0~VI.28~333~0~VI.29~444~0~~222~0~~555~0~VI.31~-333~0~VI.30~777~0~VI.32~222~0~~11110~0~~129~0~~10981~0~~234~0~~156~0~VI.33~390~0~~10591~0~~267~0~~148~0~duan~26/11/2013</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999202300026210   002012002002005008<S01-2><S>~6000~0~~1200~0~~890~0~~670~0~~689~0~~456~0~~324~0~~134~0~~10363~0~~345~0~~456~0~~278~0~~189~0~~896~0~~456~0~~376~0~~267~0~~856~0~~348~0~~249~0~~876~0~~567~0~~854~0~~17376~0"
'Barcode_Scaned str2
'str2 = "aa999202300026210   002012002002006008~~623~0~~463~0~~765~0~~487~0~~462~0~~345~0~~261~0~~654~0~~865~0~~4925~0~~752~0~~569~0~~678~0~~753~0~~654~0~~765~0~~4171~0~~26472~0~~756~0~~453~0~~27681~0~duan~26/11/2013</S></S01-2>"
'Barcode_Scaned str2
'str2 = "aa999202300026210   002012002002007008<S01-3><S>~6000~0~~3300~0~~222~0~~333~0~~444~0~~555~0~~666~0~~777~0~~888~0~~999~0~~111~0~~233~0~~456~0~~234~0~~667~0~~446~0~~246~0~~2346~0~~3436~0~~2436~0~~6543~0~~5442~0~~6523~0~~564~0~~0~0~~43867~0~~3454~0~~24332~"
'Barcode_Scaned str2
'str2 = "aa999202300026210   0020120020020080080~~71653~0~~4534~0~~634~0~~32423~0~~12342~0~~23453~0~~2332~0~~3444~0~~2342~0~~1235~0~~2342~0~~345342~0~~23423~0~~21432~0~~124~0~~12~0~~451~0~~390784~0~~464779~0~~1243~0~~1435~0~~467457~0~duan~26/11/2013</S></S01-3>"
'Barcode_Scaned str2
'
''BCTC bo QD95:
'str2 = "aa999222300026210   00201200300300100901/0101/01/1900<S01><S>~143~0~~3~0~V.01~1~0~~2~0~V.02~7~0~~3~0~~4~0~V.11~46~0~~5~0~~6~0~~7~0~~8~0~~9~0~~11~0~V.02~12~0~~75~0~~13~0~~14~0~~15~0~~16~0~~17~0~~747~0~V.11~103~0~~18~0~~19~0~~21~0~~22~0~~23~0~~190~0~V.05~49~"
'Barcode_Scaned str2
'str2 = "aa999222300026210   0020120030030020090~~24~0~~25~0~~53~0~~26~0~~27~0~V.06~57~0~~28~0~~29~0~~31~0~~65~0~~32~0~~33~0~~219~0~~34~0~~35~0~~73~0~~37~0~~36~0~V.04~38~0~~39~0~~170~0~V.07~41~0~V.09~42~0~V.10~43~0~~44~0~~890~0~~593~0~~316~0~~15~0~~23~0~~24~0~V.08~"
'Barcode_Scaned str2
'str2 = "aa999222300026210   00201200300300300928~0~~12~0~V.12~11~0~~18~0~~26~0~~13~0~~21~0~~24~0~~26~0~~19~0~~26~0~~30~0~~277~0~~24~0~V.14~17~0~~26~0~V.15~13~0~V.09~23~0~~35~0~~27~0~~54~0~~23~0~~35~0~~297~0~~297~0~~25~0~~27~0~~28~0~~29~0~~24~0~~25~0~~31~0~~25~0~~5"
'Barcode_Scaned str2
'str2 = "aa999222300026210   0020120030030040096~0~~27~0~~890~0~~100~0~~20~0~~32~0~~34~0~~35.6~0~~36~0~~45~0~~42~0~~41~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~~0~0~Kh∏nh Ph≠¨ng~11/12/2013</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999222300026210   002012003003005009<S01-1><S>~936~0~~400~0~~100~0~~200~0~~32~0~~34~0~~45~0~~46~0~~47~0~~32~0~~31~0~~905~0~~34~0~~871~0~~38~0~~833~0~~39~0~~21~0~~18~0~~851~0~VI.1~23~0~VI.2~25~0~~803~0~~26~0~Kh∏nh Ph≠¨ng~11/12/2013</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999222300026210   002012003003006009<S01-2><S>~7000~0~~100~0~~200~0~~300~0~~400~0~~500~0~~600~0~~700~0~~800~0~~900~0~~234~0~~345~0~~357~0~~12436~0~~678~0~~456~0~~367~0~~654~0~~743~0~"
'Barcode_Scaned str2
'str2 = "aa999222300026210   002012003003007009~821~0~~321~0~~4040~0~~654~0~~732~0~~821~0~~965~0~~876~0~~521~0~~4569~0~~21045~0~~432~0~~654~0~VII.34~22131~0~Kh∏nh Ph≠¨ng~11/12/2013</S></S01-2>"
'Barcode_Scaned str2
'str2 = "aa999222300026210   002012003003008009<S01-3><S>~600~0~~123~0~~234~0~~235~0~~267~0~~289~0~~1748~0~~345~0~~4~0~~345~0~~123~0~~234~0~~345~0~~456~0~~567~0~~2419~0~~678~0~~789~0~~125~0~~167~"
'Barcode_Scaned str2
'str2 = "aa999222300026210   0020120030030090090~~278~0~~564~0~~345~0~~2946~0~~278~0~~457~0~~976~0~~267~0~~378~0~~231~0~~2587~0~~7952~0~~276~0~~285~0~~8513~0~Kh∏nh Ph≠¨ng~11/12/2013</S></S01-3>"
'Barcode_Scaned str2

''BCTC bo QD48:
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

'
''01/GTGT:
'str2 = "aa999012300121087   10201300300300101601/0114/06/2006<S01><S>0010011000</S><S>~32222000~24000000~2100000~1328400~5000000~36000000~2500000~6000000~10000000~500000~20000000~2000000~4"
'Barcode_Scaned str2
'str2 = "aa999012300121087   1020130030030020161000000~2500000~1171600~500000~400000~3000000~0~0~0~33950400~40000~33910400</S><S>Huy“n Tr©m~KTV~Hoµng Lan~19/12/2013~1~~~1701~~~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003003016<S01_1><S>01GTKT~01GTKT~001~012~11/10/2013~Nam~~~5000000~0~</S><S>02GTTT~02GTTT~002~0345~12/10/2013~Chi’n~~~6000000~0~</S><S>06HDXK~06HDXK~003~0567~13/10/2013~Minh~~~10000000~50"
'Barcode_Scaned str2
'str2 = "aa999012300121087   1020130030030040160000~</S><S>07KPTQ~07KPTQ~004~0345~14/10/2013~Lan~~~20000000~2000000~</S><S>03XKNB~03XKNB~005~0123~15/10/2013~Hπnh~~~4000000~20000~</S><S>41000000~2500000~36000000</S></S01_1>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003005016<S01_2><S>02GTTT~02GTTT~006~567~10/10/2013~Giang~~~6000000~5~300000~</S><S>06HDXK~06HDXK~007~345~11/10/2013~Tr≠Íng~~~10000000~10~1000000~</S><S>01GTKT~01GTKT~008~2"
'Barcode_Scaned str2
'str2 = "aa999012300121087   10201300300300601634~12/10/2013~Ph≠¨ng~~~8000000~10~800000~</S><S>~~~~~~~~0~0~0~</S><S>04HGDL~04HGDL~009~123~19/10/2013~Kh∏nh~~~1000000~5~50000~</S><S>24000000~2100000</S></S01_2>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003007016<S01_3><S>001~12/10/2013~usd~2000000~CK~10 ngµy~S001~13/10/201"
'Barcode_Scaned str2
'str2 = "aa999012300121087   1020130030030080163~usd~3000000~003~15/10/2013~usd~1500000~009~16/10/2013~usd~60"
'Barcode_Scaned str2
'str2 = "aa999012300121087   10201300300300901600000~008~12/10/2013~usd~4000000~004~11/10/2013~usd~2500000~00"
'Barcode_Scaned str2
'str2 = "aa999012300121087   1020130030030100162~17/10/2013~usd~2700000~ch¯ng tı~ghi chÛ</S><S></S></S01_3>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003011016<S01_4A><S>2100000~300000~1000000~800000~41000000~36000000~87.8~800000~702400</S></S01_4A>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003012016<S01_4B><S>2012~5600000~5000000~400000~200000~100000000~30000000~30~1200000~360000~34000~326000</S></S01_4B>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003013016<S01_5><S>CT001~12/10/2013~CÙc thu’~~3000000</S></S01_5>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003014016<S01_6><S>c¨ sÎ 1~0102030405~10303~3000000~12000000"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003015016~15000000~270000~0</S><S>0~3000000~0~0</S></S01_6>"
'Barcode_Scaned str2
'str2 = "aa999012300121087   102013003003016016<S01_7><S>mez~Chi’c~200000~150000~cﬂn 50.000</S><S>honda~Chi’c~500000~480000~cﬂn 20</S></S01_7>"
'Barcode_Scaned str2
'
''02/GTGT:
'str2 = "aa999022300121087   10201300100100100201/0114/06/2006<S01><S>0010011000</S><S>3000000~1200000~5000000~500000~30000~20000~50000~40000~480000~120000~4560000~40000~30000~20000~4510000</S><S>Huy“n Tr©m~Hoµng Lan~KTV~19/12/2013~1~~~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999022300121087   102013001001002002<S01_2><S>02GTTT~02GTTT~001~034~15/10/2013~lan h≠¨ng~0102030405~game~5000000~10~500000~</S><S>5000000~500000</S></S01_2>"
'Barcode_Scaned str2
'
''03/GTGT
'str2 = "aa999042300121087   10201300200300100101/0114/06/2006<S01><S>0010011000</S><S>3000000~2000000~1000000~140000~200000~-1940000~0</S><S>Huy“n Tr©m~Hoµng Lan~KTV~19/12/2013~1~~~0</S></S01>"
'Barcode_Scaned str2
'
''04/GTGT:
'str2 = "aa999712300121087   10201300100100100301/0101/01/1900<S01><S>0010011000</S><S>20000000~3000000~30000~4000000~200000~5000000~150000~6000000~120000~18000000~500000~38000000~500000</S><S>Huy“n Tr©m~Hoµng Lan~KTV~19/12/2013~1~~~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999712300121087   102013001001002003<S01_1><S>01GTKT~~003~05442~12/10/2013~nga~~~40000000~</S><S>07KPTQ~~004~03243~13/10/2013~hπnh~~~5000000~</S><S>02GTTT~~001~024234~14/10/2013~linh~~~6000000~</S><S>04HGDL~~002~"
'Barcode_Scaned str2
'str2 = "aa999712300121087   102013001001003003023423~15/10/2013~∏nh~~~7000000~~06HDXK~~005~023434~16/10/2013~ph≠¨ng~~~8000000~</S><S>03XKNB~~006~2423~18/10/2013~kh´i~~~2000000~</S><S>68000000~40000000~28000000</S></S01_1>"
'Barcode_Scaned str2

''05/GTGT:
'str2 = "aa999722300121087   10201300200200100101/0114/06/2006<S01><S>0010011000</S><S>4000000~7000000~40000~140000~180000</S><S>Huy“n Tr©m~Hoµng Lan~KTV~19/12/2013~1~~~~</S></S01>"
'Barcode_Scaned str2

''01A/TNDN + PL dinh kem
'str2 = "aa999112300121087   03201300600600100301/0114/06/2006<S01><S>0010011000</S><S>300000~200000~100000~23000~12000~111000~10000~12000~89000~30000~25000~34000~20.897~12000~22"
'Barcode_Scaned str2
'str2 = "aa999112300121087   032013006006002003.678~21326~15326~7000~8000~6000~x~03~15/11/2013~4000~2000</S><S>x~x</S><S>Huy“n Tr©m~KTV~Hoµng Lan~19/12/2013~1~0~~1052</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999112300121087   032013006006003003<S01-1><S>6000</S><S>doanh nghi÷p ab~0010011000~45.67~2740~10304~c¨ sÎ 1~0010011000-001~23.88~1433~10304~c¨ sÎ 2~0010011000-002~30.45~1827~10304</S></S01-1>"
'Barcode_Scaned str2
'
''KHBS 01A/TNDN:
'str2 = "bs999112300121087   03201301101100100401/0114/06/2006<S01><S>0010011000</S><S>300000~200000~100000~230000~12000~318000~100000~12000~206000~80000~50000~76000~20.897~12000~2"
'Barcode_Scaned str2
'str2 = "bs999112300121087   0320130110110020042.678~45603~20000~4000~3000~25603~x~03~15/11/2013~4000~21603</S><S>x~x</S><S>Huy“n Tr©m~KTV~Hoµng Lan~19/12/2013~0~1~~1052</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999112300121087   032013011011003004<SKHBS><S>Thu’ TNDN ph∏t sinh~34~21326~45603~24277</S><S>Thu’ TNDN d˘ ki’n mi‘n, gi∂m~35~15326~20000~467"
'Barcode_Scaned str2
'str2 = "bs999112300121087   0320130110110040044</S><S>19/12/2013~49~972~4200~lh005~15/10/2013~10100~10107~9~100000~l˝ do nhi“u~0~0~19603</S></SKHBS>"
'Barcode_Scaned str2
'
''01B/TNDN + PL dinh kem:
'str2 = "aa999122300121087   03201300200200100301/0114/06/2006<S01><S>0010011000</S><S>x~x~560000~520000~400000~100000~20000~40000~23.786~~20.567~22.345~17.678~~~25"
'Barcode_Scaned str2
'str2 = "aa999122300121087   032013002002002003724~25724~~12000~1000~2000~13724~x~~19/11/2013~40~13684</S><S>Hoµng Lan~19/12/2013~Huy“n Tr©m~KTV~1~~1052</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999122300121087   032013002002003003<S01-1><S>13724</S><S>doanh nghi÷p mh~3600247325~45.78~6283~10703~doanh nghi÷p ab~0102030405~54.22~7441~10307</S></S01-1>"
'Barcode_Scaned str2
'
''KHBS 01B/TNDN:
'str2 = "bs999122300121087   03201300300300100401/0114/06/2006<S01><S>0010011000</S><S>x~x~540000~500000~380000~100000~20000~40000~23.786~~20.567~22.345~17.678~~12.896~2"
'Barcode_Scaned str2
'str2 = "bs999122300121087   0320130030030020045973~24746~1227~8000~1000~2000~17973~x~~19/11/2013~40~17933</S><S>Hoµng Lan~19/12/2013~Huy“n Tr©m~KTV~0~1~1052</S></S01>"
'Barcode_Scaned str2
'str2 = "bs999122300121087   032013003003003004<SKHBS><S>Thu’ TNDN cÒa hoπt ÆÈng Æ≠Óc ≠u Æ∑i v“ thu’ su t~29~0~1227~1227~Thu’ thu nhÀp doanh nghi÷p d˘ ki’n mi‘n, gi∂m~30~12000~8000~-4000</S><S>Thu’ T"
'Barcode_Scaned str2
'str2 = "bs999122300121087   032013003003004004NDN cÒa hoπt ÆÈng kh´ng ≠u Æ∑i~28~25724~24746~-978</S><S>19/12/2013~49~104~5000~kh002~16/10/2013~10100~10101~12~30000~NNT nÈp thu’~0~0~4249</S></SKHBS>"
'Barcode_Scaned str2
'
''02/TNDN:
'str2 = "aa999732300121087   03201300300300100201/0114/06/2006<S02><S>0010011000</S><S>2000000~690000~120000~200000~300000~20000~10000~40000~1310000~500000~810000~22~178200~32000~146200~2000000~5000"
'Barcode_Scaned str2
'str2 = "aa999732300121087   0320130030030020020~22~1~429000</S><S>1~~tÊ ch¯c s˘ ki÷n~012345678~Ba ß◊nh- HN~123~14/07/2013~30/07/2013</S><S>Huy“n Tr©m~KTV~Hoµng Lan~20/12/2013~1~~~1053~~x</S></S02>"
'Barcode_Scaned str2
'
''01/KK-XS:
'str2 = "aa999482300121087   11201300100100100101/0101/01/2010<S01><S>0010011000</S><S>500000~400000~200000</S><S>Hoµng Lan~Huy“n Tr©m~KTV~20/12/2013~1~~</S></S01>"
'Barcode_Scaned str2
'
''01/KK-BHDC:
'str2 = "aa999252300121087   11201300100100100101/0101/01/1900<S01><S>0010011000</S><S>4000000~200000~100000~300000~30000~100000~10000~130000~30000~21000~161000</S><S>Hoµng Lan~20/12/2013~Huy“n Tr©m~KTV~1~~~0</S></S01>"
'Barcode_Scaned str2

''02/TNCN:
'str2 = "aa999152300121087   11201300100100100101/0101/01/2010<S01><S>0010011000</S><S>5000~2000~120000~300000~400000~100000~340000~130000~100000~200000~60000~10000~40000</S><S>Hoµng Lan~20/12/2013~Huy“nTr©m~KTV~1~~</S></S01>"
'Barcode_Scaned str2
'
''03/TNCN:
'str2 = "aa999502300121087   11201300100100100201/0101/01/2010<S01><S>0010011000</S><S>tÊ ch¯c khai thay ABC~0102030405~ßinh Li÷t - Hoµn Ki’m - HN~Hoµn Ki’m~HN~0987689965~015678889~abc@yahoo.com</S>"
'Barcode_Scaned str2
'str2 = "aa999502300121087   112013001001002002<S>3000000~150000~1200000~1200~6000000~300000~5000000~500000~2000000~20000~1000000~1000</S><S>Hoµng Lan~20/12/2013~Huy“n Tr©m~KTV~1~~</S></S01>"
'Barcode_Scaned str2
'
''07/TNCN:
'str2 = "aa999362300121087   06201300100100100101/0101/01/2010<S07><S>0010011000</S><S>~30000000~2000000~6700000~4000000~1600000~500000~400000~200000~21300000~2610000~120000~5220~2604780~4000000~20~800000</S><S>Hoµng Lan~20/12/2013~Huy“n Tr©m~KTV~1~~~0</S></S07>"
'Barcode_Scaned str2
'
''08/KK-TNCN:
'str2 = "aa999742300121087   03201300100100100101/0101/01/1900<S01><S>0010011000</S><S>40000000~12000000~28000000~13650000~9000000~3600000~500000~300000~250000~14350000~717500~100000~1281~716219</S><S>Hoµng Lan~20/12/2013~Huy“n Tr©m~KTV~1~</S></S01>"
'Barcode_Scaned str2
'
''08A/TNCN:
'str2 = "aa999752300121087   03201300100100100201/0101/01/1900<S01><S>0010011000</S><S>80000000~5000000~60000000</S><S>Kh∏nh~0102030405~45.78~34335000~27468000~300000~34035000~2855250~1142100"
'Barcode_Scaned str2
'str2 = "aa999752300121087   032013001001002002~1713150~Linh~2300121087~54.22~40665000~32532000~400000~40265000~3789750~1515900~2273850</S><S>Hoµng Lan~20/12/2013~Huy“n Tr©m~KTV~1~</S></S01>"
'Barcode_Scaned str2
'
''01/NTNN:
'str2 = "aa999702300121087   10201300100100100201/0101/01/1900<S01><S>0010011000</S><S>nÈi dung 1~0102030405~hd001~4000000~12/10/2013~5000000~12.78~10~63900~3400000~34.67~200000~978780~1042680~nÈi dung 1~0010011000~hd002~6000000~15"
'Barcode_Scaned str2
'str2 = "aa999702300121087   102013001001002002/10/2013~4000000~11.45~5~22900~2890000~27.34~100000~690126~713026</S><S>40000000~100000~8000000~500000~1000000~1100000</S><S>~</S><S>Huy“nTr©m~KTV~Hoµng Lan~20/12/2013~1~~~</S></S01>"
'Barcode_Scaned str2
'
''01/TTDB + PL dinh kem:
'str2 = "aa999052300121087   10201300300300100501/0101/01/1900<S01><S>0010011000</S><S>~150000~100000.00~30000~0~20001</S><S>10200~L›t~34452.79~100000~66667.00~50.0~20000~0~13334~10300~L›t~456.78~50000~33333.00~50.0~10000~0~6667</S><S>20000~17391.00~2000~0~6"
'Barcode_Scaned str2
'str2 = "aa999052300121087   10201300300300200509</S><S>20500~~237.67~20000~17391.00~15.0~2000~0~609</S><S>0</S><S>~~0.00~0</S><S>~~0.00~0</S><S>~~0.00~0</S><S>170000~117391.00~32000~0~20610</S><S>Hoµng Lan~Huy“n Tr©m~KTV~20/12/2013~1~~0~20610~~0</S></S01>"
'Barcode_Scaned str2
'str2 = "aa999052300121087   102013003003003005<S01-1><S>KH001~S001~13/10/2013~Kh∏nh~10300~hµng h„a 1~346.35~2391.34~828240.609~KH001~S002~13/10/2013~Linh~10101~hµng h„a 2~123.67~1456.28~180098.1476</S><S>    Bia (50 %)~346.35~828241~    X◊ gµ (65 %)~123.67~180098</S></S01-1>"
'Barcode_Scaned str2
'str2 = "aa999052300121087   102013003003004005<S01-2><S>KH002~S001~16/10/2013~nguy™n li÷u 1~345.78~20000~1000~19000~KH003~S002~17/10/2013~"
'Barcode_Scaned str2
'str2 = "aa999052300121087   102013003003005005nguy™n li÷u 2~238.97~30000~1500~28500</S><S>m∆t hµng A~100~nguy™n li÷u ab~32~10</S></S01-2>"
'Barcode_Scaned str2
'
''01/TAIN
'str2 = "aa999062300121087   10201300200200100201/0114/06/2006<S01><S>0010011000</S><S>010103~Kg~345.789~12345.68~11.000~0~23.87~010208~Kg~456.782~0,00~0~30000~89.34</S><S>010213~Kg~543.367~2897.67~12.0"
'Barcode_Scaned str2
'str2 = "aa999062300121087   10201300200200200200~0~345.27</S><S>010208~Kg~653.715~3762.78~10.000~0~5987.34~010208~Kg~433.678~3243.89~10.000~0~3243.56</S><S>Huy“n Tr©m~KTV~~20/12/2013~1~~0~0~</S></S01>"
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
    frmSystem.chkQuetBangKe.value = False

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
                lValue = .value
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
            If (Trim(idToKhai) = "53" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "37" And Val(Mid(strPrefix, 21, 4)) > 2009) _
                Or (Trim(idToKhai) = "54" And Val(Mid(strPrefix, 19, 2)) > 1 And Val(Mid(strPrefix, 21, 4)) > 2009) Or (Trim(idToKhai) = "38" And Val(Mid(strPrefix, 21, 4)) > 2009) Then
                DisplayMessage "0094", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '07072011 TT28
        ' Khong nhan cac to khai theo mau cua
        idToKhai = Mid(strPrefix, 4, 2)

        If (Val(Left$(strPrefix, 3)) < 300) Then
            If Trim(idToKhai) = "01" Or Trim(idToKhai) = "02" Or Trim(idToKhai) = "04" Or Trim(idToKhai) = "11" Or Trim(idToKhai) = "12" Or Trim(idToKhai) = "46" Or Trim(idToKhai) = "47" Or Trim(idToKhai) = "48" Or Trim(idToKhai) = "49" Or Trim(idToKhai) = "15" Or Trim(idToKhai) = "16" Or Trim(idToKhai) = "50" Or Trim(idToKhai) = "51" _
            Or Trim(idToKhai) = "36" Or Trim(idToKhai) = "70" Or Trim(idToKhai) = "06" Or Trim(idToKhai) = "05" Then
                DisplayMessage "0113", msOKOnly, miInformation
                Exit Sub
            End If
        End If
        
        '06012012 TT28
        ' Khong nhan cac to khai theo mau cu GD2
        If (Val(Left$(strPrefix, 3)) < 310) Then
            If Trim$(idToKhai) = "71" Or Trim$(idToKhai) = "72" Or Trim$(idToKhai) = "73" Or Trim$(idToKhai) = "03" Or Trim$(idToKhai) = "74" Or Trim$(idToKhai) = "75" Or Trim$(idToKhai) = "80" Or Trim$(idToKhai) = "81" Or Trim$(idToKhai) = "82" Or Trim$(idToKhai) = "17" Or Trim$(idToKhai) = "42" Or Trim$(idToKhai) = "43" _
            Or Trim$(idToKhai) = "59" Or Trim$(idToKhai) = "76" Or Trim$(idToKhai) = "41" Or Trim$(idToKhai) = "77" Or Trim$(idToKhai) = "86" Or Trim$(idToKhai) = "87" Or Trim$(idToKhai) = "89" Then
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
        'If Trim(idToKhai) = "08" Or Trim(idToKhai) = "24" Then
        'nvsu -- reOpen(01/BCTL_DK)
        If Trim(idToKhai) = "08" Then
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
                ProgressBar1.value = 0
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

                If frmSystem.chkSaveQuestion.value Then
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
                ProgressBar1.value = 0
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

                If frmSystem.chkSaveQuestion.value Then
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
    'remove 24,25,26
    If (idToKhaiCheck >= 27 And idToKhaiCheck <= 35) Or (idToKhaiCheck >= 55 And idToKhaiCheck <= 58) Or (idToKhaiCheck >= 18 And idToKhaiCheck <= 21) Or idToKhaiCheck = 69 Then
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
'        ElseIf idToKhaiCheck = 11 Then
'             If ((lDataNo + 1 > lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 7)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 1
'                Exit Sub
'            End If
'            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
'            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
'            If ((lDataNo + 1 < lElementsNo And lDataNo <> 7) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 7)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 2
'                Exit Sub
'            End If
'        ElseIf idToKhaiCheck = 12 Then
'             If ((lDataNo + 1 > lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 > lElementsNo) And lDataNo = 6)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 1
'                Exit Sub
'            End If
'            ' Truong hop chuoi ma vach it chi tieu hon so chi tieu trong template
'            'If (UBound(arrStrValue) + 1 < lElementsNo) And isSheetTk Then
'            If ((lDataNo + 1 < lElementsNo And lDataNo <> 6) Or ((lDataNo + 2 < lElementsNo) And lDataNo = 6)) And isSheetTk Then
'                blnValidData = False
'                checkSoCT = 2
'                Exit Sub
'            End If
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

Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDOMdata As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDOMdata Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")
    
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
        'set ThreeMonths cho to khai thang/quy
        If strID = "01" Or strID = "02" Or strID = "04" Or strID = "71" Or strID = "95" Or strID = "36" Or strID = "68" Or strID = "25" Then
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
    strTaxIDString = strTaxID
    'strTaxID = ""
    'Neu khong co NNT thi quit luon
    If (strTaxID = vbNullString Or strTaxID = "") Then
        InitParameters = False
        MessageBox "0146", msOKOnly, miWarning
        Exit Function
    End If
    
    Dim strIDBCTC As String
    strIDBCTC = Left$(strTaxReportInfo, 2)
     If (Val(strIDBCTC) = 27 Or Val(strIDBCTC) = 28 Or Val(strIDBCTC) = 29 _
            Or Val(strIDBCTC) = 30 Or Val(strIDBCTC) = 31 Or Val(strIDBCTC) = 32 Or Val(strIDBCTC) = 33 Or Val(strIDBCTC) = 34 Or Val(strIDBCTC) = 35 _
            Or Val(strIDBCTC) = 55 Or Val(strIDBCTC) = 56 Or Val(strIDBCTC) = 57 Or Val(strIDBCTC) = 58 Or Val(strIDBCTC) = 18 Or Val(strIDBCTC) = 19 _
            Or Val(strIDBCTC) = 20 Or Val(strIDBCTC) = 21 Or Val(strIDBCTC) = 69) Then
        Set rsTaxInfor = GetTaxInfoBCTC(strTaxID, blnConnected)
    Else
        Set rsTaxInfor = GetTaxInfo(strTaxID, blnConnected)
    End If
    ' end
    
'     'Connect DB fail
'    If Not blnConnected Then _
'        Exit Function
    
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
    
    If rsTaxInfor.Fields(0) = "05" Then
        InitParameters = False
        MessageBox "0088", msOKOnly, miCriticalError
        Exit Function
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
    
'    If rsTaxInfor.Fields(0) = "05" Then
'        InitParameters = False
'        MessageBox "0088", msOKOnly, miCriticalError
'        Exit Function
'    End If

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
    
    If Val(strIDBCTC) = 1 Or Val(strIDBCTC) = 2 Or Val(strIDBCTC) = 25 Or Val(strIDBCTC) = 26 Or Val(strIDBCTC) = 4 Or Val(strIDBCTC) = 71 Or Val(strIDBCTC) = 36 Or Val(strIDBCTC) = 68 Then
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
        
        'Xu ly rieng cho to khai thang/quy
        If Val(strIDBCTC) = 1 Or Val(strIDBCTC) = 2 Or Val(strIDBCTC) = 25 Or Val(strIDBCTC) = 26 Or Val(strIDBCTC) = 4 Or Val(strIDBCTC) = 71 Or Val(strIDBCTC) = 95 Or Val(strIDBCTC) = 36 Then
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
'        If Val(strID) = 74 Then
'            arrCT = Split(strData, "~")
'            If Trim(arrCT(2)) <> "" Then
'                TuNgay = arrCT(2)
'                DenNgay = arrCT(3)
'                isTKThang = True
'            End If
'
'        End If
' 08A/TNCN
'        If Val(strID) = 75 Then
'            arrCT = Split(strData, "~")
'            If Trim(arrCT(1)) <> "" Then
'                TuNgay = Right$(arrCT(0), 7)
'                DenNgay = arrCT(1)
'                isTKThang = True
'            End If
'
'        End If
        
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
    strTaxDLTIDString = strTaxDLID
    'strTaxDLID = ""
    If strTaxDLID <> vbNullString Or strTaxDLID <> "" Then
        
        Set rsTaxDLInfor = GetTaxDLInfo(strMST, strTaxDLID, blnDLConnected)
    Else
        Set rsTaxDLInfor = Nothing
    End If
    
        
    ' DLT ngung hoat dong canh bao va thoat
    If Not rsTaxDLInfor Is Nothing Then
        If rsTaxDLInfor.Fields(7).value = "01" Then
            MessageBox "0144", msOKOnly, miWarning
            frmSystem.MousePointer = vbDefault
            Me.MousePointer = vbDefault
            Exit Function
        End If
    End If
        
    
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
                'If Not (rsTaxDLInfor Is Nothing Or rsTaxDLInfor.Fields.Count = 0) Then
                If Not (rsTaxDLInfor Is Nothing) Then
                    If Not (rsTaxDLInfor.Fields.Count = 0) Then
                        If Not objTaxBusiness Is Nothing Then
                            objTaxBusiness.strTenDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(0).value), "", rsTaxDLInfor.Fields(0).value), TCVN, UNICODE)
                            objTaxBusiness.strDiaChiDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(1).value), "", rsTaxDLInfor.Fields(1).value), TCVN, UNICODE)
                            objTaxBusiness.strDienThoaiDL = IIf(IsNull(rsTaxDLInfor.Fields(2).value), "", rsTaxDLInfor.Fields(2).value)
                            objTaxBusiness.strFaxDL = IIf(IsNull(rsTaxDLInfor.Fields(3).value), "", rsTaxDLInfor.Fields(3).value)
                            objTaxBusiness.strEmailDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(4).value), "", rsTaxDLInfor.Fields(4).value), TCVN, UNICODE)
                            objTaxBusiness.strSoHopDongDL = TAX_Utilities_Srv_New.Convert(IIf(IsNull(rsTaxDLInfor.Fields(5).value), "", rsTaxDLInfor.Fields(5).value), TCVN, UNICODE)
                            objTaxBusiness.strNgayHopDongDL = IIf(IsNull(rsTaxDLInfor.Fields(6).value), "", rsTaxDLInfor.Fields(6).value)
                        End If
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
    
    'Load co quan thue KHBS
    With fpSpread1
        Dim CQT_CAPCUC    As Variant
        Dim CQT_HOANTHUE  As Variant
        Dim tCQT_CAPCUC   As String
        Dim tCQT_HOANTHUE As String

        If TAX_Utilities_Srv_New.NodeValidity.hasChildNodes Then
            If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(TAX_Utilities_Srv_New.NodeValidity.childNodes.length - 1), "ID") = "KHBS" Then
                If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(TAX_Utilities_Srv_New.NodeValidity.childNodes.length - 1), "Active") = "1" Then
                    .Sheet = .SheetCount - 1
                    .GetText .ColLetterToNumber("BI"), .MaxRows - 15, CQT_CAPCUC
                    .GetText .ColLetterToNumber("BI"), .MaxRows - 13, CQT_HOANTHUE
                    DataDM CQT_CAPCUC, tCQT_CAPCUC
                    DataDM CQT_HOANTHUE, tCQT_HOANTHUE
                    .Col = .ColLetterToNumber("BE")

                    If tCQT_CAPCUC <> vbNullString Then
                        .Row = .MaxRows - 15
                        .Text = tCQT_CAPCUC

                    End If

                    If tCQT_HOANTHUE <> vbNullString Then
                        .Row = .MaxRows - 13
                        .Text = tCQT_HOANTHUE
 
                    End If
                End If
            End If
        End If
    
    End With
    
    ' set ma CQT
    If Not objTaxBusiness Is Nothing Then
        If (Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) >= 64 And Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) <= 68) Or Val(GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")) = 91 Then
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
        If Val(LoaiTk) = 70 Or Val(LoaiTk) = 71 Or Val(LoaiTk) = 72 Or Val(LoaiTk) = 73 Or Val(LoaiTk) = 74 Or Val(LoaiTk) = 77 Or Val(LoaiTk) = 3 Or Val(LoaiTk) = 75 Or Val(LoaiTk) = 80 Or Val(LoaiTk) = 81 Or Val(LoaiTk) = 82 Or Val(LoaiTk) = 86 Or Val(LoaiTk) = 87 Or Val(LoaiTk) = 89 Or Val(LoaiTk) = 17 Or Val(LoaiTk) = 42 Or Val(LoaiTk) = 43 Or Val(LoaiTk) = 59 Or Val(LoaiTk) = 76 Or Val(LoaiTk) = 41 Then
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
            frmSystem.chkQuetBangKe.value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    ' Loai to khai la GTGT hoa hong dai ly mau 02/GTGT thi cho hien check quet du lieu bang ke
    ElseIf (verToKhai = 0 And LoaiTk = "02") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    ' Loai to khai la Quyet toan thue TNCN mau 04/TNCN thi cho hien check quet du lieu bang ke
    ElseIf (verToKhai = 0 And LoaiTk = "17") Then
        If (TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1) Then
            frmSystem.chkQuetBangKe.Visible = True
        Else
            frmSystem.chkQuetBangKe.value = False
            frmSystem.chkQuetBangKe.Visible = False
        End If
    Else
        frmSystem.chkQuetBangKe.value = False
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
    
    If intCount > ProgressBar1.value Then
        ProgressBar1.value = ProgressBar1.value + 1
        lblFilePath.caption = ProgressBar1.value & " / " & ProgressBar1.max
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
        strMaNNT = strTaxIDString
        strResultNNT = GetDataFromESB("", "", "NNT")
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
 
'    'Neu khong co thong tin NNT thi exit luon
'    If (strTaxIDString = "" Or strTaxIDString = vbNullString) Then
'        Set rsReturn = Nothing
'        blnSuccess = False
'        Exit Function
'    End If

'    'Du lieu gia lap de test
'    Set xmlResultDLT = LoadXmlTemp("ResultDLTFromESB")
'    strResultDLT = "sdfsfds"

    If (strTaxIDDLString <> "" And strTaxIDDLString <> vbNullString) Then
        strMaDLT = strTaxIDDLString
        strResultDLT = GetDataFromESB("", "", "DLT")
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
                       If (MessageBox("0141", msOKOnly, miCriticalError) = mrNo) Then
                        Set rsReturn = Nothing
                        blnSuccess = False
                        Exit Function
                        End If
                End If
'                If (xmlResultDLT.getElementsByTagName("TrangThaiHoatDong").length > 0) Then
'                    If (xmlResultDLT.getElementsByTagName("TrangThaiHoatDong")(0).Text = "01") Then
'                            'MessageBox "0144", msOKOnly, miWarning
'                            Set rsReturn = Nothing
'                            blnSuccess = False
'                            Exit Function
'                    End If
'                End If
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
    rsReturn.Fields.Append "trang_thai", adChar, 2, adFldUpdatable
    
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
        rsReturn!trang_thai = xmlResultDLT.getElementsByTagName("TrangThaiHoatDong")(0).Text
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

'   'Du lieu gia lap de test
'    Set xmlResultNNT = LoadXmlTemp("ResultNNTFromESB")
'    strResultNNT = "test"

    If (strTaxIDString <> "" Or strTaxIDString <> vbNullString) Then
        strMaNNT = strTaxIDString
        strResultNNT = GetDataFromESB("", "", "NNT")
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

Function GenerateSQL_Details(xmlDOMdata As MSXML.DOMDocument, strSQL_DTL As String, vHdrID As Variant, lPos As Long) As String
    Dim xmlListSection As MSXML.IXMLDOMNodeList
    Dim xmlNodeSection As MSXML.IXMLDOMNode
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim lPosition As Long, strCondition As String
    Dim i As Long, j As Long, strLoaiDL As String
    
On Error GoTo ErrHandle
    Set xmlListSection = xmlDOMdata.getElementsByTagName("Section")
    For Each xmlNodeSection In xmlListSection
        If Trim(xmlNodeSection.Attributes.getNamedItem("Dynamic").nodeValue) = "1" Then
            iRowID = 0
            For i = 0 To xmlNodeSection.childNodes.length - 1
                iRowID = iRowID + 1
                For j = 0 To xmlNodeSection.childNodes(i).childNodes.length - 1
                    Set xmlAttribute = xmlDOMdata.createAttribute("RowID")
                    xmlAttribute.value = iRowID
                    Set xmlNode = xmlNodeSection.childNodes(i).childNodes(j).Attributes.setNamedItem(xmlAttribute)
                    Set xmlAttribute = Nothing
                Next
            Next
        End If
    Next
        
    strLoaiDL = Trim(TAX_Utilities_Srv_New.NodeValidity.childNodes(lPos).Attributes.getNamedItem("DataFile").nodeValue)
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
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
    Set xmlDOMdata = Nothing
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
        ProgressBar1.value = 0
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
    ProgressBar1.value = ProgressBar1.value + 1
    
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
Private Function formatMaToKhai(ByVal strID As String) As String
    Dim strTemp As String
    Dim strCode As String
    Dim strItem As String
    Dim strRetValue As String
    
    strCode = Mid$(strID, Len(strID) - 1, 2)
    strItem = Left$(strID, Len(strID) - 2)
    If (strCode = "11") Then
        strRetValue = "('" & strItem & "','" & strID & "','" & strItem & "13')"
    ElseIf strCode = "13" Then
        strRetValue = "('" & strItem & "','" & strItem & "11','" & strID & "')"
    Else
        strRetValue = "('" & strItem & "','" & strItem & "11','" & strID & "13')"
    End If
    formatMaToKhai = strRetValue
End Function

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
        If Not IsNull(rsTaxInfor.Fields(lCtrl + 1).value) Then
            'arrStrData(lCtrl) = clsConvert.Convert(rsTaxInfor.Fields(lCtrl + 1).Value, UNICODE, TCVN)
            arrStrData(lCtrl) = rsTaxInfor.Fields(lCtrl + 1).value
            
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
    Set XmlDocStandard = SetValueHeaderESB(XmlDocStandard)
    
    'Verify value for header
    XmlDocStandard.getElementsByTagName("ID_LINK")(0).Text = strMaToKhai
    XmlDocStandard.getElementsByTagName("SPARE3")(0).Text = strLoaiToKhai
    'end verify value
    
    ' Set value tag <add_info>
    XmlDocStandard.getElementsByTagName("ngay_nop_tk")(0).Text = Format(sNgayNopTK, "dd-mmm-yyyy hh:mm:ss")
    XmlDocStandard.getElementsByTagName("ky_lap_bo")(0).Text = sKyLapBo
    XmlDocStandard.getElementsByTagName("nguon_goc_tk")(0).Text = xmlConfig.getElementsByTagName("SENDER_CODE")(0).Text
    XmlDocStandard.getElementsByTagName("nguoi_nhan_tk")(0).Text = strUserID '& "." & xmlConfig.getElementsByTagName("CODE_OFFICE")(0).Text
    XmlDocStandard.getElementsByTagName("ngay_nhan_tk")(0).Text = strNgayHeThongSrv 'ConvertDate(Format(DateTime.Now, "dd/mm/yyyy"), True, "-")
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
        
        XmlDocStandard.getElementsByTagName("RECEIVER_CODE")(0).Text = "QLT_NTK_" & strMaTinhCoQuanThue
        XmlDocStandard.getElementsByTagName("RECEIVER_NAME")(0).Text = "Co so du lieu trung gian chua QLAC, BCTC"
        
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
        
        XmlDocStandard.getElementsByTagName("RECEIVER_CODE")(0).Text = "QLT_NTK_" & strMaTinhCoQuanThue
        XmlDocStandard.getElementsByTagName("RECEIVER_NAME")(0).Text = "Co so du lieu trung gian chua QLAC, BCTC"
        
        'Load template QHS
        tempQHSxml.Load GetAbsolutePath("..\InterfaceTemplates\xml\QHS.xml")
        XmlDocStandard.selectSingleNode("DATA/BODY/ROW/ADD_INFO").appendChild tempQHSxml.lastChild.lastChild

        If (GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "68") Then
            For nodeValIndex = 1 To TAX_Utilities_Srv_New.NodeValidity.childNodes.length
                Set nodeVal = TAX_Utilities_Srv_New.NodeValidity.childNodes(nodeValIndex)
                If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01_1") Then
                    XmlDocStandard.getElementsByTagName("PL_BK_01AC_01")(0).Text = "X"
                End If
                If (GetAttribute(nodeVal, "Active") = "1" And GetAttribute(nodeVal, "ID") = "01_2") Then
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


