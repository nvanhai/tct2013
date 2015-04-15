VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmKhaibosung 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5370
   ControlBox      =   0   'False
   HelpContextID   =   85
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnThoat 
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
      Left            =   2670
      TabIndex        =   2
      Top             =   2730
      Width           =   1200
   End
   Begin VB.CommandButton btnMo 
      Caption         =   "§ån&g ý"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2730
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2025
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   480
      Width           =   5055
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4695
         _Version        =   458752
         _ExtentX        =   8281
         _ExtentY        =   873
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmKhaibosung.frx":0000
         UserResize      =   1
      End
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4695
         _Version        =   458752
         _ExtentX        =   8281
         _ExtentY        =   873
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmKhaibosung.frx":0441
         UserResize      =   1
         Appearance      =   1
      End
      Begin FPUSpreadADO.fpSpread fpsNgaykhai 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   4695
         _Version        =   458752
         _ExtentX        =   8281
         _ExtentY        =   873
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   3
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmKhaibosung.frx":0803
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Chän tê khai cÇn khai bæ sung"
      BeginProperty Font 
         Name            =   "DS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image imgCaption 
      Height          =   345
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmKhaibosung"
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
Private Const f1dteTuNRow = 2
Private Const f1dteTuNCol = "C"
Private Const f4cboLoTRow = 2
Private Const f4cboLoTCol = "C"
Private Const mFormColor = -2147483633
Private Const mHeaderColor = 16709097
Private Const LOI_KY_HIEU_LUC = "1"
Private Const LOI_NGAY_BAT_DAU_NAM_TAI_CHINH = "2"
Private Const LOI_TU_NGAY_DEN_NGAY = "3"
Private lstryear As String
Private lstrMonth As String
Private lstrThreemonths As String
Private arrStrId() As String
Private Dtetun As String
Private Dteden As String
Private lSoBG As Long
Private arrCheckStatus() As Long
Private blnOpenInterfaces As Boolean
Private blnOnExit As Boolean                    'Bat su kien Click vao nut Thoat
Private arrStrXMLFileNames() As String
Private arrLngErrRows() As Long
Private lngRowFocus As Long
Private blnDKienTraCuu As Boolean               'Kiem tra dieu kien tra cuu co hop le hay ko
Dim objCvt As DateUtils
Private varFirstDay As String, varLastDay As String

Private Sub btnMo_Click()
Dim strId As String
Dim strDay As String
Dim strDateKHBS As String
Dim frmTK As frmInterfaces
Dim blnValidFinanceYear As Boolean
Dim fso As New FileSystemObject, fle As file
Dim lLoc As Long

    If checkValidate = False Then Exit Sub
    
    strId = arrStrId(fpsLoaiTK.TypeComboBoxCurSel)
    With fpsDkNgay
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        strDay = .Text
    End With
    
    'Initial parameters
    TAX_Utilities_v2.month = ""
    TAX_Utilities_v2.ThreeMonths = ""
    TAX_Utilities_v2.Year = ""
    TAX_Utilities_v2.FirstDay = ""
    TAX_Utilities_v2.LastDay = ""
    TAX_Utilities_v2.NodeMenu = getNode(CStr(strId))
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
            TAX_Utilities_v2.month = Mid$(CStr(strDay), 1, 2)
            TAX_Utilities_v2.Year = Mid$(CStr(strDay), 4, 4)
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
            TAX_Utilities_v2.ThreeMonths = CInt(Mid$(CStr(strDay), 1, 2))
            TAX_Utilities_v2.Year = Mid$(CStr(strDay), 4, 4)
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" Then
            Dim dDauKyNam As Date
            Dim dCuoiKyNam As Date
                        
            Dim objDateUtils As DateUtils
            iNgayTaiChinh = 1
            iThangTaiChinh = 1
            dDauKyNam = DateSerial(CInt(Right(strDay, 4)), iThangTaiChinh, iNgayTaiChinh)
            dCuoiKyNam = DateAdd("M", 12, dDauKyNam) - 1
            Set objDateUtils = New DateUtils
            varFirstDay = objDateUtils.ToString(dDauKyNam, "DD/MM/YYYY")
            varLastDay = objDateUtils.ToString(dCuoiKyNam, "DD/MM/YYYY")
            
            ' hlnam Edit - Begin
            ' Doi voi loai to khai quyet toan (03_TNDN, 04_TNCN) co ky quet toan la Tu ngay - Den ngay
            ' Phai lay duoc ky quet toan cua to quuyet toan thi moi bat dau cho phep sua doi bo sung
            
            If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").Text = "03") Then
                For Each fle In fso.GetFolder(TAX_Utilities_v2.DataFolder).Files
                    ' Neu data file do khong phai la file KHBS thi moi bat dau
                    If InStr(1, fle.Name, "KHBS") = 0 Then
                        ' Kiem tra xem co dung la data cua nam can sua doi khong
                        lLoc = InStr(1, fle.Name, "03_TNDN" & "_" & CStr(Right(strDay, 4)) & "_")
                        If lLoc <> 0 Then
                            lLoc = lLoc + Len("03_TNDN" & "_" & CStr(Right(strDay, 4)) & "_")
                            ' Lay ngay bat dau cua ky quyet toan TNDN
                            varFirstDay = Mid$(fle.Name, lLoc, 2) & "/" & Mid$(fle.Name, lLoc + 2, 2) & "/" & Mid$(fle.Name, lLoc + 4, 4)
                            ' Lay ngay ket thuc cua ky quyet toan TNDN
                            varLastDay = Mid$(fle.Name, lLoc + 9, 2) & "/" & Mid$(fle.Name, lLoc + 11, 2) & "/" & Mid$(fle.Name, lLoc + 13, 4)
                            Exit For
                        End If
                    End If
                Next
            End If
            
            ' hlnam Edit -End
                        
            TAX_Utilities_v2.Year = CStr(Right(strDay, 4))
            ' Gan lay ngay bat dau cua ky quet toan, neu khong co thi default la ngay '01/01'
            TAX_Utilities_v2.FirstDay = CStr(varFirstDay)
            ' Gan lay ngay ket thuc cua ky quet toan, neu khong co thi default la ngay '31/12'
            TAX_Utilities_v2.LastDay = CStr(varLastDay)
            
        Else
            TAX_Utilities_v2.Year = CStr(strDay)
        End If
        
    If Validtokhai(strId, strDay) = False Then
       DisplayMessage "0107", msOKOnly, miInformation
       Exit Sub
    End If
    
    
    With fpsNgaykhai
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        strDateKHBS = .Text
    End With
    
    If strDateKHBS <> vbNullString Then
        TAX_Utilities_v2.DateKHBS = Replace(strDateKHBS, "/", "")
    End If
    
        
        
         If GetAttribute(TAX_Utilities_v2.NodeMenu, "FinanceYear") = "1" Then
            strNgayTaiChinh = GetNgayBatDauNamTaiChinh
            If Not KiemTraNgayTaiChinh(strNgayTaiChinh, False) Then
                blnValidFinanceYear = False
            Else
                iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
                iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
                blnValidFinanceYear = True
            End If
        Else
                strNgayTaiChinh = "01/01"
                iNgayTaiChinh = 1
                iThangTaiChinh = 1
                blnValidFinanceYear = True
        End If
        
        
        strKHBS = "frmKHBS_BS"
        
        
        
        Unload Me
        Set frmTK = New frmInterfaces
        frmTK.Show

End Sub

Private Sub btnThoat_Click()
    Unload Me
    strHiddenFormName = vbNullString
End Sub

Private Sub btnThoat_LostFocus()
    blnOnExit = False
End Sub

Private Sub btnThoat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnOnExit = True
End Sub


Private Sub Form_Load()

    FormatGrid
    SetupData
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
    
    blnOnExit = False
    
    Me.Top = (frmSystem.ScaleHeight - Me.Height) / 2 - 250
    Me.Left = (frmSystem.ScaleWidth - Me.Width) / 2
    
    strInterfaceUnloadEventName = ""

End Sub
Sub FormatGrid()
    Dim i As Integer
    Dim vdtehientai As String
    Dim strarrdate() As String
    vdtehientai = format(Date, "dd/mm/yyyy")
    formatPrefix vdtehientai, strarrdate
    With fpsDkNgay
        .BackColor = mFormColor
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        .BackColor = vbWhite
        .CellType = CellTypePic
        '.TypePicMask = "99//9999"
        .TypePicMask = "9999"
    End With
    
    With fpsNgaykhai
        .BackColor = mFormColor
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        .BackColor = vbWhite
        .CellType = CellTypePic
        .TypePicMask = "99//99//9999"
        .Text = strarrdate(0) & "/" & strarrdate(1) & "/" & strarrdate(2)
        
    End With
    
    With fpsLoaiTK
        .BackColor = mFormColor
        'fpSpread4.BorderStyle = BorderStyleNone
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        .Col = .ColLetterToNumber(f4cboLoTCol)
        .Row = f4cboLoTRow
        .BackColor = vbWhite
    End With
End Sub
Sub SetupData()
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    
        With fpsDkNgay
            Dim vdtehientai As String
            vdtehientai = format(Date, "dd/mm/yyyy")
            Dim strarrdate() As String
            formatPrefix vdtehientai, strarrdate
            
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .CellType = CellTypePic
            .TypePicMask = "9999"
'            .TypePicMask = "99//9999"
'            If strarrdate(1) = 1 Then
'                strarrdate(1) = 12
'                strarrdate(2) = strarrdate(2) - 1
'            Else
'                strarrdate(1) = Right("0" & strarrdate(1) - 1, 2)
'            End If
'            .Text = strarrdate(1) & "/" & strarrdate(2)
            .Text = strarrdate(2) - 1
        End With
    'Lay du lieu cho cbo
        With fpsLoaiTK
            .Col = .ColLetterToNumber(f4cboLoTCol)
            .Row = f4cboLoTRow
            Dim xmlDocument As New MSXML.DOMDocument
            Dim xmlNode As MSXML.IXMLDOMNode
            Dim strDataFileName As String
            Dim i As Integer

            xmlDocument.Load TAX_Utilities_v2.GetAbsolutePath("ListKHBS.xml")
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
'            ReDim Preserve arrStrId(0)
'            arrStrId(0) = "00"
            
                
                
            For Each xmlNode In xmlNodeListMap
                Dim id As String
                Dim LoaiTk As String
                Dim Parentid As String
                ReDim Preserve arrStrId(i)
                arrStrId(i) = GetAttribute(xmlNode, "ID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                If arrStrId(i) = "01" Then
                    .TypeComboBoxIndex = 0
                Else
                    .TypeComboBoxIndex = -1
                End If
                .TypeComboBoxString = LoaiTk
                i = i + 1
            Next
            'lintSoTk = i
            .TypeComboBoxCurSel = 0
            lstryear = "1"
            lstrMonth = "0"
            lstrThreemonths = "0"
            Set xmlNode = Nothing
            Set xmlDocument = Nothing
        End With
End Sub
Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not blnOpenInterfaces Then
        frmTreeviewMenu.Show
    End If
    Set frmKhaibosung = Nothing
    lngRowFocus = 0
End Sub

'Private Sub fpsDkNgay_Click(ByVal Col As Long, ByVal Row As Long)
'With fpsDkNgay
'    MsgBox "Col: " & .Col & "            Row:" & .Row
'End With
'End Sub
Private Sub fpsDkNgay_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 1 Then
        If fpsDkNgay.ActiveRow = f1dteTuNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteTuNCol) Then
            fpsLoaiTK.SetFocus
            With fpsLoaiTK
                .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
            End With
        End If
        
    End If

End Sub
Private Sub fpsDkNgay_KeyPress(KeyAscii As Integer)
    'Chan ko cho nhap dau "."
    If KeyAscii = Asc(".") Then
        KeyAscii = 0
    End If
End Sub
Private Sub fpsDkNgay_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With fpsDkNgay
        Dim strarrdate() As String
        Dim strPrefix As String
        Dim vdtehientai As String
        Static blnOnDKienTraCuu_LeaveCell As Boolean          'Kiem tra su kien LeaveCell dang dc goi ???

        'Khi nguoi dung bam nut Thoat -> Bo qua ham nay.
        If blnOnExit Then Exit Sub

        '*******************************
        ' added
        'Date: 13/05/06
        'Su kien nay dang dc goi
        If blnOnDKienTraCuu_LeaveCell = True Then Exit Sub

        'Khoi tao gia tri dieu kien tra cuu
        blnOnDKienTraCuu_LeaveCell = True
        '*******************************

        'Khoi tao gia tri dieu kien tra cuu
        blnDKienTraCuu = False

        vdtehientai = format(Date, "dd/mm/yyyy")
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow

        If Trim(Replace(.Text, "/", "")) <> "" Then
            formatPrefix .Text, strarrdate
            'Bat dk thang
            If (Val(strarrdate(0)) > 12 Or Val(strarrdate(0)) <= 0) And lstrMonth = "1" Then
                .Text = ""
                DisplayMessage "0090", msOKOnly, miInformation
                blnOnDKienTraCuu_LeaveCell = False
                Exit Sub
            End If
            'Bat dk quy
            If (Val(strarrdate(0)) > 4 Or Val(strarrdate(0)) <= 0) And lstrThreemonths = "1" Then
                .Text = ""
                DisplayMessage "0091", msOKOnly, miInformation
                blnOnDKienTraCuu_LeaveCell = False
                Exit Sub
            End If
                'bat dk nam
                If lstrMonth = "0" And lstrThreemonths = "0" Then
                    Select Case Len(Trim(strarrdate(0)))
                    Case 1
                        strarrdate(0) = Year(vdtehientai)
                    Case 2
                        strarrdate(0) = "20" & Trim(strarrdate(0))
                    Case 3
                        strarrdate(0) = "2" & Trim(strarrdate(0))
                    Case 4
                        If Val(strarrdate(0)) < 2000 Then
                            strarrdate(0) = "2000"
                        End If
                End Select
            Else
                Select Case Len(Trim(strarrdate(1)))
                Case 1
                    strarrdate(1) = Year(vdtehientai)
                Case 2
                    strarrdate(1) = "20" & Trim(strarrdate(1))
                Case 3
                    strarrdate(1) = "2" & Trim(strarrdate(1))
                Case 4
                    If Val(strarrdate(1)) < 2000 Then
                          strarrdate(1) = "2000"
                    End If
                End Select
            End If
            'Hien thi lai kq
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            If lstrMonth = "0" And lstrThreemonths = "0" Then

                .SetText .Col, .Row, strarrdate(0)
            Else
                .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1)
            End If
        End If


    End With

    blnOnDKienTraCuu_LeaveCell = False
    blnDKienTraCuu = True
End Sub

Private Sub fpsLoaiTK_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    With fpsLoaiTK
        .Col = .ColLetterToNumber(f4cboLoTCol)
        .Row = f4cboLoTRow
'        If .TypeComboBoxCurSel = 0 Then
'            With fpsDkNgay
'                Dim vdtehientai As String
'                vdtehientai = Date
'                Dim strarrdate() As String
'                formatPrefix vdtehientai, strarrdate
'
'                .Col = .ColLetterToNumber(f1dteTuNCol)
'                .Row = f1dteTuNRow
'                .CellType = CellTypePic
'                .TypePicMask = "99//9999"
'                .Text = strarrdate(0) & "/" & strarrdate(2)
'                '.Text = strarrdate(2)
'
'            End With
'            Exit Sub
'        End If
        'xmlDocument.Load TAX_Utilities_v2.GetAbsolutePath("menu.xml")
        'Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            For Each xmlNode In xmlNodeListMenu
                Dim Parentid As String
                Parentid = GetAttribute(xmlNode, "PopID")
                '.TypeComboBoxIndex = 0
                If Parentid = "101" Then
                    If arrStrId(.TypeComboBoxCurSel) = GetAttribute(xmlNode, "ID") Then
                        lstryear = GetAttribute(xmlNode, "Year")
                        lstrMonth = GetAttribute(xmlNode, "Month")
                        lstrThreemonths = GetAttribute(xmlNode, "ThreeMonth")
                        Exit For
                    End If
                End If
            Next
        CreateDkKy
    End With
End Sub

Private Sub fpsLoaiTK_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 0 Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteTuNCol), f1dteTuNRow
        End With
    End If
    If KeyCode = vbKeyTab And Shift = 1 Then
        btnThoat.SetFocus
    End If
End Sub
Sub formatPrefix(strDate As String, strarrdate() As String)
    Dim lstrdate As Variant
    Dim i As Integer
    i = 0
    strarrdate = Split(strDate, "/")
    For Each lstrdate In strarrdate
        If Len(Trim(lstrdate)) < 2 Then
                strarrdate(i) = "0" & Trim(lstrdate)
         End If
         i = i + 1
    Next
End Sub
Sub CreateDkKy()
    Dim vdtehientai As String
    Dim strarrdate() As String
    vdtehientai = format(Date, "dd/mm/yyyy")
    formatPrefix vdtehientai, strarrdate
    
    With fpsDkNgay
        If lstrMonth = "0" And lstrThreemonths = "0" Then
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .CellType = CellTypePic
            .TypePicMask = "9999"
            .Text = strarrdate(2) - 1
            Exit Sub
        End If
        If lstrMonth = "1" Then
             .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            
            If strarrdate(1) = 1 Then
            strarrdate(1) = 12
            strarrdate(1) = strarrdate(1) - 1
            Else
            strarrdate(1) = Right("0" & strarrdate(1) - 1, 2)
            End If
            .Text = strarrdate(1) & "/" & strarrdate(2)
            Exit Sub
        End If
        If lstrThreemonths = "1" Then
            Dim strQuy As String
            If Val(strarrdate(1)) < 4 Then
                strQuy = "04"
                strarrdate(2) = strarrdate(2) - 1
            ElseIf Val(strarrdate(1)) >= 4 And Val(strarrdate(1)) < 7 Then
                 strQuy = "01"
            ElseIf Val(strarrdate(1)) >= 7 And Val(strarrdate(1)) < 10 Then
                strQuy = "02"
            Else
                strQuy = "03"
            End If
             .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .CellType = CellTypePic
            .TypePicMask = "99//9999"
            .Text = strQuy & "/" & strarrdate(2)
            Exit Sub
        End If
    End With
End Sub


'*********************************************************
' Description: Lay thong tin ve tat ca cac to khai theo
'              kieu kien tra cuu dua vao va ma to khai.
'       strId: Ma to khai
'       strPeriodFrom: Ky ke khai tu ky.
'       strPeriodTo: Ky ke khai den ky.
'       strPeriods(): Thong tin ve tat ca cac to khai, bao gom:
'           + Ma to khai
'           + Ten to khai
'           + Ky hieu luc
'           + Ky ke khai
'           + Ky ke khai tu ngay (neu co)
'           + Ky ke khai den ngay (neu co)
'           + Ten cac file du lieu
'           + Gia tri thue phai nop(neu co).
'           + Gia tri thue khau tru(neu co).
' Return: True neu ton tai to khai tuong ung voi ma,
'         False neu nguoc lai
'*********************************************************

Private Sub LoadXMLFileNames()
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As file
    Dim strId As String
    
    
 
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v2.DataFolder)).Files
        If Right$(fle.Name, 4) = ".xml" Then
            ReDim Preserve arrStrXMLFileNames(lngIndex)
            arrStrXMLFileNames(lngIndex) = Mid$(fle.Name, 1, Len(fle.Name) - 4)
            lngIndex = lngIndex + 1
        End If
    Next
End Sub



Private Function IsValidDate(ByVal strDate As String) As String
    Dim dDate As Variant
    
    On Error GoTo ErrHandle
    If Len(strDate) <> 8 Then Exit Function
    
    dDate = DateValue(Mid$(strDate, 1, 2) & "/" & Mid$(strDate, 3, 2) & "/" & Mid$(strDate, 5, 4))
    IsValidDate = Mid$(strDate, 1, 2) & "/" & Mid$(strDate, 3, 2) & "/" & Mid$(strDate, 5, 4)
    
    Exit Function
ErrHandle:
End Function


Private Function GetDataFileNames(arrStrDataFiles() As String, strPeriod As String) As String
    Dim intCtrl As Integer
    Dim strReturn As String
    
    For intCtrl = 0 To UBound(arrStrDataFiles)
        If intCtrl = 0 Then
            strReturn = arrStrDataFiles(intCtrl) & "_" & strPeriod
        Else
            strReturn = strReturn & "," & arrStrDataFiles(intCtrl) & "_" & strPeriod
        End If
    Next intCtrl
    
    GetDataFileNames = strReturn
End Function

Private Sub DeleteDataFiles(ByVal strFileNames As String)
    Dim arrStrDataFiles() As String
    Dim intCtrl As Integer
    Dim fso As New FileSystemObject
    
    arrStrDataFiles = Split(strFileNames, ",")
    
    For intCtrl = 0 To UBound(arrStrDataFiles)
        If fso.FileExists(GetAbsolutePath(TAX_Utilities_v2.DataFolder & arrStrDataFiles(intCtrl) & ".xml")) Then
            fso.DeleteFile GetAbsolutePath(TAX_Utilities_v2.DataFolder & arrStrDataFiles(intCtrl) & ".xml"), True
            '.DeleteRows i + 1, 1
        End If
    Next intCtrl
End Sub

Private Function Validtokhai(strId As String, strDay As String) As Boolean
   Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Dim xmlNodeMenu As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim fso As New FileSystemObject
    Dim fle As file
    
    Validtokhai = False
    
    
    xmlDocument.Load TAX_Utilities_v2.GetAbsolutePath("ListKHBS.xml")
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
    
    For Each xmlNodeMenu In xmlNodeListMap
        If GetAttribute(xmlNodeMenu, "ID") = strId Then
            Exit For
        End If
    Next
    
     If Not xmlNodeMenu Is Nothing Then
        'to khai 03_TNDN
        If strId = "03" Then
            strDataFileName = GetAttribute(xmlNodeMenu.childNodes.Item(0), "DataFile") & "_" & Trim(Replace(strDay, "/", "")) & "_" & Trim(Replace(varFirstDay, "/", "")) & "_" & Trim(Replace(varLastDay, "/", ""))
        Else
            strDataFileName = GetAttribute(xmlNodeMenu.childNodes.Item(0), "DataFile") & "_" & Trim(Replace(strDay, "/", ""))
        End If
        
     End If
  
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v2.DataFolder)).Files
        If fle.Name = strDataFileName & ".xml" Then
            Validtokhai = True
            Exit For
        End If
    Next
  Set xmlNodeMenu = Nothing
  Set xmlNodeListMap = Nothing
  Set xmlDocument = Nothing
    
End Function

Private Function checkValidate() As Boolean
    checkValidate = True
    
    With fpsDkNgay
        Dim strarrdate() As String
        Dim strPrefix As String
        Dim vdtehientai As String
        Dim vdKytinhthue As String
        
        vdtehientai = format(Date, "dd/mm/yyyy")
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        
        If Trim(Replace(.Text, "/", "")) <> "" Then
            formatPrefix .Text, strarrdate
            'Bat dk thang
            If (Val(strarrdate(0)) > 12 Or Val(strarrdate(0)) <= 0) And lstrMonth = "1" Then
                .Text = ""
                DisplayMessage "0090", msOKOnly, miInformation
                checkValidate = False
                .SetFocus
                Exit Function
            End If
            'Bat dk quy
            If (Val(strarrdate(0)) > 4 Or Val(strarrdate(0)) <= 0) And lstrThreemonths = "1" Then
                .Text = ""
                DisplayMessage "0091", msOKOnly, miInformation
                checkValidate = False
                .SetFocus
                Exit Function
            End If
                'bat dk nam
                If lstrMonth = "0" And lstrThreemonths = "0" Then
                    Select Case Len(Trim(strarrdate(0)))
                    Case 1
                        strarrdate(0) = Year(vdtehientai)
                    Case 2
                        strarrdate(0) = "20" & Trim(strarrdate(0))
                    Case 3
                        strarrdate(0) = "2" & Trim(strarrdate(0))
                    Case 4
                        If Val(strarrdate(0)) < 2000 Then
                            strarrdate(0) = "2000"
                        End If
                End Select
            Else
                Select Case Len(Trim(strarrdate(1)))
                Case 1
                    strarrdate(1) = Year(vdtehientai)
                Case 2
                    strarrdate(1) = "20" & Trim(strarrdate(1))
                Case 3
                    strarrdate(1) = "2" & Trim(strarrdate(1))
                Case 4
                    If Val(strarrdate(1)) < 2000 Then
                          strarrdate(1) = "2000"
                    End If
                End Select
            End If
            'Hien thi lai kq
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            If lstrMonth = "0" And lstrThreemonths = "0" Then
                .SetText .Col, .Row, strarrdate(0)
            Else
                .SetText .Col, .Row, strarrdate(0) & "/" & strarrdate(1)
            End If
        Else
            DisplayMessage "0062", msOKOnly, miCriticalError
            .SetFocus
            checkValidate = False
            Exit Function
        End If
        
        If Len(.Text) = 4 Then
            vdKytinhthue = "01/01/" & .Text
        Else
           If lstrThreemonths = "0" Then
                vdKytinhthue = "01/" & .Text
           Else
                Select Case Left(.Text, 2)
                    Case "01"
                        vdKytinhthue = "01/01/" & Right(.Text, 4)
                    Case "02"
                        vdKytinhthue = "01/04/" & Right(.Text, 4)
                    Case "03"
                        vdKytinhthue = "01/07/" & Right(.Text, 4)
                    Case "04"
                        vdKytinhthue = "01/10/" & Right(.Text, 4)
                 End Select
           End If
            
        End If
        
    End With
    
    With fpsNgaykhai
    .Col = .ColLetterToNumber(f1dteTuNCol)
    .Row = f1dteTuNRow
     If Len(.Text) > 0 Then
        Set objCvt = New DateUtils
        If IsNull(objCvt.ToDate(.Text, "DD/MM/YYYY")) Then
            DisplayMessage "0071", msOKOnly, miCriticalError
            .SetFocus
            checkValidate = False
            Exit Function
        Else
           If DateSerial(CInt(Mid$(.Text, 7, 4)), CInt(Mid$(.Text, 4, 2)), CInt(Mid$(.Text, 1, 2))) > DateSerial(CInt(Mid$(vdtehientai, 7, 4)), CInt(Mid$(vdtehientai, 4, 2)), CInt(Mid$(vdtehientai, 1, 2))) Or _
              DateSerial(CInt(Mid$(.Text, 7, 4)), CInt(Mid$(.Text, 4, 2)), CInt(Mid$(.Text, 1, 2))) < DateSerial(CInt(Mid$(vdKytinhthue, 7, 4)), CInt(Mid$(vdKytinhthue, 4, 2)), CInt(Mid$(vdKytinhthue, 1, 2))) Then
                DisplayMessage "0114", msOKOnly, miCriticalError
                .SetFocus
                checkValidate = False
               Exit Function
           End If
        End If
      Else
            DisplayMessage "0071", msOKOnly, miCriticalError
            .SetFocus
            checkValidate = False
            Exit Function
    End If
    
   End With
    
End Function


