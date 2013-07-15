VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTraCuuiHTKK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   6960
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   291
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CheckBox chkNhanTDong 
      Caption         =   "NhËn tù ®éng"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   6840
      Width           =   1425
   End
   Begin VB.Frame Frame5 
      Caption         =   "M· sè thuÕ"
      Height          =   570
      Left            =   210
      TabIndex        =   14
      Top             =   1560
      Width           =   5445
      Begin VB.TextBox txtMST 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         MaxLength       =   15
         TabIndex        =   15
         Top             =   170
         Width           =   2445
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tr¹ng th¸i tê khai"
      Height          =   975
      Left            =   3240
      TabIndex        =   10
      Top             =   2160
      Width           =   7860
      Begin VB.CheckBox chkCBXL 
         Caption         =   "CÇn c¸n bé thuÕ xö lý"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton optIHTKK 
         Caption         =   "Ch­a nhËn sang vïng trung gian NTK"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optRecv 
         Caption         =   "NhËn vµo vïng trung gian NTK"
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   240
         Width           =   2745
      End
   End
   Begin VB.CommandButton btnThoat 
      Caption         =   "&Tho¸t"
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton btnTraCuu 
      Caption         =   "Tra &cøu"
      Height          =   375
      Index           =   0
      Left            =   7140
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin FPUSpreadADO.fpSpread fpsKetQua 
      Height          =   3060
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   10725
      _Version        =   458752
      _ExtentX        =   18918
      _ExtentY        =   5397
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
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
      GridColor       =   16777215
      MaxCols         =   12
      MaxRows         =   13
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   14
      SpreadDesigner  =   "frmTracuuiHTKK.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän ngµy nép vµ kú lËp bé"
      Height          =   1410
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   5320
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _Version        =   458752
         _ExtentX        =   8070
         _ExtentY        =   1508
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
         MaxCols         =   7
         MaxRows         =   5
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "frmTracuuiHTKK.frx":0752
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   10995
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chän lo¹i tê khai"
      Height          =   760
      Left            =   210
      TabIndex        =   7
      Top             =   720
      Width           =   5445
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   435
         Left            =   360
         TabIndex        =   0
         Top             =   220
         Width           =   4725
         _Version        =   458752
         _ExtentX        =   8334
         _ExtentY        =   767
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
         SpreadDesigner  =   "frmTracuuiHTKK.frx":0CF4
         UserResize      =   1
      End
   End
   Begin VB.CommandButton cmdNhanTkhai 
      Caption         =   "NhËn tê khai"
      Height          =   375
      Left            =   8430
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
   End
   Begin MSForms.Label lblDangXuLy 
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      ForeColor       =   255
      Caption         =   "§ang xö lý ..."
      Size            =   "2355;450"
      BorderColor     =   0
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   210
      TabIndex        =   9
      Top             =   600
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
      Left            =   0
      Top             =   0
      Width           =   10755
   End
   Begin MSForms.Label LblSoBG 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   6550
      Width           =   3015
      BackColor       =   -2147483648
      VariousPropertyBits=   8388627
      Size            =   "5318;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTraCuuiHTKK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const f1dteTuNRow = 2
Private Const f1dteTuNCol = "C"
Private Const f1dteDeNRow = 2
Private Const f1dteDeNCol = "F"
Private Const f4cboLoTRow = 2
Private Const f4cboLoTCol = "C"

' Ky lap bo
Private Const f1KyLBRow = 4
Private Const f1KyLBCol = "C"


Private Const mFormColor = -2147483633
Private Const lminDate = "01/01/1900"
Private Const lmaxDate = "31/12/3000"
Private Const lmaxSoTk = 1000

Private lngRowFocus As Long
Private DteTuN As Date
Private DteDeN As Date
Private kyLapBo As Variant
Private lmaTK As Long
Private lSoBG As Long
Private larrId(lmaxSoTk) As Long
Private lerror As Boolean
Private strDaNhan As String

Private Sub btnThoat_Click(Index As Integer)
    Unload Me
End Sub
Private Sub btnTraCuu_Click(Index As Integer)
    traCuuToKhai
End Sub

Public Sub traCuuToKhai()
    Dim rsReturn As New ADODB.Recordset
    Dim rsTotalReturn As New ADODB.Recordset
    
    
    Dim strSQL As String
    Dim strTotalSQL As String
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    
    Dim rsCLob As New ADODB.Recordset
    Dim strTemp As String
    Dim iCountClob As Integer
    
    'Xoa bo ket qua cu tren Grid
    fpsKetQua.ClearRange 1, 1, fpsKetQua.MaxCols, lSoBG, True
    If Not KiemTraDKngay Then
         Exit Sub
    End If
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    
    If chkCBXL.Value = 1 Then
        'neu tra cuu theo can bo xu ly thi danh dau lai de thuc hien ghi du lieu
        TAX_Utilities_Srv_New.isCanBoXuLyGhiTK = True
    Else
        TAX_Utilities_Srv_New.isCanBoXuLyGhiTK = False
    End If
    
    'Lay cau lenh truy van
    xmlSQL.Load App.path & "\SQL.xml"
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraCuuiHTKK")
    strSQL = Replace(strSQL, "nhohon=", "<=")
    strSQL = Replace(strSQL, "ma_tkhai", "" & changeLoaiToKhai(lmaTK) & "")
    strSQL = Replace(strSQL, "strDa_Nhan", "" & strDaNhan & "")
    strSQL = Replace(strSQL, "ngay_nop_dau", "To_date('" & format(DteTuN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    strSQL = Replace(strSQL, "ngay_nop_cuoi", "To_date('" & format(DteDeN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    
'    ' Dem tong so ban ghi
'    strTotalSQL = GetAttribute(xmlSQL.childNodes(1), "SqlCountiHTKK")
'    strTotalSQL = Replace(strTotalSQL, "nhohon=", "<=")
'    strTotalSQL = Replace(strTotalSQL, "ma_tkhai", "" & changeLoaiToKhai(lmaTK) & "")
'    strTotalSQL = Replace(strTotalSQL, "strDa_Nhan", "" & strDaNhan & "")
'    strTotalSQL = Replace(strTotalSQL, "ngay_nop_dau", "To_date('" & format(DteTuN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
'    strTotalSQL = Replace(strTotalSQL, "ngay_nop_cuoi", "To_date('" & format(DteDeN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    
    'Khoi tao bien
    lCol = 2
    lRow = 1
    lSoBG = 0
    
'    ' dem tong so ban ghi
'    Set rsTotalReturn = clsDAO.Execute(strTotalSQL)
'    If Not rsTotalReturn Is Nothing Then
'        ' Hien thi progress
'        ProgressBar1.Visible = True
'        lblDangXuLy.Visible = True
'        ProgressBar1.Value = 0
'        ' end
'        ProgressBar1.max = rsTotalReturn.Fields(0).Value
'    Else
'        ProgressBar1.Visible = False
'        lblDangXuLy.Visible = False
'    End If
    
    'Thuc hien cau lenh sql
    Set rsReturn = clsDAO.Execute(strSQL)
    'Hien thi du lieu len grid
    If rsReturn.Fields.Count > 0 Then
        Do While Not rsReturn.EOF
            fpsKetQua.MaxRows = lRow + 1
            fpsKetQua.InsertRows lRow, 1
            fpsKetQua.SetText 1, lRow, lSoBG + 1
'            ' progress
'            ProgressBar1.Value = lSoBG + 1
'            'end
            For lIndex = 1 To rsReturn.Fields.Count
                If Not (IsNull(rsReturn.Fields(lIndex - 1).Value)) Or UCase(Trim(rsReturn.Fields(lIndex - 1).Name)) = "DLIEU_MVACH" Then
                    ' Cho phep tich vao cac check_box
                    fpsKetQua.Col = 12
                    fpsKetQua.Row = lRow
                    fpsKetQua.CellType = CellTypeCheckBox
                    fpsKetQua.TypeHAlign = TypeHAlignCenter
                    fpsKetQua.Lock = False
                    If UCase(Trim(rsReturn.Fields(lIndex - 1).Name)) = "DLIEU_MVACH" Then
'                        strTemp = vbNullString
'                        strSQL = vbNullString
'                        For iCountClob = 0 To 1000
'                            'Debug.Print "select DBMS_LOB.substr(DLIEU_MVACH,4000,((4000 * " & iCountClob & ")+1)) from rcv_ihtkk_mvach where ID= " & rsReturn.Fields("ID").Value
'                            strSQL = "select DBMS_LOB.substr(DLIEU_MVACH,4000,((4000 * " & iCountClob & ")+1)) as DLIEU_MVACH from rcv_ihtkk_mvach where ID =" & rsReturn.Fields("ID").Value
'                            Set rsCLob = clsDAO.Execute(strSQL)
'                            If rsCLob.Fields.Count > 0 Then
'                                If Not IsNull(rsCLob.Fields("DLIEU_MVACH").Value) Then
'                                    strTemp = strTemp & Trim(rsCLob.Fields("DLIEU_MVACH").Value)
'                                Else
'                                    Exit For
'                                End If
'                            Else
'                                Exit For
'                            End If
'                        Next
                        strTemp = ""
                        fpsKetQua.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(strTemp, TCVN, UNICODE)
                    Else
                        fpsKetQua.SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
                    End If
                    lCol = lCol + 1
                    fpsKetQua.RowHeight(lRow) = fpsKetQua.MaxTextRowHeight(lRow)
                End If
            Next lIndex
            lCol = 2
            lRow = lRow + 1
            
            rsReturn.MoveNext
            lSoBG = lSoBG + 1
        Loop
    Else
        lSoBG = 0
        DisplayMessage "0072", msOKOnly
    End If
    LblSoBG.Visible = True
    LblSoBG.TextAlign = fmTextAlignLeft
    LblSoBG.caption = "Sè tê khai t×m thÊy: " & lSoBG
    
    'An thanh progress di
    ProgressBar1.Visible = False
    lblDangXuLy.Visible = False

        
    '
    If Trim(strDaNhan) = "= 'Y'" Then
        cmdNhanTkhai.Enabled = False
    Else
        cmdNhanTkhai.Enabled = True
    End If
    'Chinh sua lai Grid ket cho dep
    With fpsKetQua
        .MaxRows = lSoBG
'        If lSoBG <= 15 Then
'            Dim i As Integer
'            For i = 1 To 15 - lRow
'               ' .MaxRows = lSoBG + i + 1
'                .InsertRows lSoBG + i, 1
'            Next
'        Else
'            .RowHeight(lRow) = 0
'        End If
        
        If lngRowFocus = 0 Then
            If lSoBG <> 0 Then
                lngRowFocus = SetRowFocus(1, 1, True)
            End If
        ElseIf lngRowFocus > .MaxRows Then
            lngRowFocus = SetRowFocus(1, .MaxRows, True)
        Else
            lngRowFocus = SetRowFocus(1, lngRowFocus, True)
        End If
        .SetFocus
    End With
End Sub


Private Function changeLoaiToKhai(ByVal strLoaiMaToKhai As String) As String
    If strLoaiMaToKhai = "102" Then changeLoaiToKhai = " IN ('01', '02', '04', '07','71','72')"
    If strLoaiMaToKhai = "103" Then changeLoaiToKhai = " IN ('03', '11', '12', '14','73')"
    If strLoaiMaToKhai = "104" Then changeLoaiToKhai = " IN ('46', '47', '48', '49', '15', '16', '53', '37', '50', '51', '54', '38', '39', '40', '36', '17','42','43','41','59','74','75')"
    If strLoaiMaToKhai = "105" Then changeLoaiToKhai = " IN ('06', '08', '09','77')"
    If strLoaiMaToKhai = "106" Then changeLoaiToKhai = " IN ('05')"
    If strLoaiMaToKhai = "108" Then changeLoaiToKhai = " IN ('18','55','56','57','58','69') "
    If strLoaiMaToKhai = "109" Then changeLoaiToKhai = " IN ('19','24','25','26','27') "
    If strLoaiMaToKhai = "110" Then changeLoaiToKhai = " IN ('20','28','29','30','31') "
    If strLoaiMaToKhai = "111" Then changeLoaiToKhai = " IN ('21','32','33','34','35') "
    If strLoaiMaToKhai = "112" Then changeLoaiToKhai = " IN ('64','65','66','67','68') "
    If strLoaiMaToKhai = "101" Then changeLoaiToKhai = " IN ('70','80','81','82') "
    If strLoaiMaToKhai = "113" Then changeLoaiToKhai = " IN ('86','87','89') "
    
    If strLoaiMaToKhai = "0" Then changeLoaiToKhai = " LIKE '%' "
    If Trim(txtMST.Text) <> vbNullString Then
        changeLoaiToKhai = changeLoaiToKhai & ") And (TIN like '%" & UCase(Trim(txtMST.Text)) & "%'"
    End If
End Function

Private Function validToKhaiiHTKK(ByVal maSoThue As String, ByVal maToKhai As String, ByVal KyKeKhai As String, ByVal lanNop As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    validToKhaiiHTKK = False
    
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    ' Lay tat ca cac to khai trong cung mot ky tinh thue, cua mot loai to khai, cua mot nguoi nop thue va co so lan nop den thoi diem kiem tra
    strSQL = "select ID from RCV_IHTKK_MVACH where TIN = '" & maSoThue & "' and MA_TKHAI = '" & maToKhai & "' and KY_KKHAI = '" & KyKeKhai & "' and DA_NHAN IN ('N')and LAN_NOP < " & lanNop
    Set rs = clsDAO.Execute(strSQL)
    
    'Neu con to khac lan nop truoc do chua duoc chuyen vao QLT_NTK thi tra lai ket qua la true
    If rs.Fields.Count > 0 Then
        validToKhaiiHTKK = True
    End If
End Function

Private Sub chkCBXL_Click()
    If chkCBXL.Value = 1 Then
        strDaNhan = " = 'W'"
    Else
        strDaNhan = " = 'N'"
    End If
End Sub

Private Sub chkNhanTDong_Click()
    If optIHTKK.Value And chkNhanTDong.Value Then
        chkCBXL.Enabled = True
    Else
        chkCBXL.Value = 0
        chkCBXL.Enabled = False
    End If
End Sub

Private Sub cmdNhanTkhai_Click()
    Dim varBarcodeiHTKK As Variant
    Dim varBarcodeScan As String
    Dim varTkhaiIdIhtkk As Variant
    
    Dim varNgayNop As Variant
    
    Dim varTkhaiMaSoThue As Variant
    Dim varTkhaiMaToKhai As Variant
    Dim varTkhaiKyKeKhai As Variant
    Dim varTkhaiLanNop As Variant
    Dim msgMessValidiHTKK As Boolean
    msgMessValidiHTKK = False
    
    Dim rsCLob As New ADODB.Recordset
    Dim rsReturn As New ADODB.Recordset
    Dim strTemp As String
    Dim iCountClob As Integer
    Dim strSQL As String
    
    Dim kyLB As Variant
    Dim arrKyLB() As String
    
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    
   If chkNhanTDong.Value = 1 Then
   ' Kiem tra ky LB
        With fpsDkNgay
             ' Ky lap bo
            .Col = .ColLetterToNumber(f1KyLBCol)
            .Row = f1KyLBRow
            .GetText .Col, .Row, kyLB
            If Trim(kyLB) = "" Or Trim(kyLB) = "../...." Then
                DisplayMessage "0128", msOKOnly, miCriticalError
                Exit Sub
            End If
            ' Ky lB khong dc lon hon ky hien tai
            If DateSerial(CInt(Right$(kyLB, 4)), CInt(Left$(kyLB, 2)), 1) > Date Then
                DisplayMessage "0129", msOKOnly, miCriticalError
                Exit Sub
            End If
            ' Set ky LB
            frmInterfaces.kyLapBo_IHTKK = CStr(kyLB)
            TAX_Utilities_Srv_New.HthucNopIHTKK = True
            TAX_Utilities_Srv_New.KyLBIHTKK = kyLB
            TAX_Utilities_Srv_New.NhanTuDongIHTKK = True
        End With
   
   End If
   
    
    ' Nhan phat mot, voi cac to khai chua chuyen sang QLT_NTK
    If (chkNhanTDong.Value = 0 And optRecv.Value = False) Then
        frmInterfaces.isNhanTuDong = False
        
        With fpsKetQua
            If lngRowFocus = 0 Then Exit Sub
            ' Lay thong tin ve chuoi ma vach
            .GetText 5, lngRowFocus, varBarcodeiHTKK        ' Get du lieu ma vach
            ' Lay thong tin ve Ma so thue
            .GetText 2, lngRowFocus, varTkhaiMaSoThue         ' Get du lieu ma so thue
            ' Lay thong tin ve Ma to khai
            .GetText 10, lngRowFocus, varTkhaiMaToKhai         ' Get du lieu ma to khai
            ' Lay thong tin ve ky ke khai
            .GetText 4, lngRowFocus, varTkhaiKyKeKhai         ' Get du lieu ky ke khai
            ' Lay thong tin ve lan nop
            .GetText 11, lngRowFocus, varTkhaiLanNop         ' Get du lieu lan nop
            
            ' Lay ve ID cua to khai
            .GetText 8, lngRowFocus, varTkhaiIdIhtkk        ' Get ID cua iHTKK
            'Lay ve du lieu cua ma vach
                strTemp = vbNullString
                strSQL = vbNullString
                For iCountClob = 0 To 1000
                    strSQL = "select DBMS_LOB.substr(DLIEU_MVACH,4000,((4000 * " & iCountClob & ")+1)) as DLIEU_MVACH from rcv_ihtkk_mvach where ID =" & CDbl(Trim(varTkhaiIdIhtkk))
                    Set rsCLob = clsDAO.Execute(strSQL)
                    If rsCLob.Fields.Count > 0 Then
                        If Not IsNull(rsCLob.Fields("DLIEU_MVACH").Value) Then
                            strTemp = strTemp & Trim(rsCLob.Fields("DLIEU_MVACH").Value)
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            ' end
            varBarcodeiHTKK = strTemp
            
            If Trim$(varBarcodeiHTKK) <> vbNullString Then
                ' Kiem tra xem co to nao cung ky ke khai, va duoc gui truoc nhung chua duoc nhan vao hay chua?
                If validToKhaiiHTKK(varTkhaiMaSoThue, varTkhaiMaToKhai, varTkhaiKyKeKhai, CInt(varTkhaiLanNop)) = True Then
                    msgMessValidiHTKK = True
                    .Row = lngRowFocus
                    .ForeColor = vbBlue
                Else
                    ' Chuyen cac tham so va sang man hinh nhan to khai
                    varBarcodeScan = Trim(varBarcodeiHTKK)
                    frmInterfaces.SetReceiveByBarcode True
                    frmInterfaces.Show
                    frmInterfaces.isIHTKK = True
                    ' Lay thong tin ve ID cua iHTKK
                    frmInterfaces.tkhai_ID_IHTKK = CDbl(varTkhaiIdIhtkk)
                    ' Lay thong tin ngay nop
                    .GetText 7, lngRowFocus, varNgayNop
                    frmInterfaces.ngay_nop_IHTKK = CStr(varNgayNop)
                    frmInterfaces.Barcode_Scaned varBarcodeScan
                End If
            End If
        End With
    ElseIf optRecv.Value = True Then ' Doi voi to khai da chuyen sang QLT_NTK, chi cho xem khong cho ghi lai
        With fpsKetQua
            If lngRowFocus = 0 Then Exit Sub
            ' Lay thong tin ve chuoi ma vach
            .GetText 5, lngRowFocus, varBarcodeiHTKK        ' Get du lieu ma vach
            
            ' Lay ve ID cua to khai
            .GetText 8, lngRowFocus, varTkhaiIdIhtkk        ' Get ID cua iHTKK
            'Lay ve du lieu cua ma vach
                strTemp = vbNullString
                strSQL = vbNullString
                For iCountClob = 0 To 1000
                    strSQL = "select DBMS_LOB.substr(DLIEU_MVACH,4000,((4000 * " & iCountClob & ")+1)) as DLIEU_MVACH from rcv_ihtkk_mvach where ID =" & CDbl(Trim(varTkhaiIdIhtkk))
                    Set rsCLob = clsDAO.Execute(strSQL)
                    If rsCLob.Fields.Count > 0 Then
                        If Not IsNull(rsCLob.Fields("DLIEU_MVACH").Value) Then
                            strTemp = strTemp & Trim(rsCLob.Fields("DLIEU_MVACH").Value)
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            ' end
            varBarcodeiHTKK = strTemp
            
            If Trim$(varBarcodeiHTKK) <> vbNullString Then
                ' Chuyen cac tham so va sang man hinh nhan to khai
                varBarcodeScan = Trim(varBarcodeiHTKK)
                frmInterfaces.SetReceiveByBarcode True
                frmInterfaces.Show
                frmInterfaces.isIHTKK = True
                ' Lay thong tin ve ID cua iHTKK
                .GetText 8, lngRowFocus, varTkhaiIdIhtkk        ' Get ID cua iHTKK
                frmInterfaces.tkhai_ID_IHTKK = CDbl(varTkhaiIdIhtkk)
                ' Lay thong tin ngay nop
                .GetText 7, lngRowFocus, varNgayNop
                frmInterfaces.ngay_nop_IHTKK = CStr(varNgayNop)
                frmInterfaces.Barcode_Scaned varBarcodeScan
                frmInterfaces.cmdSave.Enabled = False
                frmInterfaces.cmdClear.Enabled = False
                frmInterfaces.cmdViewNow.Enabled = False
            End If
        End With
    ElseIf chkNhanTDong.Value = 1 And optIHTKK.Value = True Then  ' Nhan luon tat ca
        frmInterfaces.isNhanTuDong = True
        
        If lngRowFocus = 0 Then Exit Sub
        Dim i As Long
        Dim strValueChk As Variant
        Dim blValueChk As Boolean
        
        With fpsKetQua
            '.EventEnabled(EventAllEvents) = False
            blValueChk = False ' Dat kiem tra xem co check vao cac to khai duoc chon ko, default la ko chon
            For i = 1 To .MaxRows
                .GetText 12, i, strValueChk
                If strValueChk = "1" Then
                    blValueChk = True ' Neu co bat ke to nao duoc chon vao check thi dat la true
                    Exit For
                End If
            Next
            If blValueChk = False Then ' Neu ko chon to nao ma nhan tu dong thi nhan tat
                For i = 1 To .MaxRows
                    ' Lay thong tin ve chuoi ma vach
                    .GetText 5, i, varBarcodeiHTKK         ' Get du lieu ma vach
                    ' Lay thong tin ve Ma so thue
                    .GetText 2, i, varTkhaiMaSoThue         ' Get du lieu ma so thue
                    ' Lay thong tin ve Ma to khai
                    .GetText 10, i, varTkhaiMaToKhai         ' Get du lieu ma to khai
                    ' Lay thong tin ve ky ke khai
                    .GetText 4, i, varTkhaiKyKeKhai         ' Get du lieu ky ke khai
                    ' Lay thong tin ve lan nop
                    .GetText 11, i, varTkhaiLanNop         ' Get du lieu lan nop
            
                    
                    ' Lay ve ID cua to khai
                    .GetText 8, i, varTkhaiIdIhtkk        ' Get ID cua iHTKK
                    'Lay ve du lieu cua ma vach
                        strTemp = vbNullString
                        strSQL = vbNullString
                        For iCountClob = 0 To 1000
                            strSQL = "select DBMS_LOB.substr(DLIEU_MVACH,4000,((4000 * " & iCountClob & ")+1)) as DLIEU_MVACH from rcv_ihtkk_mvach where ID =" & CDbl(Trim(varTkhaiIdIhtkk))
                            Set rsCLob = clsDAO.Execute(strSQL)
                            If rsCLob.Fields.Count > 0 Then
                                If Not IsNull(rsCLob.Fields("DLIEU_MVACH").Value) Then
                                    strTemp = strTemp & Trim(rsCLob.Fields("DLIEU_MVACH").Value)
                                Else
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    ' end
                    varBarcodeiHTKK = strTemp
                    
                    
                    If Trim$(varBarcodeiHTKK) <> vbNullString Then
                        ' Kiem tra xem co to nao cung ky ke khai, va duoc gui truoc nhung chua duoc nhan vao hay chua?
                        If validToKhaiiHTKK(varTkhaiMaSoThue, varTkhaiMaToKhai, varTkhaiKyKeKhai, CInt(varTkhaiLanNop)) = True Then
                            msgMessValidiHTKK = True
                            .Row = i
                            .ForeColor = vbBlue
                        Else
                            ' Chuyen cac tham so va sang man hinh nhan to khai
                            varBarcodeScan = Trim(varBarcodeiHTKK)
                            frmInterfaces.SetReceiveByBarcode True
                            frmInterfaces.Show
                            frmInterfaces.isIHTKK = True
                            ' set gia tri true de khong load lai du lieu sau khi nhan tu dong
                            frmInterfaces.isCheckList = True
                            ' Lay thong tin ve ID cua iHTKK
                            .GetText 8, i, varTkhaiIdIhtkk        ' Get ID cua iHTKK
                            frmInterfaces.tkhai_ID_IHTKK = CDbl(varTkhaiIdIhtkk)
                            ' Lay thong tin ngay nop
                            .GetText 7, i, varNgayNop
                            frmInterfaces.ngay_nop_IHTKK = CStr(varNgayNop)
                            frmInterfaces.lanNopTKIHTKK = varTkhaiLanNop ' lay so lan nop to khai
                            frmInterfaces.Barcode_Scaned varBarcodeScan
                            frmInterfaces.cmdSave_Click
                        End If
                    End If
                Next
            Else ' Neu chon to nao thi lay to khai do chuyen vao vung trung gian QLT_NTK
                For i = 1 To .MaxRows
                    ' Lay danh sach cac check, neu check vao to khai nao thi moi chuyen vao to khai do
                    .GetText 12, i, strValueChk
                    If strValueChk = "1" Then
                        ' Lay thong tin ve chuoi ma vach
                        .GetText 5, i, varBarcodeiHTKK         ' Get du lieu ma vach
                        ' Lay thong tin ve Ma so thue
                        .GetText 2, i, varTkhaiMaSoThue         ' Get du lieu ma so thue
                        ' Lay thong tin ve Ma to khai
                        .GetText 10, i, varTkhaiMaToKhai         ' Get du lieu ma to khai
                        ' Lay thong tin ve ky ke khai
                        .GetText 4, i, varTkhaiKyKeKhai         ' Get du lieu ky ke khai
                        ' Lay thong tin ve lan nop
                        .GetText 11, i, varTkhaiLanNop         ' Get du lieu lan nop
                        
                        
                         ' Lay ve ID cua to khai
                        .GetText 8, i, varTkhaiIdIhtkk        ' Get ID cua iHTKK
                        'Lay ve du lieu cua ma vach
                            strTemp = vbNullString
                            strSQL = vbNullString
                            For iCountClob = 0 To 1000
                                strSQL = "select DBMS_LOB.substr(DLIEU_MVACH,4000,((4000 * " & iCountClob & ")+1)) as DLIEU_MVACH from rcv_ihtkk_mvach where ID =" & CDbl(Trim(varTkhaiIdIhtkk))
                                Set rsCLob = clsDAO.Execute(strSQL)
                                If rsCLob.Fields.Count > 0 Then
                                    If Not IsNull(rsCLob.Fields("DLIEU_MVACH").Value) Then
                                        strTemp = strTemp & Trim(rsCLob.Fields("DLIEU_MVACH").Value)
                                    Else
                                        Exit For
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ' end
                        varBarcodeiHTKK = strTemp
                       
                        
                        If Trim$(varBarcodeiHTKK) <> vbNullString Then
                            If validToKhaiiHTKK(varTkhaiMaSoThue, varTkhaiMaToKhai, varTkhaiKyKeKhai, CInt(varTkhaiLanNop)) = True Then
                                msgMessValidiHTKK = True
                                .Row = i
                                .ForeColor = vbBlue
                            Else
                                ' Chuyen cac tham so va sang man hinh nhan to khai
                                varBarcodeScan = Trim(varBarcodeiHTKK)
                                frmInterfaces.SetReceiveByBarcode True
                                frmInterfaces.Show
                                frmInterfaces.isIHTKK = True
                                frmInterfaces.isCheckList = True
                                ' Lay thong tin ve ID cua iHTKK
                                .GetText 8, i, varTkhaiIdIhtkk        ' Get ID cua iHTKK
                                frmInterfaces.tkhai_ID_IHTKK = CDbl(varTkhaiIdIhtkk)
                                ' Lay thong tin ngay nop
                                .GetText 7, i, varNgayNop
                                frmInterfaces.ngay_nop_IHTKK = CStr(varNgayNop)
                                frmInterfaces.lanNopTKIHTKK = varTkhaiLanNop ' lay so lan nop to khai
                                frmInterfaces.Barcode_Scaned varBarcodeScan
                                frmInterfaces.cmdSave_Click
                            End If
                        End If
                    End If
                Next
            End If
            traCuuToKhai
            
            '.EventEnabled(EventAllEvents) = True
        End With
        ' Neu con to khai bi loi thi hien ra thong bao
    End If
    'reset bien dung de check khi nhan iHTKK
    TAX_Utilities_Srv_New.HthucNopIHTKK = False
    TAX_Utilities_Srv_New.KyLBIHTKK = ""
    TAX_Utilities_Srv_New.NhanTuDongIHTKK = False
    If msgMessValidiHTKK Then
        With fpsKetQua
            For i = 1 To .MaxRows
                ' Lay thong tin ve chuoi ma vach
                .GetText 5, i, varBarcodeiHTKK         ' Get du lieu ma vach
                ' Lay thong tin ve Ma so thue
                .GetText 2, i, varTkhaiMaSoThue         ' Get du lieu ma so thue
                ' Lay thong tin ve Ma to khai
                .GetText 10, i, varTkhaiMaToKhai         ' Get du lieu ma to khai
                ' Lay thong tin ve ky ke khai
                .GetText 4, i, varTkhaiKyKeKhai         ' Get du lieu ky ke khai
                ' Lay thong tin ve lan nop
                .GetText 11, i, varTkhaiLanNop         ' Get du lieu lan nop
                If Trim$(varBarcodeiHTKK) <> vbNullString Then
                    If validToKhaiiHTKK(varTkhaiMaSoThue, varTkhaiMaToKhai, varTkhaiKyKeKhai, CInt(varTkhaiLanNop)) = True Then
                        msgMessValidiHTKK = True
                        .Row = i
                        .ForeColor = vbBlue
                    End If
                End If
            Next
        End With
        DisplayMessage "0104", msOKOnly, miInformation
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    SetControlCaption Me, "frmTraCuuiHTKK"
    FormatGrid
    SetupData
    'lerror = False
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
End Sub
Sub FormatGrid()
    With fpsDkNgay
        .BackColor = mFormColor
      'fpSpread1.BorderStyle = BorderStyleNone
        .ColHeadersShow = False
        .RowHeadersShow = False
        .EditModePermanent = True
        .EditModeReplace = True
        '.BorderStyle = BorderStyleNone
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        .BackColor = vbWhite
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        .BackColor = vbWhite
        
        ' Ky LB
        .Col = .ColLetterToNumber(f1KyLBCol)
        .Row = f1KyLBRow
        .BackColor = vbWhite
    End With
    With fpsKetQua
        .EditModePermanent = True
        .EditModeReplace = True
        .ColHeadersShow = True
        .RowHeadersShow = False
        .AllowColMove = False
        .MaxRows = 17
        .TextTip = TextTipFloating
        .CursorType = CursorTypeLockedCell
        .CursorStyle = CursorStyleArrow
        .SetTextTipAppearance "Tahoma", 8, False, False, RGB(255, 255, 235), &H0
        
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
    LblSoBG.Visible = False
    
End Sub
Sub SetupData()
        With fpsDkNgay
            Dim vdtehientai As Variant
            Dim vKyLB As Variant
            
            vdtehientai = Date
            
            vKyLB = format(vdtehientai, "mm/yyyy")
            
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .SetText .Col, .Row, vdtehientai
            
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .SetText .Col, .Row, vdtehientai
            
            ' Ky LB
            .Col = .ColLetterToNumber(f1KyLBCol)
            .Row = f1KyLBRow
            .SetText .Col, .Row, vKyLB
        End With
    'Lay du lieu cho cbo
        With fpsLoaiTK
            .Col = .ColLetterToNumber(f4cboLoTCol)
            .Row = f4cboLoTRow
            Dim xmlDocument As New MSXML.DOMDocument
            Dim xmlNode As MSXML.IXMLDOMNode
            Dim strDataFileName As String
            Dim i As Integer
            xmlDocument.Load TAX_Utilities_Srv_New.GetAbsolutePath("menu.xml")
            Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            i = 0
            larrId(0) = 0
            For Each xmlNode In xmlNodeListMenu
                Dim LoaiTk As String
                Dim Parentid As String
                Parentid = GetAttribute(xmlNode, "ParentID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                '.TypeComboBoxIndex = 0
                If Parentid = "101" Then
                    i = i + 1
                    .TypeComboBoxIndex = -1
                    .TypeComboBoxString = LoaiTk
                    larrId(i) = Val(GetAttribute(xmlNode, "ID"))
                End If
            Next
            .TypeComboBoxCurSel = 0
            Set xmlNode = Nothing
            Set xmlDocument = Nothing
        End With
        'strDaNhan = " IS NULL "
        strDaNhan = " = 'N'"
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    frmTraCuuiHTKK.Top = (frmSystem.Height - frmTraCuuiHTKK.Height) / 2
    frmTraCuuiHTKK.Left = (frmSystem.Width - frmTraCuuiHTKK.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lngRowFocus = 0
End Sub

Private Sub fpsDkNgay_GotFocus()
    'btnTraCuu(0).Default = True
End Sub
Private Sub fpsDkNgay_Keydown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyTab And Shift = 0 Then 'And Not lerror
'    If fpsDkNgay.ActiveRow = f1dteDeNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteDeNCol) Then
'            fpsKetQua.SetFocus
'    End If
'End If

If KeyCode = vbKeyTab And Shift = 0 Then 'And Not lerror
    If fpsDkNgay.ActiveRow = f1KyLBRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1KyLBCol) Then
            fpsKetQua.SetFocus
    End If
End If
If KeyCode = vbKeyTab And Shift = 1 Then
    If fpsDkNgay.ActiveRow = f1dteTuNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteTuNCol) Then
        fpsLoaiTK.SetFocus
        With fpsLoaiTK
            .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
        End With
    End If
    If fpsDkNgay.ActiveRow = f1dteDeNRow And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber(f1dteDeNCol) Then
        fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteTuNCol), f1dteTuNRow
        End With
    End If
End If
End Sub
Private Sub fpsDkNgay_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With fpsDkNgay
         .Col = .ColLetterToNumber(f1dteTuNCol)
         .Row = f1dteTuNRow
           Dim vKyLB As Variant
           Dim arrDate() As String
           Dim vdtetun As Variant
            .GetText .Col, .Row, vdtetun
            If Not IsDate(vdtetun) And vdtetun <> "" Then
                .SetText .Col, .Row, ""
                DisplayMessage "0073", msOKOnly, miInformation
                Cancel = True
                'lerror = True
                .SetActiveCell .Col, .Row
                Exit Sub
            End If
            
            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
           Dim vdteden As Variant
            .GetText .Col, .Row, vdteden
            If Not IsDate(vdteden) And vdteden <> "" Then
                .SetText .Col, .Row, ""
                DisplayMessage "0073", msOKOnly, miInformation
                Cancel = True
                'lerror = True
                .SetFocus
                .SetActiveCell .Col, .Row
                'lerror = True
                Exit Sub
            End If
            
            ' Kiem tra ky lap bo
            .Col = .ColLetterToNumber(f1KyLBCol)
            .Row = f1KyLBRow
            .SetText .Col, .Row, Format_mmyyyy(CStr(.Text))
            .GetText .Col, .Row, vKyLB
            arrDate = Split(vKyLB, "/")
            If UBound(arrDate) > 0 Then
                If CInt(arrDate(0)) > 12 And vKyLB <> "" Then
                    DisplayMessage "0117", msOKOnly, miInformation
                    Cancel = True
                    'lerror = True
                    .SetFocus
                    .SetActiveCell .Col, .Row
                    'lerror = True
                    Exit Sub
                End If
            End If

    End With
    'lerror = False
End Sub

Private Sub fpsKetQua_Click(ByVal Col As Long, ByVal Row As Long)
    With fpsKetQua
        .ReDraw = False
        lngRowFocus = SetRowFocus(lngRowFocus, Row, True)
        .ReDraw = True
    End With
End Sub

Private Sub fpsKetQua_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = 1 Then
       fpsDkNgay.SetFocus
        With fpsDkNgay
            .SetActiveCell .ColLetterToNumber(f1dteDeNCol), f1dteDeNRow
        End With
    End If
    If KeyCode = vbKeyTab And Shift = 0 Then
       btnTraCuu(0).SetFocus
    End If
    If KeyCode = vbKeyDown And lngRowFocus < fpsKetQua.MaxRows Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus + 1)
    ElseIf KeyCode = vbKeyUp And lngRowFocus > 1 Then
        lngRowFocus = SetRowFocus(lngRowFocus, lngRowFocus - 1)
    End If
End Sub

Private Sub fpsLoaiTK_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    With fpsLoaiTK
        lmaTK = larrId(.TypeComboBoxCurSel)
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
        btnThoat(1).SetFocus
    End If
End Sub
Function KiemTraDKngay() As Boolean
With fpsDkNgay
        KiemTraDKngay = True
        .Col = .ColLetterToNumber(f1dteTuNCol)
        .Row = f1dteTuNRow
        Dim vdtetun As Variant
        .GetText .Col, .Row, vdtetun
        If vdtetun = "" Then
            DteTuN = format(lminDate, "DD/MM/YYYY")
        Else
            If IsDate(vdtetun) Then
                DteTuN = format(vdtetun, "DD/MM/YYYY")
            Else
                KiemTraDKngay = False
                Exit Function
            End If
        End If
        .Col = .ColLetterToNumber(f1dteDeNCol)
        .Row = f1dteDeNRow
        Dim vdteden As Variant
        .GetText .Col, .Row, vdteden
        If vdteden = "" Then
            DteDeN = format(lmaxDate, "DD/MM/YYYY")
        Else
            If IsDate(vdteden) Then
                DteDeN = format(vdteden, "DD/MM/YYYY")
            Else
                KiemTraDKngay = False
                Exit Function
            End If
        End If
        If DteTuN > DteDeN Then
            DisplayMessage "0071", msOKOnly
            KiemTraDKngay = False
            .SetActiveCell .ColLetterToNumber(f1dteDeNCol), f1dteDeNRow
        End If
        
         ' Ky lap bo
        .Col = .ColLetterToNumber(f1KyLBCol)
        .Row = f1KyLBRow
        .GetText .Col, .Row, kyLapBo
    End With
End Function

Private Function SetRowFocus(ByVal lngRow As Long, ByVal lngNewRow As Long, Optional ByVal blnClickEvent As Boolean = False) As Long
    With fpsKetQua
        .Col = -1
        .Row = lngRow
        .BackColor = vbWhite
        
        .Row = lngNewRow
        .BackColor = RGB(212, 343, 423)
        
        If blnClickEvent Then
            .SetActiveCell 8, lngNewRow
        Else
            .SetActiveCell 8, lngRow
        End If
        
    End With
    SetRowFocus = lngNewRow
End Function

Private Sub optIHTKK_Click()
    'strDaNhan = " IS NULL"
    strDaNhan = " = 'N'"
    If chkNhanTDong.Value Then
        chkCBXL.Enabled = True
    Else
        chkCBXL.Enabled = False
    End If
End Sub

Private Sub optRecv_Click()
    strDaNhan = " = 'Y'"
    chkCBXL.Value = 0
    chkCBXL.Enabled = False
End Sub



Function Format_mmyyyy(str As String) As String
    Dim m As String, y As String
    
    On Error GoTo e
    m = Left(str, InStr(str, "/") - 1)
    y = Right(str, Len(str) - InStr(str, "/"))
    y = Replace(y, ".", "")
    If IsNumeric(m) And IsNumeric(y) Then
        If Val(m) >= 1 And Val(m) <= 12 Then
            Format_mmyyyy = format(m, "0#")
        Else
            GoTo e
        End If
        
        If Val(y) >= 0 And Val(y) <= 9999 Then
            
            If Val(y) >= 0 And Val(y) <= 999 Then y = CStr(2000 + Val(y))
            If Val(y) < 1900 Then GoTo e
            Format_mmyyyy = Format_mmyyyy & "/" & format(y, "####")
        Else
            GoTo e
        End If
    End If
    Exit Function
e:
    Format_mmyyyy = ""
End Function
