VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmTruyennhanTK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "M· sè thuÕ"
      Height          =   615
      Left            =   90
      TabIndex        =   14
      Top             =   1770
      Width           =   3465
      Begin VB.TextBox txtMST 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   930
         MaxLength       =   15
         TabIndex        =   15
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tr¹ng th¸i tê khai"
      Height          =   615
      Left            =   3630
      TabIndex        =   10
      Top             =   1770
      Width           =   7335
      Begin VB.OptionButton optIHTKK 
         Caption         =   "Tê khai ch­a göi lªn Côc"
         Height          =   255
         Left            =   540
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optRecv 
         Caption         =   "Tê khai ®· göi lªn Côc"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   2865
      End
   End
   Begin VB.CommandButton btnThoat 
      Caption         =   "&Tho¸t"
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   4
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton btnTraCuu 
      Caption         =   "Tra cøu"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin FPUSpreadADO.fpSpread fpsKetQua 
      Height          =   3675
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   10635
      _Version        =   458752
      _ExtentX        =   18759
      _ExtentY        =   6482
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
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
      GridColor       =   16777215
      MaxCols         =   11
      MaxRows         =   13
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmTruyennhanTK.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän ngµy nép vµ kú lËp bé"
      Height          =   1245
      Left            =   5580
      TabIndex        =   5
      Top             =   360
      Width           =   5385
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   915
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5175
         _Version        =   458752
         _ExtentX        =   9128
         _ExtentY        =   1614
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
         SpreadDesigner  =   "frmTruyennhanTK.frx":0A0A
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      Height          =   4095
      Left            =   90
      TabIndex        =   6
      Top             =   2370
      Width           =   10875
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chän lo¹i tê khai"
      Height          =   1245
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   5445
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   885
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5205
         _Version        =   458752
         _ExtentX        =   9181
         _ExtentY        =   1561
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
         SpreadDesigner  =   "frmTruyennhanTK.frx":0CA2
         UserResize      =   1
      End
   End
   Begin VB.CommandButton cmdNhanTkhai 
      Caption         =   "Göi tê khai"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   6600
      Width           =   1215
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3015
      BackColor       =   -2147483648
      VariousPropertyBits=   8388627
      Size            =   "5318;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCaption 
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   120
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
      Left            =   120
      Top             =   0
      Width           =   9195
   End
   Begin MSForms.Label LblSoBG 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   6600
      Width           =   3015
      BackColor       =   -2147483648
      VariousPropertyBits=   8388627
      Size            =   "5318;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTruyennhanTK"
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
Private Const mFormColor = -2147483633
Private Const lminDate = "01/01/1900"
Private Const lmaxDate = "31/12/3000"
Private Const lmaxSoTk = 1000
Private Const mHeaderColor = 16709097

Private lngRowFocus As Long
Private DteTuN As Date
Private DteDeN As Date


Private lmaTK As Long
Private lSoBG As Long
Private larrId(lmaxSoTk) As Long
Private lerror As Boolean
Private strDaNhan As String
Private strTracuu As String

Private Sub btnThoat_Click(Index As Integer)
    Unload Me
End Sub
Private Sub btnTraCuu_Click(Index As Integer)
    strTracuu = "TC"
    traCuuToKhai
End Sub

Public Sub traCuuToKhai()
    Dim rsReturn As New ADODB.Recordset
    Dim strSQL As String
    Dim lIndex As Long, lCtrl As Long
    Dim lCol As Long, lRow As Long
    
    Dim rsCLob As New ADODB.Recordset
    Dim strTemp As String
    Dim iCountClob As Integer
    
    Dim ngayNop As String
    Dim ngayNop1 As Variant
    Dim arrDate() As String
    
    'Xoa bo ket qua cu tren Grid
    'fpsKetQua.ClearRange 1, 1, fpsKetQua.MaxCols, lSoBG, True
    
    
    
    If Not KiemTraDKngay Then
         Exit Sub
    End If
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString spathVat & "\NTK_TG"
        clsDAO.Connect
    End If
    
    'kiem tra dieu kien ky lap bo
    Dim KLAPBO As Variant
    With fpsDkNgay
        .Col = .ColLetterToNumber("C")
        .Row = 4
        .GetText .Col, .Row, KLAPBO
     End With
     
    'strDaNhan = " = 'C'"
     'Lay cau lenh truy van
    xmlSQL.Load App.path & "\SQL.xml"
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraCuuTN")
    strSQL = Replace(strSQL, "nhohon=", "<=")
    strSQL = Replace(strSQL, "ma_tkhai", "" & changeLoaiToKhai(lmaTK) & "")
    strSQL = Replace(strSQL, "strDa_Nhan", "" & strDaNhan & "")
    strSQL = Replace(strSQL, "ngay_nop_dau", " CTOD('" & format(DteTuN, "mm/dd/yyyy") & "')")
    strSQL = Replace(strSQL, "ngay_nop_cuoi", " CTOD('" & format(DteDeN, "mm/dd/yyyy") & "')")
    If KLAPBO <> "" Then
    strSQL = strSQL & " And (Kylbo = '" & KLAPBO & "' ) "
    End If
    strSQL = strSQL & " Order By TIN, loai_tkhai "
 'Khoi tao bien
    lCol = 3
    lRow = 2
    lSoBG = 0
    'Thuc hien cau lenh sql
    Set rsReturn = clsDAO.Execute(strSQL)
    'Hien thi du lieu len grid
    With fpsKetQua
        'Xoa bo ket qua cu tren Grid
        .DeleteRows 2, fpsKetQua.MaxRows - 1
        .MaxRows = 2
        
        'Xoa trang thai nut Check
        .Col = 2
        .Row = 1
        .Value = "0"
    
        If rsReturn.Fields.Count > 0 Then
            Do While Not rsReturn.EOF
                If lRow > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
            
                .MaxRows = lRow
                .InsertRows lRow, 1
                .SetText 1, lRow, lSoBG + 1
                 'Check
                .Col = 2
                .Row = lRow
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .Lock = False
                
                
                For lIndex = 1 To rsReturn.Fields.Count
                    If Not (IsNull(rsReturn.Fields(lIndex - 1).Value)) Then
                        If lIndex - 1 = 5 Then
                            ngayNop = rsReturn.Fields(lIndex - 1).Value
                            arrDate = Split(ngayNop, "/")
                            ngayNop = arrDate(1) & "/" & arrDate(0) & "/" & arrDate(2)
                            'ngayNop = DateSerial(arrDate(2), arrDate(0), arrDate(1))
                            'ngayNop = format(ngayNop, "dd/mm/yyyy")
                            .SetText lCol, lRow, ngayNop
                        Else
                            .SetText lCol, lRow, TAX_Utilities_Svr_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
                        End If
                        lCol = lCol + 1
                        .RowHeight(lRow) = .MaxTextRowHeight(lRow)
                    End If
                Next lIndex
                lCol = 3
                lRow = lRow + 1
                
                rsReturn.MoveNext
                lSoBG = lSoBG + 1
            Loop
        Else
            If strTracuu = "TC" Then
                lSoBG = 0
                DisplayMessage "0072", msOKOnly
            End If
        End If
        If strTracuu = "TC" Then
            LblSoBG.Visible = True
            LblSoBG.TextAlign = fmTextAlignLeft
            LblSoBG.caption = "Sè tê khai t×m thÊy: " & lSoBG
        Else
            LblSoBG.Visible = False
        End If
    End With
    ' Bat nut gui to khai len
'    If lSoBG = 0 Then
'        cmdNhanTkhai.Enabled = False
'    Else
'        If optRecv.Value = True Then
'            cmdNhanTkhai.Enabled = False
'        Else
'            cmdNhanTkhai.Enabled = True
'
'        End If
'    End If
    'Chinh sua lai Grid ket cho dep
'    With fpsKetQua
'        If lSoBG <= 15 Then
'            Dim i As Integer
'            For i = 1 To 15 - lRow
'                .MaxRows = lSoBG + i + 1
'                .InsertRows lSoBG + i, 1
'            Next
'        Else
'            .RowHeight(lRow) = 0
'        End If
'
'        If lngRowFocus = 0 Then
'            lngRowFocus = SetRowFocus(1, 1, True)
'        ElseIf lngRowFocus > .MaxRows Then
'            lngRowFocus = SetRowFocus(1, .MaxRows, True)
'        Else
'            lngRowFocus = SetRowFocus(1, lngRowFocus, True)
'        End If
'        .SetFocus
'    End With
    clsDAO.Disconnect
End Sub


Private Function changeLoaiToKhai(ByVal strLoaiMaToKhai As String) As String
    If strLoaiMaToKhai = "15" Then changeLoaiToKhai = " = '02A_TNCN11'"
    If strLoaiMaToKhai = "16" Then changeLoaiToKhai = " = '02B_TNCN11'"
    If strLoaiMaToKhai = "50" Then changeLoaiToKhai = " ='03A_TNCN11'"
    If strLoaiMaToKhai = "51" Then changeLoaiToKhai = "  ='03B_TNCN11'"
    If strLoaiMaToKhai = "36" Then changeLoaiToKhai = " ='07_TNCN11'"
    If strLoaiMaToKhai = "74" Then changeLoaiToKhai = " ='08_TNCN11'"
    If strLoaiMaToKhai = "75" Then changeLoaiToKhai = " ='08A_TNCN11'"
    
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
'    If Not clsDAO.Connected Then
'        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
'        clsDAO.Connect
'    End If
'    ' Lay tat ca cac to khai trong cung mot ky tinh thue, cua mot loai to khai, cua mot nguoi nop thue va co so lan nop den thoi diem kiem tra
'    strSQL = "select ID from RCV_IHTKK_MVACH where TIN = '" & maSoThue & "' and MA_TKHAI = '" & maToKhai & "' and KY_KKHAI = '" & kyKeKhai & "' and DA_NHAN IS NULL and LAN_NOP < " & lanNop
'    Set rs = clsDAO.Execute(strSQL)
'
'    'Neu con to khac lan nop truoc do chua duoc chuyen vao QLT_NTK thi tra lai ket qua la true
'    If rs.Fields.Count > 0 Then
'        validToKhaiiHTKK = True
'    End If
End Function

Private Sub cmdNhanTkhai_Click()
    Dim sSQL            As String
    Dim rsHDR           As ADODB.Recordset
    Dim rsDTl           As ADODB.Recordset
    Dim maxR            As Integer
    Dim strSQLHdr       As String, strSQLHdrTemp As String
    Dim strSQLInsPkgTmp As String, strSQLInsTupTmp As String, strSQLInsMupHDRTmp As String, strSQLInsMupDTLTmp As String
    Dim strSQLInsPkg    As String, strSQLInsTup As String, strSQLInsMupHDR As String, strSQLInsMupDTL As String
    Dim strSQLDtl       As String, strSQLDtlTemp As String
    Dim strSQLUpdate    As String
    Dim strValueChk     As Variant
    '     'Bien luu du lieu bang HDR
    Dim strMST          As Variant, strTen As Variant, strDiaChi As Variant
    Dim strLoaiTK       As Variant, strNgayNop As Variant, strKyKKtu As Variant
    Dim strKyKKden      As Variant, strKylbTu As Variant
    Dim strKylbDen      As Variant, strNgayCN As Variant, strNguoiCN As Variant
    Dim strLoiDD        As Variant, strLanQuet As Variant, strPhongXL As Variant
    Dim strKkbs         As Variant, strKyLb As Variant, strKyKK As Variant
    Dim strTTHTK        As Variant, strID As Variant, strMaCQT As Variant, strThueOnDinh As Variant
    '     'Bien luu du lieu DTL
    '   dntai 06/02/2012 them bien rowID de luu gia tri truong rowID trong to 08A_TNCN
    Dim strKyHieu       As Variant, strGiaTri As Variant, rowID As Variant
    Dim i               As Integer
    Dim dataPkgId       As String
    Dim tupId           As String
    Dim mupId           As String
    Dim noiLamViec      As String
    Dim noiNhan         As String
    Dim bln             As Boolean
    Dim strCreateDate   As Variant
    Dim strTnsCode      As String
    Dim tranNum         As Integer
    Dim maxRowSen       As Integer
    Dim totalCount      As Integer
    Dim countTKIns      As Integer
    Dim numberPkg       As Integer, stepPkg As Integer
    Dim flagPkgLast     As Boolean
    Dim clsConn         As New TAX_Utilities_Svr_New.clsADO
    'pit
    Dim MADLT           As Variant
    Dim SOHDDL          As Variant
    Dim NGAYHDDL        As Variant
    Dim LANBS           As Variant
    
    On Error GoTo ErrHandle

    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString spathVat & "\NTK_TG"
        clsDAO.Connect
    End If

    '
    ' Tao ket noi toi DB Cuc
    If Not clsConn.Connected Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI"
        clsConn.Connect
    End If

    xmlSQL.Load App.path & "\SQL.xml"
    strSQLHdrTemp = GetAttribute(xmlSQL.childNodes(1), "SqlDLGuiCuc")
    
    strSQLDtlTemp = GetAttribute(xmlSQL.childNodes(1), "SqlDLGuiCucDtl")
    strSQLInsPkgTmp = GetAttribute(xmlSQL.childNodes(1), "strInsData_pkg")
    strSQLInsTupTmp = GetAttribute(xmlSQL.childNodes(1), "strInsTup_exc")
    strSQLInsMupHDRTmp = GetAttribute(xmlSQL.childNodes(1), "strInsMupHDR_DTL_exc")

    'strSQLInsMupDTLTmp = GetAttribute(xmlSQL.childNodes(1), "strInsMupDTL_exc")
    If IsNumeric(GetAttribute(xmlSQL.childNodes(1), "MaxRowSend")) Then
        maxRowSen = Val(GetAttribute(xmlSQL.childNodes(1), "MaxRowSend"))
    Else
        maxRowSen = 1000
    End If
    
    strTracuu = "NTK"
    strCreateDate = DateSerial(Int(Year(Now())), Int(Month(Now())), Int(Day(Now())))
    strCreateDate = "CTOD('" & format(strCreateDate, "mm/dd/yyyy") & "')"
    tranNum = 0

    ' duyet nhung ban ghi duoc check de chuyen len Cuc
    If clsConn.Connected Then

        With fpsKetQua
            totalCount = 0

            For i = 2 To .MaxRows
                .GetText 2, i, strValueChk

                If strValueChk = "1" Then
                    totalCount = totalCount + 1
                End If

            Next

            numberPkg = IIf((totalCount Mod maxRowSen) = 0, totalCount \ maxRowSen, totalCount \ maxRowSen + 1)
            countTKIns = 0
            stepPkg = 1
            flagPkgLast = True

            For i = 2 To .MaxRows
                noiLamViec = GetNoiLamViec
                noiNhan = GetNoiNhan(noiLamViec)
                .GetText 2, i, strValueChk
                .GetText 3, i, strMST
                .GetText 5, i, strKyKK
                .GetText 6, i, strKyLb
                .GetText 7, i, strLoaiTK
                .GetText 8, i, strNgayNop
                .GetText 9, i, strTTHTK
                .GetText 10, i, strLanQuet
                .GetText 11, i, strID
                strTnsCode = "PT" 'GetTnsCode(changeTK2TabCode(Trim(CStr(strLoaiTK))))

                If strValueChk = "1" Then
                    countTKIns = countTKIns + 1

                    ' Lay ID cua data_pkg
                    If totalCount <= maxRowSen And flagPkgLast Then
                        dataPkgId = GetDataPkgId
                        tranNum = totalCount
                        ' Ghi du lieu vao data_pkg
                        strSQLInsPkg = strSQLInsPkgTmp
                        strSQLInsPkg = strSQLInsPkg & "'" & dataPkgId & "','" & strTnsCode & "'," & strCreateDate & "," & tranNum & ",0,'" & noiLamViec
                        strSQLInsPkg = strSQLInsPkg & "','" & noiNhan & "','0','" & dataPkgId & "'," & strCreateDate & ",'" & noiNhan & "','','','00','',0,0,0," & strCreateDate & ",0)"
                        bln = clsConn.ExecuteDLL(strSQLInsPkg)
                        ' end
                        flagPkgLast = False
                    ElseIf totalCount > maxRowSen And flagPkgLast Then

                        If stepPkg < numberPkg And ((countTKIns = (stepPkg - 1) * maxRowSen + 1) Or countTKIns = 1) Then
                            dataPkgId = GetDataPkgId
                            tranNum = maxRowSen
                            ' Ghi du lieu vao data_pkg
                            strSQLInsPkg = strSQLInsPkgTmp
                            strSQLInsPkg = strSQLInsPkg & "'" & dataPkgId & "','" & strTnsCode & "'," & strCreateDate & "," & tranNum & ",0,'" & noiLamViec
                            strSQLInsPkg = strSQLInsPkg & "','" & noiNhan & "','0','" & dataPkgId & "'," & strCreateDate & ",'" & noiNhan & "','','','00','',0,0,0," & strCreateDate & ",0)"
                            bln = clsConn.ExecuteDLL(strSQLInsPkg)
                            ' end
                            stepPkg = stepPkg + 1
                        ElseIf stepPkg = numberPkg And (countTKIns > (stepPkg - 1) * maxRowSen) And flagPkgLast Then
                            dataPkgId = GetDataPkgId
                            tranNum = totalCount - (stepPkg - 1) * maxRowSen
                            ' Ghi du lieu vao data_pkg
                            strSQLInsPkg = strSQLInsPkgTmp
                            strSQLInsPkg = strSQLInsPkg & "'" & dataPkgId & "','" & strTnsCode & "'," & strCreateDate & "," & tranNum & ",0,'" & noiLamViec
                            strSQLInsPkg = strSQLInsPkg & "','" & noiNhan & "','0','" & dataPkgId & "'," & strCreateDate & ",'" & noiNhan & "','','','00','',0,0,0," & strCreateDate & ",0)"
                            bln = clsConn.ExecuteDLL(strSQLInsPkg)
                            ' end
                            stepPkg = stepPkg + 1
                            flagPkgLast = False
                        End If
  
                    End If

                    ' ghep voi dieu kien loc de lay ban ghi trong bang HDR
                    strSQLHdr = strSQLHdrTemp
                    strSQLHdr = strSQLHdr + " where id = " & Val(Trim(CStr(strID)))
                    Set rsHDR = clsDAO.Execute(strSQLHdr)

                    ' ghi du lieu vao bang HDR tren Cuc
                    If rsHDR.Fields.Count > 0 Then

                        Do While Not rsHDR.EOF
                            strMST = rsHDR.Fields(0).Value
                            strTen = rsHDR.Fields(1).Value
                            strDiaChi = rsHDR.Fields(2).Value
                            strNgayNop = rsHDR.Fields(4).Value
                            strKylbTu = rsHDR.Fields(5).Value
                            strKylbDen = rsHDR.Fields(6).Value
                            strKyKKtu = rsHDR.Fields(7).Value
                            strKyKKden = rsHDR.Fields(8).Value
                            strNgayCN = rsHDR.Fields(9).Value
                            strNguoiCN = rsHDR.Fields(10).Value
                            strLoiDD = rsHDR.Fields(11).Value
                            strLanQuet = rsHDR.Fields(12).Value
                            strPhongXL = rsHDR.Fields(13).Value
                            strID = rsHDR.Fields(15).Value
                            strMaCQT = rsHDR.Fields(16).Value
                            strThueOnDinh = rsHDR.Fields(17).Value

                            ' bo sung them 4 trg cua pit

'                            If Trim(SOHDDL) = vbNullString Then SOHDDL = "''"
                            MADLT = rsHDR.Fields(18).Value
                            SOHDDL = rsHDR.Fields(19).Value
                            NGAYHDDL = rsHDR.Fields(20).Value
                            LANBS = rsHDR.Fields(21).Value
                            If Trim(format(NGAYHDDL, "mm/dd/yyyy")) = vbNullString Then
                                NGAYHDDL = "12/30/1899"
                            End If

                            '                            If Trim(NGAYHDDL) = vbNullString Then
                            '                                NGAYHDDL = "CTOD('')"
                            '                            Else
                            '                                ' NGAYHDDL = ToDate(Trim(NGAYHDDL), DDMMYYYY)
                            '                                NGAYHDDL = "CTOD('" & format(NGAYHDDL, "mm/dd/yyyy") & "')"
                            '                            End If
                            
                            tupId = GetTranNo
                            
                            If strTTHTK = "2" Then
                                strKkbs = "1"
                            Else
                                strKkbs = "0"
                            End If

                            ' Ghi du lieu vao bang tup_exc
                            strSQLInsTup = strSQLInsTupTmp
                            strSQLInsTup = strSQLInsTup & "'" & tupId & "','" & tupId & "','PT','" & dataPkgId & "',"
                            strSQLInsTup = strSQLInsTup & "'30','30','" & noiLamViec & "','" & noiLamViec & "','" & noiNhan & "',"
                            strSQLInsTup = strSQLInsTup & strCreateDate & ",'00','')"
                            bln = clsConn.ExecuteDLL(strSQLInsTup)
                            ' Ghi du lieu HDR vao bang mup_exc
                            mupId = GetMupId
                            strSQLInsMupHDR = strSQLInsMupHDRTmp
                            strSQLInsMupHDR = strSQLInsMupHDR & "'" & mupId & "','" & tupId & "','" & Trim(strMaCQT) & "','" & Trim(strMST) & "','"
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strTen) & "','" & Trim(strLoaiTK) & "','" & Trim(strNguoiCN) & "','" & Trim(strPhongXL) & "','',"
                            strSQLInsMupHDR = strSQLInsMupHDR & "'','" & Trim(strThueOnDinh) & "','','','','','','" & Trim(MADLT) & "','" & Trim(SOHDDL) & "','','','','','','','','','','','','','','',"
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strID) & "," & Trim(strLanQuet) & "," & Trim(strKkbs) & ",0,0," & Trim(LANBS) & ",0,0,0,0,0,0,0,0,0,"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strNgayNop, "mm/dd/yyyy") & "'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKylbTu, "mm/dd/yyyy") & "'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKylbDen, "mm/dd/yyyy") & "'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKyKKtu, "mm/dd/yyyy") & "'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKyKKden, "mm/dd/yyyy") & "'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strNgayCN, "mm/dd/yyyy") & "'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(NGAYHDDL, "mm/dd/yyyy") & "'),CTOD(''),CTOD(''),CTOD(''))"
                            
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & "'" & mupId & "','" & tupId & "'," & Trim(strID) & ",'" & Trim(strMaCQT) & "','" & Trim(strMst) & "','"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strTen) & "','" & Trim(strLoaiTK) & "',"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strNgayNop, "mm/dd/yyyy") & "'),"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKylbTu, "mm/dd/yyyy") & "'),"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKylbDen, "mm/dd/yyyy") & "'),"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKyKKtu, "mm/dd/yyyy") & "'),"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strKyKKden, "mm/dd/yyyy") & "'),"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(strNgayCN, "mm/dd/yyyy") & "'),'"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strNguoiCN) & "'," & Trim(strLanQuet) & ",'" & Trim(strPhongXL) & "',"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strKkbs) & ",'','',0,'" & strThueOnDinh & "',"
                            '                            pit
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(MADLT) & "," & Trim(SOHDDL) & ","
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & " CTOD('" & format(NGAYHDDL, "mm/dd/yyyy") & "'),"
                            '                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(LANBS) & ")"
                            bln = clsConn.ExecuteDLL(strSQLInsMupHDR)
                            ' Ghi du lieu vao bang mup_exc
                            'dntai 06/02/2012 vi to 08A_TNCN luu vao bang # nen phai sua lai cau truy van
                            If Trim(strLoaiTK) = "08A_TNCN11" Then
                                strSQLDtlTemp = GetAttribute(xmlSQL.childNodes(1), "SqlDLGuiCucDtl_pl")
                            Else
                                strSQLDtlTemp = GetAttribute(xmlSQL.childNodes(1), "SqlDLGuiCucDtl")
                            End If
                            strSQLDtl = strSQLDtlTemp
                            strSQLDtl = strSQLDtl & " where hdr_id = " & Trim(strID)
                            Set rsDTl = clsDAO.Execute(strSQLDtl)

                            If rsDTl.Fields.Count > 0 Then
                                
                                Do While Not rsDTl.EOF
                                    strGiaTri = rsDTl.Fields(0).Value
                                    strKyHieu = rsDTl.Fields(1).Value
                                    'lay rowID trong to 08A_TNCN
                                    If Trim(strLoaiTK) = "08A_TNCN11" Then
                                        rowID = rsDTl.Fields(2).Value
                                        rowID = "'" & Trim(rowID) & "'"
                                    Else
                                        rowID = "''"
                                    End If
                                    mupId = GetMupId
                                    'Ghep chuoi cau insert DTl
                                    strSQLInsMupDTL = strSQLInsMupHDRTmp
                                    strSQLInsMupDTL = strSQLInsMupDTL & "'" & mupId & "','" & tupId & "','','','','','','','','','','" & Trim(strMaCQT) & "','"
                                    strSQLInsMupDTL = strSQLInsMupDTL & Trim(strLoaiTK) & "','" & Trim(strKyHieu) & "','" & Trim(strGiaTri) & "'," & rowID & ","
                                    strSQLInsMupDTL = strSQLInsMupDTL & "'','','','','','','','','','','','','','','','',"
                                    strSQLInsMupDTL = strSQLInsMupDTL & "0,0,0,0," & Trim(strID) & ",0,0,0,0,0,0,0,0,0,0,"
                                    strSQLInsMupDTL = strSQLInsMupDTL & "CTOD('12/30/1899'),CTOD('12/30/1899'),CTOD('12/30/1899'),CTOD('12/30/1899'),CTOD('12/30/1899'),CTOD('12/30/1899'),CTOD('12/30/1899'),CTOD(''),CTOD(''),CTOD(''))"
                                    '                                    strSQLInsMupDTL = strSQLInsMupDTLTmp
                                    '                                    strSQLInsMupDTL = strSQLInsMupDTL & "'" & mupId & "','" & tupId & "'," & Trim(strID) & ",'" & Trim(strMaCQT) & "','"
                                    '                                    strSQLInsMupDTL = strSQLInsMupDTL & Trim(strLoaiTK) & "','" & Trim(strKyHieu) & "','" & Trim(strGiaTri) & "','')"
                                    bln = clsConn.ExecuteDLL(strSQLInsMupDTL)
                                    rsDTl.MoveNext
                                Loop

                            End If

                            ' update pkg_id
                            strSQLUpdate = "update tmp_tncn_hdr set pkg_id = '" & dataPkgId & "' where  id = " & Trim(strID)
                            bln = clsDAO.ExecuteDLL(strSQLUpdate)
                            ' Set trang thai cua to khai da duoc chuyen len Cuc
                            ' HDR
                            strSQLUpdate = "update tmp_tncn_hdr set da_nhan = 1 where  id = " & Trim(strID)
                            bln = clsDAO.ExecuteDLL(strSQLUpdate)
                            'DTL
                            strSQLUpdate = "update tmp_tncn_dtl set danhan = 1 where hdr_id = " & Trim(strID)
                            bln = clsDAO.ExecuteDLL(strSQLUpdate)
                            'end set trang thai
                            rsHDR.MoveNext
                        Loop

                    End If

                End If

            Next

        End With

        clsDAO.Disconnect
        ' Load lai danh sach to khai chua gui
        traCuuToKhai
    Else
        DisplayMessage "0131", msOKOnly, miCriticalError
        Exit Sub
    End If

    ' end
ErrHandle:
    SaveErrorLog Me.Name, "cmdNhanTkhai_Click", Err.Number, Err.Description

    'Rollback du lieu tren DB vat
    ' Set trang thai cua to khai da duoc chuyen len Cuc
    ' HDR
    If clsDAO.Connected Then
        strSQLUpdate = "update tmp_tncn_hdr set da_nhan = 0, pkg_id = '' where  id = " & Trim(strID)
        clsDAO.ExecuteDLL (strSQLUpdate)
        'DTL
        strSQLUpdate = "update tmp_tncn_dtl set danhan = 0 where hdr_id = " & Trim(strID)
        clsDAO.ExecuteDLL (strSQLUpdate)
        clsDAO.Disconnect
        'end set trang thai
    End If

End Sub







Private Sub Form_Load()
    SetControlCaption Me, "frmTruyennhanTK"
    FormatGrid
    SetupData
    'lerror = False
    With fpsLoaiTK
        .SetActiveCell .ColLetterToNumber(f4cboLoTCol), f4cboLoTRow
    End With
End Sub
Sub FormatGrid()
    Dim i As Integer
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
        
        .Col = .ColLetterToNumber("C")
        .Row = 4
        .BackColor = vbWhite
    End With
    With fpsKetQua
        .MaxCols = 11
        .EditModePermanent = True
        .EditModeReplace = True
        .CursorType = CursorTypeLockedCell
        .CursorStyle = CursorStyleArrow
        .TypeNumberNegStyle = TypeNumberNegStyle1
        .ColWidth(12) = 0
        .Row = 1
        .RowHeight(1) = 25

        For i = 1 To .MaxCols
            .Col = i
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignCenter
            .BackColor = mHeaderColor
        Next
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
            vdtehientai = Date
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .SetText .Col, .Row, vdtehientai

            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .SetText .Col, .Row, vdtehientai
            
            .Col = .ColLetterToNumber("C")
            .Row = 4
            .SetText .Col, .Row, format(Date, "mm/yyyy")
            
        End With
    'Lay du lieu cho cbo
        With fpsLoaiTK
            .Col = .ColLetterToNumber(f4cboLoTCol)
            .Row = f4cboLoTRow
            Dim xmlDocument As New MSXML.DOMDocument
            Dim xmlNode As MSXML.IXMLDOMNode
            Dim strDataFileName As String
            Dim i As Integer
            
            xmlDocument.Load TAX_Utilities_Svr_New.GetAbsolutePath("menu.xml")
            Set xmlNodeListMenu = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
            i = 0
            larrId(0) = 0
            For Each xmlNode In xmlNodeListMenu
                Dim LoaiTk As String
                Dim Parentid As String
                Dim strID As Variant
                Parentid = GetAttribute(xmlNode, "ParentID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                strID = Val(GetAttribute(xmlNode, "ID"))
                '.TypeComboBoxIndex = 0
                If Parentid = "101_11" Then
                    If strID = 15 Or strID = 16 Or strID = 50 Or strID = 51 Or strID = 36 Or strID = 74 Or strID = 75 Then
                        i = i + 1
                        .TypeComboBoxIndex = -1
                        .TypeComboBoxString = LoaiTk
                        larrId(i) = Val(GetAttribute(xmlNode, "ID"))
                    End If
                End If
            Next
            .TypeComboBoxCurSel = 0
            Set xmlNode = Nothing
            Set xmlDocument = Nothing
        End With
        strDaNhan = " =0 "
End Sub

Private Sub Form_Resize()
    SetFormCaption Me, imgCaption, lblCaption
    frmTruyennhanTK.Top = (frmSystem.Height - frmTruyennhanTK.Height) / 2
    frmTruyennhanTK.Left = (frmSystem.Width - frmTruyennhanTK.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lngRowFocus = 0
End Sub

Private Sub fpsDkNgay_GotFocus()
    'btnTraCuu(0).Default = True
End Sub
Private Sub fpsDkNgay_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab And Shift = 0 Then 'And Not lerror
    If fpsDkNgay.ActiveRow = 4 And fpsDkNgay.ActiveCol = fpsDkNgay.ColLetterToNumber("C") Then
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
            
            Dim KLAPBO As Variant
            Dim arrDate() As String
            .Col = .ColLetterToNumber("C")
            .Row = 4
            .GetText .Col, .Row, KLAPBO
            If Trim(KLAPBO) <> "" And KLAPBO <> "../...." Then
                 arrDate = Split(KLAPBO, "/")
                 
                If CInt(arrDate(0)) <= 0 Or CInt(arrDate(0)) > 12 Then
                     DisplayMessage "0142", msOKOnly, miInformation
                     Cancel = True
                     .SetActiveCell .Col, .Row
                     Exit Sub
                End If
            End If
    End With
    'lerror = False
End Sub

Private Sub fpsKetQua_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim vxoatat As Variant, varValue As Variant
    Dim i As Integer, blnCheck As Boolean
    
    With fpsKetQua
        GetCellSpan fpsKetQua, Col, Row
        If Col = 2 Then
            .EventEnabled(EventAllEvents) = False
            If Row = 1 Then
                .GetText Col, Row, vxoatat
                If vxoatat = "1" Then
                    For i = 2 To lSoBG + 1
                        .SetText Col, i, "1"
                        blnCheck = True
                    Next
                Else
                    For i = 2 To lSoBG + 1
                        .SetText Col, i, "0"
                    Next
                    blnCheck = False
                End If
            Else
                blnCheck = False
                .SetText Col, 1, "1"
                For i = 2 To lSoBG + 1
                    .GetText Col, i, varValue
                    If CStr(varValue) = "0" Or CStr(varValue) = "" Then
                        .SetText Col, 1, "0"
                        'Exit For
                    Else
                        blnCheck = True
                    End If
                Next i
            End If
            .EventEnabled(EventAllEvents) = True
        End If
        
        If Row > 1 Then
            .ReDraw = False
            lngRowFocus = SetRowFocus(lngRowFocus, Row, True)
            .ReDraw = True
        End If
        
        'Set enable to cmdNhanTkhai
        If blnCheck And Trim(strDaNhan) = "=0" Then
            cmdNhanTkhai.Enabled = True
        Else
            cmdNhanTkhai.Enabled = False
        End If
        '********************************************
    End With
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
        
        Dim KYLBO As Variant
        .GetText .ColLetterToNumber("C"), 3, KYLBO
        
        
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
    strDaNhan = " =0 "
End Sub

Private Sub optRecv_Click()
    strDaNhan = " =1"
End Sub

' Ham tra ve ten cua doi tuong nop thue
Private Function getTENDTNT(ByVal maDTNT As String) As String
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    If Not clsConn.Connected Then
        clsConn.CreateConnectionString spathVat & "\DTNT\"
        clsConn.Connect
    End If
    'Lay cau lenh truy van
    strSQL = "select TENGOI from DTNT2 where MADTNT ='" & Trim(maDTNT) & "'"
    Set rs = clsConn.Execute(strSQL)
    If Not rs Is Nothing Then
        If Not IsNull(Trim(rs.Fields("TENGOI"))) Then
            getTENDTNT = rs.Fields("TENGOI")
        Else
            getTENDTNT = ""
        End If
    End If
    clsConn.Disconnect
End Function
