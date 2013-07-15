VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTruyennhanTK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "M· sè thuÕ"
      Height          =   650
      Left            =   90
      TabIndex        =   14
      Top             =   1320
      Width           =   5445
      Begin VB.TextBox txtMST 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   15
         Top             =   240
         Width           =   3645
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tr¹ng th¸i tê khai"
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   2010
      Width           =   7335
      Begin VB.OptionButton optIHTKK 
         Caption         =   "Tê khai ch­a göi lªn Côc"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optRecv 
         Caption         =   "Tê khai ®· göi lªn Côc"
         Height          =   255
         Left            =   4080
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
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton btnTraCuu 
      Caption         =   "Tra cøu"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin FPUSpreadADO.fpSpread fpsKetQua 
      Height          =   3555
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   10635
      _Version        =   458752
      _ExtentX        =   18759
      _ExtentY        =   6271
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
      Height          =   1605
      Left            =   5580
      TabIndex        =   5
      Top             =   360
      Width           =   5385
      Begin FPUSpreadADO.fpSpread fpsDkNgay 
         Height          =   795
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4815
         _Version        =   458752
         _ExtentX        =   8493
         _ExtentY        =   1402
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
         SpreadDesigner  =   "frmTruyennhanTK.frx":0A6C
         UserResize      =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KÕt qu¶"
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2720
      Width           =   10875
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chän lo¹i tê khai"
      Height          =   885
      Left            =   90
      TabIndex        =   7
      Top             =   360
      Width           =   5445
      Begin FPUSpreadADO.fpSpread fpsLoaiTK 
         Height          =   525
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5205
         _Version        =   458752
         _ExtentX        =   9181
         _ExtentY        =   926
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
         SpreadDesigner  =   "frmTruyennhanTK.frx":108A
         UserResize      =   1
      End
   End
   Begin VB.CommandButton cmdNhanTkhai 
      Caption         =   "Göi tê khai"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
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
      Left            =   240
      TabIndex        =   8
      Top             =   6960
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
' Ky lap bo
Private Const f1KyLBRow = 4
Private Const f1KyLBCol = "C"


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
Private kyLapBo As Variant
Private lmaTK As Long
Private lSoBG As Long
Private larrId(lmaxSoTk) As Long
Private lerror As Boolean
Private strDaGui As String
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
    Dim strDaNhan As String
    'Xoa bo ket qua cu tren Grid
    'fpsKetQua.ClearRange 1, 1, fpsKetQua.MaxCols, lSoBG, True
    If Not KiemTraDKngay Then
         Exit Sub
    End If
    'connect to database QLT
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
    strDaNhan = "='T'"
    'Lay cau lenh truy van
    xmlSQL.Load App.path & "\SQL.xml"
    strSQL = GetAttribute(xmlSQL.childNodes(1), "SqlTraTKTN")
    strSQL = Replace(strSQL, "nhohon=", "<=")
    strSQL = Replace(strSQL, "ma_tkhai", "" & changeLoaiToKhai(lmaTK) & "")
    strSQL = Replace(strSQL, "str_tt_gui", "" & strDaGui & "")
    strSQL = Replace(strSQL, "strDa_Nhan", "" & strDaNhan & "")
    strSQL = Replace(strSQL, "ngay_nop_dau", "To_date('" & format(DteTuN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    strSQL = Replace(strSQL, "ngay_nop_cuoi", "To_date('" & format(DteDeN, "dd/mm/yyyy") & "','dd/mm/yyyy') ")
    If Trim(kyLapBo) <> "" And kyLapBo <> vbNullString Then
        strSQL = strSQL & " and  hdr.kylb_tu_ngay= to_date('" & "01/" & kyLapBo & "'," & "'dd/mm/yyyy')"
    End If
    
    strSQL = strSQL & " Order By hdr.loai_tkhai, hdr.ngay_nop"
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
                        .SetText lCol, lRow, TAX_Utilities_Srv_New.Convert(rsReturn.Fields(lIndex - 1).Value, TCVN, UNICODE)
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
             .MaxRows = lRow
        
            If strTracuu = "TC" Then
                lSoBG = 0
                DisplayMessage "0072", msOKOnly
            End If
        End If
       ' If strTracuu = "TC" Then
            LblSoBG.Visible = True
            LblSoBG.TextAlign = fmTextAlignLeft
            LblSoBG.caption = "Sè tê khai t×m thÊy: " & lSoBG
        'Else
        '    LblSoBG.Visible = False
        'End If
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
    
    If strLoaiMaToKhai = "46" Then changeLoaiToKhai = " ='01A_TNCN_BH11'"
    If strLoaiMaToKhai = "47" Then changeLoaiToKhai = " ='01B_TNCN_BH11'"
    If strLoaiMaToKhai = "48" Then changeLoaiToKhai = " ='01A_TNCN_XS11'"
    If strLoaiMaToKhai = "49" Then changeLoaiToKhai = " ='01B_TNCN_XS11'"
    
    If strLoaiMaToKhai = "74" Then changeLoaiToKhai = " ='08_TNCN11'"
    If strLoaiMaToKhai = "75" Then changeLoaiToKhai = " ='08A_TNCN11'"
    
    If strLoaiMaToKhai = "0" Then changeLoaiToKhai = " in ('02A_TNCN11','02B_TNCN11','03A_TNCN11','03B_TNCN11','07_TNCN11','01A_TNCN_BH11','01B_TNCN_BH11','01A_TNCN_XS11','01B_TNCN_XS11','08_TNCN11','08A_TNCN11') "
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
     Dim rsHDR As ADODB.Recordset
     Dim rsDTl As ADODB.Recordset
     Dim strSQLHdr As String, strSQLHdrTemp As String
     Dim strSQLInsPkgTmp As String, strSQLInsTupTmp As String, strSQLInsMupHDRTmp As String, strSQLInsMupDTLTmp As String
     Dim strSQLInsPkg As String, strSQLInsTup As String, strSQLInsMupHDR As String, strSQLInsMupDTL As String
     Dim strSQLDtl As String, strSQLDtlTemp As String
     Dim strSQLUpdate As String
     Dim strValueChk As Variant
'     'Bien luu du lieu bang HDR
     Dim strMst As Variant, strTen As Variant, strDiaChi As Variant
     Dim strLoaiTK As Variant, strNgayNop As Variant, strKyKKtu As Variant
     Dim strKyKKden As Variant, strKylbTu As Variant
     Dim strKylbDen As Variant, strNgayCN As Variant, strNguoiCN As Variant
     Dim strLoiDD As Variant, strLanQuet As Variant, strPhongXL As Variant
     Dim strKkbs As Variant, strKyLb As Variant, strKyKK As Variant
     Dim strTTHTK As Variant, strID As Variant, strMaCQT As Variant, strThueOnDinh As Variant
     Dim strDaiLyThue, strNgayHopDong, strSoHopDong, strLanBS As Variant
     Dim strHinhThucQT, strTKThangQuy As Variant
     Dim strHinhThucNop, strITkhaiID As Variant
'     'Bien luu du lieu DTL
     Dim strKyHieu As Variant, strGiaTri As Variant, strRowID As Variant
     Dim i As Integer
     Dim dataPkgId As String
     Dim tupId As String
     Dim mupId As String
     Dim noiLamViec As String
     Dim noiNhan As String
     Dim bln As Boolean
     Dim strCreateDate As Variant
     Dim strTnsCode As String
     Dim tranNum As Integer
     Dim maxRowSen As Integer
     Dim totalCount As Integer
     Dim countTKIns As Integer
     Dim numberPkg As Integer, stepPkg As Integer
     Dim flagPkgLast As Boolean
     Dim clsConn As New TAX_Utilities_Srv_New.clsADO
     ' dung transaction de inser tran_no
     Dim strSQLTransaction As String
     
On Error GoTo ErrHandle
    If Not clsDAO.Connected Then
        clsDAO.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsDAO.Connect
    End If
'
    ' Tao ket noi toi DB Cuc
    If Not clsConn.Connected Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
        clsConn.Connect
    End If

    xmlSQL.Load App.path & "\SQL.xml"
    strSQLHdrTemp = GetAttribute(xmlSQL.childNodes(1), "SqlDLGuiCuc")
    strSQLDtlTemp = GetAttribute(xmlSQL.childNodes(1), "SqlDLGuiCucDtl")
    strSQLInsPkgTmp = GetAttribute(xmlSQL.childNodes(1), "strInsData_pkg")
    strSQLInsTupTmp = GetAttribute(xmlSQL.childNodes(1), "strInsTup_exc")
    strSQLInsMupHDRTmp = GetAttribute(xmlSQL.childNodes(1), "strInsMupHDR_exc")
    strSQLInsMupDTLTmp = GetAttribute(xmlSQL.childNodes(1), "strInsMupDTL_exc")
    If IsNumeric(GetAttribute(xmlSQL.childNodes(1), "MaxRowSend")) Then
        maxRowSen = Val(GetAttribute(xmlSQL.childNodes(1), "MaxRowSend"))
    Else
        maxRowSen = 1000
    End If
    
    strTracuu = "NTK"
    strCreateDate = " sysdate "
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
            numberPkg = IIf((totalCount Mod maxRowSen) = 0, _
                            totalCount \ maxRowSen, totalCount \ maxRowSen + 1)
            countTKIns = 0
            stepPkg = 1
            flagPkgLast = True
            
            clsConn.BeginTrans
            noiLamViec = GetNoiLamViec
            noiNhan = GetNoiNhan(noiLamViec)
            For i = 2 To .MaxRows
'                noiLamViec = GetNoiLamViec
'                noiNhan = GetNoiNhan(noiLamViec)
                .GetText 2, i, strValueChk
                .GetText 11, i, strID
                strTnsCode = "PT"
                 If strValueChk = "1" Then
                    countTKIns = countTKIns + 1
                      ' Lay ID cua data_pkg
                      ' Begin
                      'clsConn.BeginTrans
                    If totalCount <= maxRowSen And flagPkgLast Then
                        dataPkgId = GetDataPkgId
                        tranNum = totalCount
                        ' Ghi du lieu vao data_pkg
                        strSQLInsPkg = strSQLInsPkgTmp
                        strSQLInsPkg = strSQLInsPkg & "'" & dataPkgId & "','" & strTnsCode & "'," & strCreateDate & "," & tranNum & ",0,'" & noiLamViec
                        strSQLInsPkg = strSQLInsPkg & "','" & noiNhan & "','0','" & dataPkgId & "'," & strCreateDate & ",'" & noiNhan & "','','','00','',0,0,0," & strCreateDate & ",0)"
                        bln = clsConn.ExecuteQuery(strSQLInsPkg)
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
                                bln = clsConn.ExecuteQuery(strSQLInsPkg)
                                ' end
                                stepPkg = stepPkg + 1
                          ElseIf stepPkg = numberPkg And (countTKIns > (stepPkg - 1) * maxRowSen) And flagPkgLast Then
                                dataPkgId = GetDataPkgId
                                tranNum = totalCount - (stepPkg - 1) * maxRowSen
                                ' Ghi du lieu vao data_pkg
                                strSQLInsPkg = strSQLInsPkgTmp
                                strSQLInsPkg = strSQLInsPkg & "'" & dataPkgId & "','" & strTnsCode & "'," & strCreateDate & "," & tranNum & ",0,'" & noiLamViec
                                strSQLInsPkg = strSQLInsPkg & "','" & noiNhan & "','0','" & dataPkgId & "'," & strCreateDate & ",'" & noiNhan & "','','','00','',0,0,0," & strCreateDate & ",0)"
                                bln = clsConn.ExecuteQuery(strSQLInsPkg)
                                ' end
                                stepPkg = stepPkg + 1
                                flagPkgLast = False
                          End If
  
                    End If
                    'clsConn.CommitTrans
                    
                    ' ghep voi dieu kien loc de lay ban ghi trong bang HDR
                    
                    strSQLHdr = strSQLHdrTemp
                    strSQLHdr = strSQLHdr + " where id = " & Val(Trim(CStr(strID)))
                    Set rsHDR = clsDAO.Execute(strSQLHdr)
                    ' ghi du lieu vao bang HDR tren Cuc
                    If rsHDR.Fields.Count > 0 Then
                        Do While Not rsHDR.EOF
                            'clsConn.BeginTrans
                            strMst = rsHDR.Fields(0).Value
                            strTen = IIf(IsNull(rsHDR.Fields(1)), "", rsHDR.Fields(1).Value)
                            strTen = Replace(strTen, "'", "''")
                            strDiaChi = IIf(IsNull(rsHDR.Fields(2)), "", rsHDR.Fields(2).Value)
                            strDiaChi = Replace(strDiaChi, "'", "''")
                            strLoaiTK = rsHDR.Fields(3).Value
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
                            strKkbs = rsHDR.Fields(14).Value
                            strID = rsHDR.Fields(15).Value
                            strThueOnDinh = rsHDR.Fields(16).Value
                            strDaiLyThue = rsHDR.Fields(17).Value
                            strSoHopDong = rsHDR.Fields(18).Value
                            strNgayHopDong = rsHDR.Fields(19).Value
                            strLanBS = rsHDR.Fields(20).Value
                            strHinhThucQT = rsHDR.Fields(21).Value
                            strTKThangQuy = rsHDR.Fields(22).Value
                            strITkhaiID = rsHDR.Fields(23).Value
                            strHinhThucNop = rsHDR.Fields(24).Value
                            
                            ' To khai 08/TNCN va 08A/TNCN
                            If strTKThangQuy = 1 Then
                                strKyKKtu = rsHDR.Fields(25).Value
                                strKyKKden = rsHDR.Fields(26).Value
                            End If
                            
                            If strTKThangQuy = vbNullString Or IsNull(strTKThangQuy) Then
                                strTKThangQuy = "null"
                            End If
                            
                            strMaCQT = strTaxOfficeId
                            
                            tupId = GetTranNo
       
                            ' Ghi du lieu vao bang tup_exc
                            strSQLInsTup = strSQLInsTupTmp
                            strSQLInsTup = strSQLInsTup & "'" & tupId & "','" & tupId & "','PT','" & dataPkgId & "',"
                            strSQLInsTup = strSQLInsTup & "'30','30','" & noiLamViec & "','" & noiLamViec & "','" & noiNhan & "',"
                            strSQLInsTup = strSQLInsTup & strCreateDate & ",'00','')"
                            'bln = clsConn.ExecuteQuery(strSQLInsTup)
                            
                            ' Ghi du lieu HDR vao bang mup_exc
                            mupId = GetMupId
                            strSQLInsMupHDR = strSQLInsMupHDRTmp
                            strSQLInsMupHDR = strSQLInsMupHDR & "'" & mupId & "','" & tupId & "'," & Trim(strID) & ",'" & Trim(strMaCQT) & "','" & Trim(strMst) & "','"
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strTen) & "','" & Trim(strLoaiTK) & "',"
                            strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strNgayNop & "','mm/dd/yyyy'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strKylbTu & "','mm/dd/yyyy'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strKylbDen & "','mm/dd/yyyy'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strKyKKtu & "','mm/dd/yyyy'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strKyKKden & "','mm/dd/yyyy'),"
                            strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strNgayCN & "','mm/dd/yyyy'),'"
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strNguoiCN) & "'," & Trim(strLanQuet) & ",'" & Trim(strPhongXL) & "',"
                            'strSQLInsMupHDR = strSQLInsMupHDR & Trim(strKkbs) & ",'','',0,'" & strThueOnDinh & "','"
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strKkbs) & ",'','"
                            strSQLInsMupHDR = strSQLInsMupHDR & strHinhThucNop & "',"
                            strSQLInsMupHDR = strSQLInsMupHDR & IIf(IsNull(strITkhaiID), "null", strITkhaiID) & ",'" & strThueOnDinh & "','"
                            
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strDaiLyThue) & "','"
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strSoHopDong) & "',"
                            If Trim(strNgayHopDong) = "" Or strNgayHopDong = vbNullString Then
                                strSQLInsMupHDR = strSQLInsMupHDR & " null,"
                            Else
                                strSQLInsMupHDR = strSQLInsMupHDR & " to_date('" & strNgayHopDong & "','mm/dd/yyyy'),"
                            End If
                            strSQLInsMupHDR = strSQLInsMupHDR & Trim(strLanBS) & ",'" & Trim(strHinhThucQT) & "'," & Trim(strTKThangQuy) & ")"
                            
                            'bln = clsConn.ExecuteQuery(strSQLInsMupHDR)
                            
                            ' add cac cau insert vao tran
                            strSQLTransaction = "begin"
                            strSQLTransaction = strSQLTransaction & vbCrLf & strSQLInsTup & ";"
                            strSQLTransaction = strSQLTransaction & vbCrLf & strSQLInsMupHDR & ";"
                            'end
                            ' Ghi du lieu DTL vao bang mup_exc
                            strSQLDtl = strSQLDtlTemp
                            strSQLDtl = strSQLDtl & " where hdr_id = " & Trim(strID)
                            Set rsDTl = clsDAO.Execute(strSQLDtl)
                            If rsDTl.Fields.Count > 0 Then
                                Do While Not rsDTl.EOF
                                    strGiaTri = rsDTl.Fields(0).Value
                                    strKyHieu = rsDTl.Fields(1).Value
                                    strRowID = rsDTl.Fields(2).Value
                                    strRowID = IIf(IsNull(strRowID), "", strRowID)
                                    mupId = GetMupId
                                    'Ghep chuoi cau insert DTl
                                    strSQLInsMupDTL = strSQLInsMupDTLTmp
                                    strSQLInsMupDTL = strSQLInsMupDTL & "'" & mupId & "','" & tupId & "'," & Trim(strID) & ",'" & Trim(strMaCQT) & "','"
                                    strSQLInsMupDTL = strSQLInsMupDTL & Trim(strLoaiTK) & "','" & Trim(strKyHieu) & "','" & Trim(strGiaTri) & "','" & Trim(strRowID) & "')"
                                    'bln = clsConn.ExecuteQuery(strSQLInsMupDTL)
                                    ' add cac cau insert vao tran
                                    strSQLTransaction = strSQLTransaction & vbCrLf & strSQLInsMupDTL & ";"
                                    ' end
                                    rsDTl.MoveNext
                                Loop
                            End If
                            'clsConn.BeginTrans
                            strSQLTransaction = strSQLTransaction & vbCrLf & "end;"
                            bln = clsConn.ExecuteQuery(strSQLTransaction)
                            'clsConn.CommitTrans
                            clsDAO.BeginTrans
                            ' update pkg_id
                            strSQLUpdate = "update rcv_tkhai_hdr set pkg_id = '" & dataPkgId & "' where  id = " & Trim(strID)
                            bln = clsDAO.ExecuteQuery(strSQLUpdate)
                            ' Set trang thai cua to khai da duoc chuyen len Cuc
                            ' HDR
                            strSQLUpdate = "update rcv_tkhai_hdr set tt_gui = 2 where  id = " & Trim(strID)
                            bln = clsDAO.ExecuteQuery(strSQLUpdate)
                            clsDAO.CommitTrans
                            'end set trang thai
                            rsHDR.MoveNext
                        Loop
                        
                    End If

                 End If
            Next
            clsConn.CommitTrans
        End With
        clsDAO.Disconnect
        clsConn.Disconnect
        ' Load lai danh sach to khai chua gui
        traCuuToKhai
    Else
        DisplayMessage "0131", msOKOnly, miCriticalError
        Exit Sub
    End If
    ' end
ErrHandle:
    If Err.Number <> 0 Then
        SaveErrorLog Me.Name, "cmdNhanTkhai_Click", Err.Number, Err.Description
    End If
    'Rollback du lieu tren DB vat
    ' Set trang thai cua to khai da duoc chuyen len Cuc
    ' HDR
    If clsDAO.Connected Then
        clsDAO.RollbackTrans
    End If
    If clsConn.Connected Then
        clsConn.RollbackTrans
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
        
        .Col = .ColLetterToNumber(f1KyLBCol)
        .Row = f1KyLBRow
        .BackColor = vbWhite
        
'        .TypePicDefaultText = "../...."
'        .TypePicMask = "99//9999"

        
        'SetDateFormat fpsDkNgay, 1, .ColLetterToNumber(f1KyLBCol), f1KyLBRow, "MM/YYYY"
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
            Dim vKyLB As Variant
            Dim arrDate() As String
            vdtehientai = Date
            vKyLB = format(vdtehientai, "mm/yyyy")
            
            .Col = .ColLetterToNumber(f1dteTuNCol)
            .Row = f1dteTuNRow
            .SetText .Col, .Row, vdtehientai

            .Col = .ColLetterToNumber(f1dteDeNCol)
            .Row = f1dteDeNRow
            .SetText .Col, .Row, vdtehientai
            ' Set gia tri cho ky lap bo
            
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
                Dim strID As Variant
                Parentid = GetAttribute(xmlNode, "ParentID")
                LoaiTk = GetAttribute(xmlNode, "Caption")
                strID = Val(GetAttribute(xmlNode, "ID"))
                '.TypeComboBoxIndex = 0
                If Parentid = "104" Then
                    If strID = 15 Or strID = 16 Or strID = 50 Or strID = 51 Or strID = 36 Or strID = 46 Or strID = 47 Or strID = 48 Or strID = 49 Or strID = 74 Or strID = 75 Then
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
        strDaGui = " =1 "
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
           Dim vdtetun As Variant
           Dim vKyLB As Variant
           Dim arrDate() As String
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
        If blnCheck And Trim(strDaGui) = "=1" Then
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
    strDaGui = " =1 "
End Sub

Private Sub optRecv_Click()
    strDaGui = " =2"
End Sub


'Sub SetDateFormat(FpSpd As fpSpread, SheetNumber As Integer, RowNumber As Long, ColNumber As Long, strFormat As String)
'    FpSpd.Sheet = SheetNumber
'    FpSpd.Row = RowNumber
'    FpSpd.Col = ColNumber
'    FpSpd.CellType = CellTypePic
'    ' Set the characters to center
'    FpSpd.TypeHAlign = TypeHAlignCenter
'    FpSpd.TypeVAlign = TypeHAlignCenter
'    FpSpd.TypePicDefaultText = "../../...."
'
'    Select Case LCase(strFormat)
'        Case LCase("DD/MM/YYYY")
'            FpSpd.TypePicMask = "99//99//9999"
'        Case LCase("DD/MM")
'            FpSpd.TypePicMask = "99//99"
'        Case LCase("MM/YYYY")
'            FpSpd.TypePicDefaultText = "../...."
'            FpSpd.TypePicMask = "99//9999"
'        Case LCase("YYYY")
'            FpSpd.TypePicDefaultText = "...."
'            FpSpd.TypePicMask = "9999"
'    End Select
'End Sub
'
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

