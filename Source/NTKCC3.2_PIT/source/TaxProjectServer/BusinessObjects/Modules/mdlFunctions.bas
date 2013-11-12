Attribute VB_Name = "mdlFunctions"
Public clsDAO As New TAX_Utilities_Svr_New.clsADO

Public Function GetAttribute(xmlNodeCell As MSXML.IXMLDOMNode, pAttributeName As String) As String
    On Error Resume Next
    GetAttribute = Replace(xmlNodeCell.Attributes.getNamedItem(pAttributeName).nodeValue, "'", "''")
End Function

Public Function GetMessageCellById(ByVal strId As String) As MSXML.IXMLDOMNode
    Dim xmlInforNode As MSXML.IXMLDOMNode
   
    For Each xmlInforNode In TAX_Utilities_Svr_New.NodeMessage
        If GetAttribute(xmlInforNode, "ID") = strId Then
            Set GetMessageCellById = xmlInforNode
            Exit Function
        End If
    Next
End Function

'Public Sub CellEditFormatNumber(fps As fpSpread, ByVal lSheet As Long, ByVal lCol As Long, ByVal lRow As Long, ByVal intKeyAscii As Integer, Optional blnEventEnable As Boolean = False, Optional evEventType As EVENTENABLEDConstants = EventAllEvents)
'    Dim strNumber As String, varValue As Variant
'
'    strNumber = "0123456789"
'    With fps
'        If Not blnEventEnable Then .EventEnabled(evEventType) = False
'
'        .Sheet = lSheet
'        .GetText lCol, lRow, varValue
'        If InStr(1, strNumber, Chr$(intKeyAscii)) <> 0 And Len(CStr(varValue)) < 10 Then
'            'GetCellSpan fps, lCol, lRow
'            .SetText lCol, lRow, CStr(varValue) & Val(Chr$(intKeyAscii))
'            UpdateCell fps, lCol, lRow, CStr(varValue) & Val(Chr$(intKeyAscii))
'        ElseIf intKeyAscii = vbKeyBack And Len(CStr(varValue)) > 0 Then
'            .SetText lCol, lRow, Mid$(CStr(varValue), 1, Len(CStr(varValue)) - 1)
'            UpdateCell fps, lCol, lRow, Mid$(CStr(varValue), 1, Len(CStr(varValue)) - 1)
'        Else
'            .SetText lCol, lRow, CStr(varValue)
'        End If
'
'        If Not blnEventEnable Then .EventEnabled(evEventType) = True
'    End With
'End Sub

''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 fps      : fpspread that you want to handle
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Public Sub UpdateCell(fps As fpSpread, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_Svr_New.Data(fps.Sheet - 1).nodeFromID(GetCellID(fps, pCol, pRow))
    If Not xmlNodeCell Is Nothing Then
        xmlNodeCell.Attributes.getNamedItem("Value").nodeValue = pValue
    End If
        
    Set xmlNodeCell = Nothing
End Sub

''' GetCellSpan description
''' Get cell span of current cell
''' Parameter1 pGrid    : the current fpSpread grid (input value)
''' Parameter2 pCol     : the current column (input/ output value)
''' Parameter3 pRow     : the current row (input/ output value)
''' Parameter4 pNumsRow : number of row with span (output value)
''' Parameter5 pNumsCol : number of column with span (output value)
Public Sub GetCellSpan(pGrid As fpSpread, pCol As Long, pRow As Long, Optional pNumsRow As Variant, Optional pNumsCol As Variant)
    On Error GoTo ErrorHandle
    
    Dim lRowAnchor As Variant, lColAnchor As Variant
    
    pGrid.GetCellSpan pCol, pRow, lColAnchor, lRowAnchor, pNumsCol, pNumsRow
    If lRowAnchor <> -1 And lColAnchor <> -1 Then
        pRow = Val(lRowAnchor)
        pCol = Val(lColAnchor)
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "GetCellSpan", Err.Number, Err.Description
End Sub

''' GetCellID description
''' Get CellID of current cell
''' Parameter1 pGrid    : the current fpSpread grid (input value)
''' Parameter2 pCol     : the current column (input value)
''' Parameter3 pRow     : the current row (input value)
Public Function GetCellID(pGrid As fpSpread, ByVal pCol As Long, ByVal pRow As Long) As String
    GetCellID = pGrid.ColNumberToLetter(pCol) & "_" & CStr(pRow)
End Function

'*******************************************************
'Description: SaveErrorLog sub write errors to log file
'Author:TuanLM
'Date:17/10/2005
'Paramter: pFormName: form has error
'          pFunctionName: function has error
'          pErrorNumber: error number
'          pErrorDesc: description of error
'Return:
'*******************************************************
Public Sub SaveErrorLog(pFormName As String, pFunctionName As String, pErrorNumber As Long, pErrorDesc As String)
    Dim msg As String
    Dim FileNum As Long
    Dim path As String
    path = App.path & "\LogFile.txt"
    msg = Now & " " & pFormName & " " & pFunctionName & vbCrLf
    msg = msg & str(pErrorNumber) & " : " & pErrorDesc
    FileNum = FreeFile
    Open path For Append As FileNum
        Print #FileNum, msg
        Print #FileNum, "------------------------------------------------------------"
    Close #FileNum
End Sub

Public Function GetCatalogueFileName(Optional lSheet As Long = 1) As String
    Dim strCatalogueName As String, strCatalogueID As String
    Dim xmlCatalogeValidNode As MSXML.IXMLDOMNode
    
    'Get valid catalogue node
    Set xmlCatalogeValidNode = GetValidityNode("102", TAX_Utilities_Svr_New.Month, TAX_Utilities_Svr_New.ThreeMonths, TAX_Utilities_Svr_New.Year)
    
    'Get catalogue ID
    strCatalogueID = GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "CatalogueID")
    
    'Get catalogue pattern name
    strCatalogueName = GetCatalogueName(xmlCatalogeValidNode, strCatalogueID)
    
    GetCatalogueFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet - 1), "TemplateFolder") & _
        strCatalogueName & ".xml")
End Function

Private Function GetCatalogueName(xmlCatalogueNode As MSXML.IXMLDOMNode, strId As String) As String
Dim xmlNode As MSXML.IXMLDOMNode

For Each xmlNode In xmlCatalogueNode.childNodes
    If GetAttribute(xmlNode, "ID") = strId Then
        GetCatalogueName = GetAttribute(xmlNode, "DataFile")
        Exit Function
    End If
Next
End Function

Public Function GetValidityNode(Id As String, Optional strMonth As String, Optional strThreeMonths As String, Optional strYear As String) As MSXML.IXMLDOMNode
    On Error GoTo ErrorHandle
    Dim xmlNodeListValidity As MSXML.IXMLDOMNodeList
    Dim xmlNodeValidity As MSXML.IXMLDOMNode
    
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlDomMenu As New MSXML.DOMDocument
    Dim xmlNodeListMenu As MSXML.IXMLDOMNodeList
    
    Dim ValidityDate As Date, StartDate As Date, MaxDate As Date
    Dim strNgayTaiChinh As String
    Dim iNgayTaiChinh As Integer
    Dim iThangTaiChinh As Integer
    
    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "FinanceYear") = "1" Then
        strNgayTaiChinh = TAX_Utilities_Svr_New.FinanceStartDate
        iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
        iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
    Else
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
    End If
    
    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
        Select Case strMonth
            Case "01", "03", "05", "07", "08", "10", "12"
                ValidityDate = Format("31/" & strMonth & "/" & strYear, "dd/mm/yyyy")
            Case "02"
                 If CInt(strYear) / 4 = CInt(strYear) \ 4 And CInt(strYear) \ 100 <> CInt(strYear) / 100 Then
                    ValidityDate = Format("29/" & strMonth & "/" & strYear, "dd/mm/yyyy")
                Else
                    ValidityDate = Format("28/" & strMonth & "/" & strYear, "dd/mm/yyyy")
                End If
            Case "04", "06", "09", "11"
                ValidityDate = Format("30/" & strMonth & "/" & strYear, "dd/mm/yyyy")
        End Select
        
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then
        Select Case strThreeMonths
            Case "1", "2", "3", "4"
                ValidityDate = GetNgayCuoiQuy(CInt(strThreeMonths), _
                            CInt(strYear), iNgayTaiChinh, iThangTaiChinh)
        End Select
    '*******************************************
    ' ThanhDX modified
    ' Date: 04/04/06
    ' ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Day") = "1" Then
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
    '*******************************************
       ValidityDate = NgayCuoiNamTaiChinh(CInt(strYear), iThangTaiChinh, iNgayTaiChinh)
    Else
        ValidityDate = Date
    End If
    
    xmlDomMenu.Load GetAbsolutePath("menu.xml")
    
    Set xmlNodeListMenu = xmlDomMenu.getElementsByTagName("Root").Item(0).childNodes
    For Each xmlNode In xmlNodeListMenu
        If Id = GetAttribute(xmlNode, "ID") Then
            Set xmlNodeListValidity = xmlNode.selectNodes("Validity")
            Exit For
        End If
    Next
    'Set xmlNodeListValidity = xmlDomMenu.selectNodes("Validity")
    'Set xmlNodeListValidity = TAX_Utilities_Svr_New.NodeMenu.selectNodes("Validity")
    For Each xmlNodeValidity In xmlNodeListValidity
        StartDate = Format(GetAttribute(xmlNodeValidity, "StartDate"), "dd/mm/yyyy")
        If ValidityDate >= StartDate Then
            If StartDate > MaxDate Then
                MaxDate = StartDate
                Set GetValidityNode = xmlNodeValidity
            End If
        End If
    Next
    
    Set xmlDomMenu = Nothing
    Set xmlNodeListMenu = Nothing
    Set xmlNodeListValidity = Nothing
    
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunctions", "GetValidityNode", Err.Number, Err.Description
End Function

Private Function GetNgayCuoiQuy(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
    Dim mTaiChinhDau As Integer
    Dim mTaiChinhCuoi As Integer
    Dim yTaiChinhDau As Integer
    Dim yTaiChinhCuoi As Integer
    Dim iInterval As Integer
    
    mTaiChinhDau = (q - 1) * 3 + dThangTaiChinh + 2 'Thang cuoi quy
    If dNgayTaiChinh = 1 Then
        mTaiChinhCuoi = mTaiChinhDau + 1 'Thang dau quy sau
        yTaiChinhDau = Y
        yTaiChinhCuoi = Y
        If mTaiChinhDau > 12 Then
            mTaiChinhDau = mTaiChinhDau - 12
            yTaiChinhDau = Y + 1
        End If
        If mTaiChinhCuoi > 12 Then
            mTaiChinhCuoi = mTaiChinhCuoi - 12
            yTaiChinhCuoi = Y + 1
        End If
        
        'Limitation of year
        If yTaiChinhCuoi >= 10000 Then
            yTaiChinhCuoi = 9999
        End If
        
        iInterval = DateDiff("D", DateSerial(yTaiChinhDau, mTaiChinhDau, 1), DateSerial(yTaiChinhCuoi, mTaiChinhCuoi, 1)) - 1
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1) + iInterval
    Else
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1)
    End If
End Function
Private Function NgayCuoiNamTaiChinh(Y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer) As Date
    Dim dNgayTC As Date
    
    dNgayTC = DateSerial(Y, dThangTaiChinh, dNgayTaiChinh)
    NgayCuoiNamTaiChinh = DateAdd("M", 12, dNgayTC)
    NgayCuoiNamTaiChinh = DateAdd("d", -1, NgayCuoiNamTaiChinh)
End Function

Private Function GetThangTaiChinh(strDate As String) As Integer
    Dim arrDateUnit() As String
    Dim i As Integer
    
    GetThangTaiChinh = -1
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        arrDateUnit(1) = Trim(arrDateUnit(1))
        GetThangTaiChinh = Val(arrDateUnit(1))
    End If
End Function
Private Function GetNgayTaiChinh(strDate As String) As Integer
    Dim arrDateUnit() As String
    Dim i As Integer
    
    GetNgayTaiChinh = -1
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        arrDateUnit(0) = Trim(arrDateUnit(0))
        GetNgayTaiChinh = Val(arrDateUnit(0))
    End If
End Function

Private Function MSTBoGach(ByVal strMST As String) As String
    Dim TEMP As String
    TEMP = strMST
    TEMP = Replace(TEMP, "-", "", 1)
    MSTBoGach = TEMP
End Function


Public Sub SetDateFormat(FpSpd As fpSpread, SheetNumber As Integer, RowNumber As Long, ColNumber As Long, strFormat As String)
    FpSpd.Sheet = SheetNumber
    FpSpd.Row = RowNumber
    FpSpd.Col = ColNumber
    FpSpd.CellType = CellTypePic
    ' Set the characters to center
    FpSpd.TypeHAlign = TypeHAlignCenter
    FpSpd.TypeVAlign = TypeHAlignCenter
    FpSpd.TypePicDefaultText = "../../...."
    
    Select Case LCase(strFormat)
        Case LCase(DDMMYYYY)
            FpSpd.TypePicMask = "99//99//9999"
        Case LCase(DDMM)
            FpSpd.TypePicMask = "99//99"
        Case LCase(MMYYYY)
            FpSpd.TypePicDefaultText = "../...."
            FpSpd.TypePicMask = "99//9999"
        Case LCase(YYYY)
            FpSpd.TypePicDefaultText = "...."
            FpSpd.TypePicMask = "9999"
    End Select
End Sub

Public Function Format_ddmmyyyy(str As String) As String
    Dim DD As String, MM As String, YYYY As String, dDate As Date
    
  If str <> "" Or Len(str) > 0 Then
    On Error GoTo e
    DD = Left(str, InStr(str, "/") - 1)
    MM = Mid(str, 4, 2)
    YYYY = Right("0000" & str, 4)
 
    
        If Val(DD) >= 1 And Val(DD) <= 31 Then
            DD = Format(DD, "0#")
        Else
            GoTo e
        End If
        
        If Val(MM) >= 1 And Val(MM) <= 12 Then
            MM = Format(MM, "0#")
        Else
            GoTo e
        End If
        
        If Val(YYYY) >= 0 And Val(YYYY) <= 9999 Then
            
            If Val(YYYY) >= 0 And Val(YYYY) <= 999 Then YYYY = CStr(2000 + Val(YYYY))
            If Val(YYYY) < 1900 Then GoTo e
            YYYY = Format(YYYY, "####")
        Else
            GoTo e
        End If
        
        dDate = Format(MM & "/" & DD & "/" & YYYY, "mm/dd/yyyy")
        'Format_ddmm = dd & "/" & mm
        Format_ddmmyyyy = DD & "/" & MM & "/" & YYYY
    End If
    Exit Function
e:
    DisplayMessage "0073", msOKOnly, miCriticalError
    Format_ddmmyyyy = ""
End Function


Public Function Format_mmyyyy(str As String) As String
    Dim m As String, Y As String
    
    On Error GoTo e
    m = Left(str, InStr(str, "/") - 1)
    Y = Right(str, Len(str) - InStr(str, "/"))
    Y = Replace(Y, ".", "")
    If IsNumeric(m) And IsNumeric(Y) Then
        If Val(m) >= 1 And Val(m) <= 12 Then
            Format_mmyyyy = Format(m, "0#")
        Else
            GoTo e
        End If
        
        If Val(Y) >= 0 And Val(Y) <= 9999 Then
            
            If Val(Y) >= 0 And Val(Y) <= 999 Then Y = CStr(2000 + Val(Y))
            If Val(Y) < 1900 Then GoTo e
            Format_mmyyyy = Format_mmyyyy & "/" & Format(Y, "####")
        Else
            GoTo e
        End If
    End If
    Exit Function
e:
    Format_mmyyyy = ""
End Function

' Ham sinh ra SQL HDR va DTL
' Ham de ghi du lieu vao bang trung gian de gui to khai len TC
Public Function InsertHDR_TGTC(ByRef hdr As TNCN_HDR) As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim sSQL As String
    
    Dim kykk_tu_ngay As Variant
    Dim kykk_den_ngay As Variant
    Dim kyLB_tu_ngay As Variant
    Dim kyLB_den_ngay As Variant
    Dim tKyKekhai As Variant
    Dim tKyLB As Variant
    Dim dDate As Date, strDate() As String
    
     'Ky/ Quy KK
    If hdr.kieu_ky = "M" Then
        'Ngay dau ky ke khai va ngay cuoi ky ke khai
        tKyKekhai = hdr.KYKKHAI
        tKyKekhai = Replace(tKyKekhai, "'", "")
        strDate = Split(tKyKekhai, "/")
        dDate = DateSerial(Int(strDate(1)), Int(strDate(0)), 1)
        kykk_tu_ngay = dDate
        kykk_den_ngay = DateAdd("m", 1, kykk_tu_ngay)
        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
        
        If Trim(kykk_den_ngay) = vbNullString Then
                kykk_den_ngay = "CTOD('')"
        Else
                kykk_den_ngay = "CTOD('" & Format(kykk_den_ngay, "mm/dd/yyyy") & "')"
        End If
        
        If Trim(kykk_tu_ngay) = vbNullString Then
                kykk_tu_ngay = "CTOD('')"
        Else
                kykk_tu_ngay = "CTOD('" & Format(kykk_tu_ngay, "mm/dd/yyyy") & "')"
        End If
    ElseIf hdr.kieu_ky = "Q" Then
        tKyKekhai = hdr.KYKKHAI
        tKyKekhai = Replace(tKyKekhai, "'", "")
        strDate = Split(tKyKekhai, "/")
        dDate = GetNgayDauQuy(Int(strDate(0)), Int(strDate(1)))
        kykk_tu_ngay = dDate
        kykk_den_ngay = DateAdd("m", 3, kykk_tu_ngay)
        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
        If Trim(kykk_tu_ngay) = vbNullString Then
                kykk_tu_ngay = "CTOD('')"
        Else
                kykk_tu_ngay = "CTOD('" & Format(kykk_tu_ngay, "mm/dd/yyyy") & "')"
        End If
        If Trim(kykk_den_ngay) = vbNullString Then
                kykk_den_ngay = "CTOD('')"
        Else
                kykk_den_ngay = "CTOD('" & Format(kykk_den_ngay, "mm/dd/yyyy") & "')"
        End If
    Else
        kykk_tu_ngay = "CTOD('')"
        kykk_den_ngay = "CTOD('')"
    End If
    
    'Ky lb
    If hdr.kieu_ky = "M" Or hdr.kieu_ky = "Q" Then
        'Ngay dau ky lap bo va ngay cuoi ky lap bo
        tKyLB = hdr.KYLBO
        tKyLB = Replace(tKyLB, "'", "")
        strDate = Split(tKyLB, "/")
        dDate = DateSerial(Int(strDate(1)), Int(strDate(0)), 1)
        kyLB_tu_ngay = dDate
        kyLB_den_ngay = DateAdd("m", 1, kyLB_tu_ngay)
        kyLB_den_ngay = DateAdd("d", -1, kyLB_den_ngay)
        If Trim(kyLB_den_ngay) = vbNullString Then
                kyLB_den_ngay = "CTOD('')"
        Else
                kyLB_den_ngay = "CTOD('" & Format(kyLB_den_ngay, "mm/dd/yyyy") & "')"
        End If
        
        If Trim(kyLB_tu_ngay) = vbNullString Then
                kyLB_tu_ngay = "CTOD('')"
        Else
                kyLB_tu_ngay = "CTOD('" & Format(kyLB_tu_ngay, "mm/dd/yyyy") & "')"
        End If
    ElseIf hdr.kieu_ky = "Q" Then
        
    Else
        kyLB_den_ngay = "CTOD('')"
        kyLB_tu_ngay = "CTOD('')"
    End If
    
    
    
   sSQLCol = "INSERT INTO tmp_tncn_hdr (id,tin, ten_dtnt, dia_chi, loai_tkhai, ngay_nop, kylb_tu_ng, kylb_den_n, kykk_tu_ng, kykk_den_n, ngay_cap_n,"
    sSQLCol = sSQLCol + " nguoi_cn, co_loi_dda, so_hieu_te, so_tt_tk, da_nhan, ghi_chu_lo, khoa_so, phong_xly, kkbs, tthtk, kylbo, kykkhai, ma_cqt, thueondinh,Ma_dl_thue,So_hd_dl,Ngay_hd_dl,Lan_bs,Loai_kykk) "
    sSQLCol = sSQLCol + " values ("

    sSQLVal = hdr.Id & "," & hdr.Tin & "," & hdr.ten_dtnt & "," & hdr.DIA_CHI & "," & hdr.loai_tkhai & "," & hdr.Ngay_nop & "," & kyLB_tu_ngay & "," & _
    kyLB_den_ngay & "," & kykk_tu_ngay & "," & kykk_den_ngay & "," & hdr.ngay_cap_nhat & "," & hdr.Nguoi_cn & "," & hdr.co_loi_dda & "," & _
    hdr.so_hieu_tep & "," & hdr.So_tt_tk & "," & hdr.DA_NHAN & "," & hdr.ghi_chu_loi & "," & hdr.khoa_so & "," & hdr.Phong_xly & "," & hdr.kkbs & "," & hdr.TTHTK & "," & hdr.KYLBO & "," & hdr.kykkhai & "," & hdr.MA_CQT & "," & hdr.thueondinh & "," & hdr.ma_dl_thue & "," & hdr.so_hd_dl & "," & hdr.ngay_hd_dl & "," & hdr.lan_bs & ",'" & hdr.loai_kykk & "'"
    
    
    sSQL = sSQLCol & sSQLVal & " )"
   
   InsertHDR_TGTC = sSQL
End Function
' end
' Ham de ghi du lieu vao bang trung gian de gui to khai len TC
Public Function InsertDTL_TGTC(ByRef dtl As TNCN_DTL) As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim sSQL As String
    
    sSQLCol = "INSERT INTO tmp_tncn_dtl (id, hdr_id,matkhai, madtnt, kylbo, kykkhai, tthtk, ngnop, cttn, giatri, danhan, lan_quet, ky_hieu, ma_cqt) "
    sSQLCol = sSQLCol + " values ("

    sSQLVal = dtl.Id & "," & dtl.hdr_id & "," & dtl.MATKHAI & "," & dtl.MADTNT & "," & dtl.KYLBO & "," & dtl.KYKKHAI & "," & dtl.TTHTK & "," & dtl.NGNOP & "," & _
    dtl.CTTN & "," & dtl.GIATRI & "," & dtl.DANHAN & "," & dtl.LAN_QUET & "," & dtl.ky_hieu & "," & dtl.ma_cqt
    
    sSQL = sSQLCol & sSQLVal & " )"
    
   InsertDTL_TGTC = sSQL
End Function
' end
' tinh ngay dau cua Quy
Public Function GetNgayDauQuy(q As Integer, Y As Integer) As Date
    Dim intYear As Integer, intDay As Integer, intMonth As Integer
    intDay = 1
    intYear = Y
    Select Case q
        Case 1
            intMonth = 1
        Case 2
            intMonth = 4
        Case 3
            intMonth = 7
        Case 4
            intMonth = 10
    End Select
    GetNgayDauQuy = DateSerial(intYear, intMonth, intDay)
End Function

' Get ve ID cua bang HDR
Public Function GetHdrId(strPath As String) As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim hdrId As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileNameHDR = strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_parm.DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
        End If
            sSQL = "SELECT prm_value FROM tmp_parm " & _
                " WHERE prm_name = 'rcv_xltk_hdr_seq' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 hdrId = rs.Fields("prm_value")
            Else
                 hdrId = 0
            End If
            sSQL = "update tmp_parm set prm_value ='" & Val(CStr(hdrId)) + 1 & "' where prm_name = 'rcv_xltk_hdr_seq' "
            clsConn.ExecuteDLL (sSQL)
            clsConn.Disconnect
    Else
        hdrId = 0
    End If
    GetHdrId = Val(Trim(CStr(hdrId)))
End Function

' Get ve ID cua bang DTL
Public Function GetDtlId(strPath As String) As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim dtlId As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileNameHDR = strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_parm.DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
        End If
            sSQL = "SELECT prm_value FROM tmp_parm " & _
                " WHERE prm_name = 'rcv_xltk_dtl_seq' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 dtlId = rs.Fields("prm_value")
            Else
                 dtlId = 0
            End If
            sSQL = "update tmp_parm set prm_value ='" & Val(CStr(dtlId)) + 1 & "' where prm_name = 'rcv_xltk_dtl_seq' "
            clsConn.ExecuteDLL (sSQL)
            clsConn.Disconnect
    Else
        dtlId = 0
    End If
    GetDtlId = Val(Trim(CStr(dtlId)))
End Function

' Thong bao AC
' Get ve ID cua bang HDR
'VTToan sua
Public Function GetHdrIdAC(strPath As String) As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim hdrId As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileNameHDR = strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_parm_ac.DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
        End If
            sSQL = "SELECT prm_value FROM tmp_parm_ac " & _
                " WHERE prm_name = 'rcv_xltk_hdr_ac' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 hdrId = rs.Fields("prm_value")
            Else
                 hdrId = 0
            End If
            sSQL = "update tmp_parm_ac set prm_value ='" & Val(CStr(hdrId)) + 1 & "' where prm_name = 'rcv_xltk_hdr_ac' "
            clsConn.ExecuteDLL (sSQL)
            clsConn.Disconnect
    Else
        hdrId = 0
    End If
    GetHdrIdAC = Val(Trim(CStr(hdrId)))
End Function

' Get ve ID cua bang DTL
Public Function GetDtlIdAC(strPath As String) As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim dtlId As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileNameHDR = strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_parm_ac.DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString strPath & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
        End If
            sSQL = "SELECT prm_value FROM tmp_parm_ac " & _
                " WHERE prm_name = 'rcv_xltk_dtl_ac' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 dtlId = rs.Fields("prm_value")
            Else
                 dtlId = 0
            End If
            sSQL = "update tmp_parm_ac set prm_value ='" & Val(CStr(dtlId)) + 1 & "' where prm_name = 'rcv_xltk_dtl_ac' "
            clsConn.ExecuteDLL (sSQL)
            clsConn.Disconnect
    Else
        dtlId = 0
    End If
    GetDtlIdAC = Val(Trim(CStr(dtlId)))
End Function

Public Function InsertHDR_TGTC_08(ByRef hdr As TNCN_HDR, kKKhaiTuNgay As Variant, kKKhaiDenNgay As Variant, KIEUQT As Variant) As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim sSQL As String

    Dim kykk_tu_ngay As Variant
    Dim kykk_den_ngay As Variant
    Dim kyLB_tu_ngay As Variant
    Dim kyLB_den_ngay As Variant
    Dim tKyKekhai As Variant
    Dim tKyLB As Variant
    Dim dDate As Date, strDate() As String
    Dim TU_NGAY As Variant, DEN_NGAY As Variant, NN_KD As Variant, TK_LAN_PS As Variant
    

     'Ky/ Quy KK
     'set TU_NGAY,DEN_NGAY,NN_KD,TK_LAN_PS
     TU_NGAY = "CTOD('')"
     DEN_NGAY = "CTOD('')"
     NN_KD = "''"
     TK_LAN_PS = 0
    If hdr.kieu_ky = "M" Then
        'Ngay dau ky ke khai va ngay cuoi ky ke khai
        tKyKekhai = hdr.KYKKHAI
        tKyKekhai = Replace(tKyKekhai, "'", "")
        strDate = Split(tKyKekhai, "/")
        dDate = DateSerial(Int(strDate(1)), Int(strDate(0)), 1)
        kykk_tu_ngay = kKKhaiTuNgay
        kykk_den_ngay = kKKhaiDenNgay
'        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
        
        If Trim(kykk_den_ngay) = vbNullString Then
                kykk_den_ngay = "CTOD('')"
        End If
        
        If Trim(kykk_tu_ngay) = vbNullString Then
                kykk_tu_ngay = "CTOD('')"
        End If
        'set lai gia tri TU_NGAY, DEN_NGAY, NN_KD
        TU_NGAY = kKKhaiTuNgay
        DEN_NGAY = kKKhaiDenNgay
        TK_LAN_PS = 1
        NN_KD = "'" & Trim(KIEUQT) & "'"
    ElseIf hdr.kieu_ky = "Q" Then
        tKyKekhai = hdr.KYKKHAI
        tKyKekhai = Replace(tKyKekhai, "'", "")
        strDate = Split(tKyKekhai, "/")
        dDate = GetNgayDauQuy(Int(strDate(0)), Int(strDate(1)))
        kykk_tu_ngay = dDate
        kykk_den_ngay = DateAdd("m", 3, kykk_tu_ngay)
        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
        If Trim(kykk_tu_ngay) = vbNullString Then
                kykk_tu_ngay = "CTOD('')"
        Else
                kykk_tu_ngay = "CTOD('" & Format(kykk_tu_ngay, "mm/dd/yyyy") & "')"
        End If
        If Trim(kykk_den_ngay) = vbNullString Then
                kykk_den_ngay = "CTOD('')"
        Else
                kykk_den_ngay = "CTOD('" & Format(kykk_den_ngay, "mm/dd/yyyy") & "')"
        End If
    Else
        kykk_tu_ngay = "CTOD('')"
        kykk_den_ngay = "CTOD('')"
    End If

    'Ky lb
    If hdr.kieu_ky = "M" Or hdr.kieu_ky = "Q" Then
        'Ngay dau ky lap bo va ngay cuoi ky lap bo
        tKyLB = hdr.KYLBO
        tKyLB = Replace(tKyLB, "'", "")
        strDate = Split(tKyLB, "/")
        dDate = DateSerial(Int(strDate(1)), Int(strDate(0)), 1)
        kyLB_tu_ngay = dDate
        kyLB_den_ngay = DateAdd("m", 1, kyLB_tu_ngay)
        kyLB_den_ngay = DateAdd("d", -1, kyLB_den_ngay)
        If Trim(kyLB_den_ngay) = vbNullString Then
                kyLB_den_ngay = "CTOD('')"
        Else
                kyLB_den_ngay = "CTOD('" & Format(kyLB_den_ngay, "mm/dd/yyyy") & "')"
        End If

        If Trim(kyLB_tu_ngay) = vbNullString Then
                kyLB_tu_ngay = "CTOD('')"
        Else
                kyLB_tu_ngay = "CTOD('" & Format(kyLB_tu_ngay, "mm/dd/yyyy") & "')"
        End If
    ElseIf hdr.kieu_ky = "Q" Then

    Else
        kyLB_den_ngay = "CTOD('')"
        kyLB_tu_ngay = "CTOD('')"
    End If



    sSQLCol = "INSERT INTO tmp_tncn_hdr (id,tin, ten_dtnt, dia_chi, loai_tkhai, ngay_nop, kylb_tu_ng, kylb_den_n, kykk_tu_ng, kykk_den_n, ngay_cap_n,"
    sSQLCol = sSQLCol + " nguoi_cn, co_loi_dda, so_hieu_te, so_tt_tk, da_nhan, ghi_chu_lo, khoa_so, phong_xly, kkbs, tthtk, kylbo, kykkhai, ma_cqt, thueondinh,Ma_dl_thue,So_hd_dl,Ngay_hd_dl,Lan_bs,TU_NGAY,DEN_NGAY,NN_KD,TK_LAN_PS) "
    sSQLCol = sSQLCol + " values ("

    sSQLVal = hdr.ID & "," & hdr.tin & "," & hdr.ten_dtnt & "," & hdr.DIA_CHI & "," & hdr.loai_tkhai & "," & hdr.ngay_nop & "," & kyLB_tu_ngay & "," & _
    kyLB_den_ngay & "," & kykk_tu_ngay & "," & kykk_den_ngay & "," & hdr.ngay_cap_nhat & "," & hdr.nguoi_cn & "," & hdr.co_loi_dda & "," & _
    hdr.so_hieu_tep & "," & hdr.so_tt_tk & "," & hdr.DA_NHAN & "," & hdr.ghi_chu_loi & "," & hdr.khoa_so & "," & hdr.phong_xly & "," & hdr.kkbs & "," & hdr.TTHTK & "," & hdr.KYLBO & "," & hdr.KYKKHAI & "," & hdr.ma_cqt & "," & hdr.thueondinh & "," & hdr.ma_dl_thue & "," & hdr.so_hd_dl & "," & hdr.ngay_hd_dl & "," & hdr.lan_bs & "," & TU_NGAY & "," & DEN_NGAY & "," & NN_KD & "," & TK_LAN_PS


    sSQL = sSQLCol & sSQLVal & " )"

   InsertHDR_TGTC_08 = sSQL
End Function

' end
' Ham de ghi du lieu vao bang trung gian de gui to khai len TC dung cho to TNCN 08
Public Function InsertDTL_TGTC08(ByRef dtl As TNCN_DTL, rowID As Integer) As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim sSQL As String
    
    If rowID = 0 Then
        sSQLCol = "INSERT INTO tmp_tncn_dtl_plus (id, hdr_id,matkhai, madtnt, kylbo, kykkhai, tthtk, ngnop, cttn, giatri, danhan, lan_quet, ky_hieu, ma_cqt) "
        sSQLCol = sSQLCol + " values ("
    
        sSQLVal = dtl.Id & "," & dtl.Hdr_id & "," & dtl.MATKHAI & "," & dtl.MADTNT & "," & dtl.KYLBO & "," & dtl.KYKKHAI & "," & dtl.TTHTK & "," & dtl.NGNOP & "," & _
        dtl.CTTN & "," & dtl.GIATRI & "," & dtl.DANHAN & "," & dtl.LAN_QUET & "," & dtl.ky_hieu & "," & dtl.MA_CQT
        
        sSQL = sSQLCol & sSQLVal & " )"
    Else
        sSQLCol = "INSERT INTO tmp_tncn_dtl_plus (id, hdr_id,matkhai, madtnt, kylbo, kykkhai, tthtk, ngnop, cttn, giatri, danhan, lan_quet, ky_hieu, ma_cqt,rowid) "
        sSQLCol = sSQLCol + " values ("
    
        sSQLVal = dtl.Id & "," & dtl.Hdr_id & "," & dtl.MATKHAI & "," & dtl.MADTNT & "," & dtl.KYLBO & "," & dtl.KYKKHAI & "," & dtl.TTHTK & "," & dtl.NGNOP & "," & _
        dtl.CTTN & "," & dtl.GIATRI & "," & dtl.DANHAN & "," & dtl.LAN_QUET & "," & dtl.ky_hieu & "," & dtl.MA_CQT & "," & rowID
        
        sSQL = sSQLCol & sSQLVal & " )"

    End If
    

    
   InsertDTL_TGTC08 = sSQL
End Function


