Attribute VB_Name = "mdlFunctions"
Option Explicit
Public Const DDMMYYYY = "DD/MM/YYYY"
Public Const DDMM = "DD/MM"
Public Const MMYYYY = "MM/YYYY"
Public Const YYYY = "YYYY"
Public Const DD = "DD"
Public Const MM = "MM"

Public Const SqlHdr_TNCN = "INSERT INTO RCV_TKHAI_HDR(ID,KKBS,TIN,TEN_DTNT,DIA_CHI,LOAI_TKHAI,NGAY_NOP,KYLB_TU_NGAY,KYLB_DEN_NGAY,KYKK_TU_NGAY,KYKK_DEN_NGAY,NGAY_CAP_NHAT,NGUOI_CAP_NHAT,CO_LOI_DDANH,SO_HIEU_TEP,SO_TT_TK,DA_NHAN,GHI_CHU_LOI,CO_GTRINH_02A,CO_GTRINH_02B,CO_GTRINH_02C, CO_PLUC_GTGT_01, CO_PLUC_GTGT_02, CO_PLUC_GTGT_03, TU_NGAY, DEN_NGAY, PHONG_XLY, CO_BANG_KE, CO_GHAN, TT_GUI,TIN_DLY,SO_HOP_DONG,NGAY_HOP_DONG,LAN_BS) VALUES("

Public Const SqlHdr_08TNCN_PIT = "INSERT INTO RCV_TKHAI_HDR(ID,KKBS,TIN,TEN_DTNT,DIA_CHI,LOAI_TKHAI,NGAY_NOP,KYLB_TU_NGAY,KYLB_DEN_NGAY,KYKK_TU_NGAY,KYKK_DEN_NGAY,NGAY_CAP_NHAT,NGUOI_CAP_NHAT,CO_LOI_DDANH,SO_HIEU_TEP,SO_TT_TK,DA_NHAN,GHI_CHU_LOI,CO_GTRINH_02A,CO_GTRINH_02B,CO_GTRINH_02C, CO_PLUC_GTGT_01, CO_PLUC_GTGT_02, CO_PLUC_GTGT_03, TU_NGAY, DEN_NGAY, PHONG_XLY, CO_BANG_KE, CO_GHAN, TT_GUI,TIN_DLY,SO_HOP_DONG,NGAY_HOP_DONG,LAN_BS,NGANH_NGHE_KD,TO_KHAI_LAN_PS) VALUES("

Public Const SqlHdr_08TNCN = "INSERT INTO RCV_TKHAI_HDR(ID,KKBS,TIN,TEN_DTNT,DIA_CHI,LOAI_TKHAI,NGAY_NOP,KYLB_TU_NGAY,KYLB_DEN_NGAY,KYKK_TU_NGAY,KYKK_DEN_NGAY,NGAY_CAP_NHAT,NGUOI_CAP_NHAT,CO_LOI_DDANH,SO_HIEU_TEP,SO_TT_TK,DA_NHAN,GHI_CHU_LOI,CO_GTRINH_02A,CO_GTRINH_02B,CO_GTRINH_02C, CO_PLUC_GTGT_01, CO_PLUC_GTGT_02, CO_PLUC_GTGT_03, TU_NGAY, DEN_NGAY, PHONG_XLY, CO_BANG_KE, CO_GHAN,TIN_DLY,SO_HOP_DONG,NGAY_HOP_DONG,LAN_BS,NGANH_NGHE_KD,TO_KHAI_LAN_PS) VALUES("

Public Const SqlHdr_AC = "INSERT INTO RCV_BCAO_HDR_AC (ID,TIN,LOAI_BC,NGAY_NOP,KYBC_TU_NGAY,KYBC_DEN_NGAY,NGAY_CAP_NHAT, NGUOI_CAP_NHAT, SO_TT_TK, DA_NHAN, PHONG_XLY, PHONG_QLY, CO_BANG_KE, HTHUC_NOP, ITKHAI_ID, TEN_DV_CQ,TIN_DV_CQ, NGAY_BC,NGUOI_DAI_DIEN,TEN_CQ_TIEP_NHAN,LY_DO_MAT, NGAY_MAT_HUY,PHUONG_PHAP_HUY,DUNG_DN_CQ, GHI_CHU, MA_CQT,LOAI_BC26,NGUOI_LAP_BIEU,QUY_BC) VALUES("
Public Const SqlHdr_BLP = "INSERT INTO RCV_BCAO_HDR_AC (ID,TIN,LOAI_BC,NGAY_NOP,KYBC_TU_NGAY,KYBC_DEN_NGAY,NGAY_CAP_NHAT, NGUOI_CAP_NHAT, SO_TT_TK, DA_NHAN, PHONG_XLY, PHONG_QLY, CO_BANG_KE, HTHUC_NOP, ITKHAI_ID, TEN_DV_CQ,TIN_DV_CQ, NGAY_BC,NGUOI_DAI_DIEN,TEN_CQ_TIEP_NHAN,LY_DO_MAT, NGAY_MAT_HUY,PHUONG_PHAP_HUY,DUNG_DN_CQ, GHI_CHU, MA_CQT,LOAI_BC26,NGUOI_LAP_BIEU,QUY_BC,NGAY_TB_PH) VALUES("


Public Function GetAttribute(xmlNodeCell As MSXML.IXMLDOMNode, pAttributeName As String) As String
    On Error Resume Next
    GetAttribute = Replace(xmlNodeCell.Attributes.getNamedItem(pAttributeName).nodeValue, "'", "''")
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
    
    Set xmlNodeCell = TAX_Utilities_Srv_New.Data(fps.Sheet - 1).nodeFromID(GetCellID(fps, pCol, pRow))
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
        Set xmlCatalogeValidNode = GetValidityNode("107", TAX_Utilities_Srv_New.Month, TAX_Utilities_Srv_New.ThreeMonths, TAX_Utilities_Srv_New.Year)

    
    'Get catalogue ID
    strCatalogueID = GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "CatalogueID")    'CatalogueID
    
    'Get catalogue pattern name
    strCatalogueName = GetCatalogueName(xmlCatalogeValidNode, strCatalogueID)
    
    GetCatalogueFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet - 1), "TemplateFolder") & _
       strCatalogueName & ".xml")
End Function

Public Function GetCatalogueFileNameTk_03TD_TAIN(Optional lSheet As Long = 1, Optional tkQuy As String = "0") As String
    Dim strCatalogueName     As String, strCatalogueID As String
    Dim xmlCatalogeValidNode As MSXML.IXMLDOMNode
    
    'Get valid catalogue node
    If tkQuy = "1" Then
        Set xmlCatalogeValidNode = GetValidityNode("107", TAX_Utilities_Srv_New.Month, TAX_Utilities_Srv_New.ThreeMonths, TAX_Utilities_Srv_New.Year)

    Else
        Set xmlCatalogeValidNode = GetValidityNode("107", TAX_Utilities_Srv_New.Month, "", TAX_Utilities_Srv_New.Year)

    End If
    
    'Get catalogue ID
    strCatalogueID = GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "CatalogueID")    'CatalogueID
    
    'Get catalogue pattern name
    strCatalogueName = GetCatalogueName(xmlCatalogeValidNode, strCatalogueID)
    
    GetCatalogueFileNameTk_03TD_TAIN = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet - 1), "TemplateFolder") & strCatalogueName & ".xml")
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
    
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1" Then
        strNgayTaiChinh = TAX_Utilities_Srv_New.FinanceStartDate
        iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
        iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
    Else
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
    End If
    
    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID") = "96" And strThreeMonths <> "" Then
        ValidityDate = GetNgayCuoiQuy(CInt(strThreeMonths), _
                            CInt(strYear), iNgayTaiChinh, iThangTaiChinh)
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
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
        
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
        Select Case strThreeMonths
            Case "1", "2", "3", "4"
                ValidityDate = GetNgayCuoiQuy(CInt(strThreeMonths), _
                            CInt(strYear), iNgayTaiChinh, iThangTaiChinh)
        End Select
    '*******************************************
    ' ThanhDX modified
    ' Date: 04/04/06
    ' ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Day") = "1" Then
    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
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
    'Set xmlNodeListValidity = TAX_Utilities_Srv_New.NodeMenu.selectNodes("Validity")
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


' Lay ve Row and Col
Public Sub ParserCellID(pGrid As fpSpread, pCellID As String, pCol As Long, pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim lPos As Long
    
    lPos = InStr(1, pCellID, "_", vbTextCompare)
    
    If lPos > 0 Then
        pCol = pGrid.ColLetterToNumber(Left(pCellID, lPos - 1))
        pRow = Val(Right(pCellID, Len(pCellID) - lPos))
    'ThanhDX added
    Else
        pCol = 0
        pRow = 0
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "ParserCellID", Err.Number, Err.Description
End Sub


Public Function GetMessageCellById(ByVal strId As String) As MSXML.IXMLDOMNode
    Dim xmlInforNode As MSXML.IXMLDOMNode
   
    For Each xmlInforNode In TAX_Utilities_Srv_New.NodeMessage
        If GetAttribute(xmlInforNode, "ID") = strId Then
            Set GetMessageCellById = xmlInforNode
            Exit Function
        End If
    Next
End Function
