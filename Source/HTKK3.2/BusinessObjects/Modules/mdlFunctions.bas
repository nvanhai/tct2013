Attribute VB_Name = "mdlFuntions"
Public Const DDMMYYYY = "DD/MM/YYYY"
Public Const DDMM = "DD/MM"
Public Const MMYYYY = "MM/YYYY"
Public Const yyyy = "YYYY"

Public Const mErrorColor = &HC0C0FF
Public Const mNonErrorColor = vbWhite

Public Const mFormColor = -2147483633
Public Const mLockNonErrorColor = 16709097

Public Const mAlertColor = 12713215

Public Const SxMST1Row = 2
Public Const SxMST1Col = "C"
Public Const SxMST2Row = 2
Public Const SxMST2Col = "D"
Public Const SxMST3Row = 2
Public Const SxMST3Col = "E"
Public Const SxMST4Row = 2
Public Const SxMST4Col = "F"
Public Const SxMST5Row = 2
Public Const SxMST5Col = "G"
Public Const SxMST6Row = 2
Public Const SxMST6Col = "H"
Public Const SxMST7Row = 2
Public Const SxMST7Col = "I"
Public Const SxMST8Row = 2
Public Const SxMST8Col = "J"
Public Const SxMST9Row = 2
Public Const SxMST9Col = "K"
Public Const SxMST10Row = 2
Public Const SxMST10Col = "L"
Public Const SxMST11Row = 2
Public Const SxMST11Col = "N"
Public Const SxMST12Row = 2
Public Const SxMST12Col = "O"
Public Const SxMST13Row = 2
Public Const SxMST13Col = "P"

Public Const Str_01BLP = "01BLP"
Public Const Str_02BLP = "02BLP"

Public strMauSoHD_01GTKT As Variant

Public mCurrentSheet As Integer
Public strFirstTimeRunID As String
Public arrErrorCells As Scripting.Dictionary

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
    SaveErrorLog "mdlFunctions", "GetCellSpan", Err.number, Err.Description
End Sub

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



'Public Sub GetCellSpan(fps As fpSpread, pCol As Long, pRow As Long, Optional pNumsRow As Variant, Optional pNumsCol As Variant)
'    Dim lRowAnchor As Variant, lColAnchor As Variant
'    Debug.Print fps.Sheet
'    fps.GetCellSpan pCol, pRow, lColAnchor, lRowAnchor, pNumsCol, pNumsRow
'    If lRowAnchor <> -1 And lColAnchor <> -1 Then
'        pRow = Val(lRowAnchor)
'        pCol = Val(lColAnchor)
'    End If
'
'End Sub



'format a month/year string as mm/yyyy
'if not able to format, out: vbnullstring
'if able, out a mm/yyyy string
Public Function Format_mmyyyy(str As String) As String
    Dim m As String, Y As String
    
    On Error GoTo e
    m = Left(str, InStr(str, "/") - 1)
    Y = Right(str, Len(str) - InStr(str, "/"))
    If IsNumeric(m) And IsNumeric(Val(Y)) Then
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

'format a day/month string as dd/mm
'if not able to format, out: vbnullstring
'if able, out a dd/mm string
Public Function Format_ddmm(str As String) As String
    Dim dd As String, mm As String, dDate As Date
    On Error GoTo e
    dd = Left(str, InStr(str, "/") - 1)
    mm = Right(str, Len(str) - InStr(str, "/"))
    If IsNumeric(dd) And IsNumeric(mm) Then
        If Val(dd) >= 1 And Val(dd) <= 31 Then
            dd = Format(dd, "0#")
        Else
            GoTo e
        End If
        
        If Val(mm) >= 1 And Val(mm) <= 12 Then
            mm = Format(mm, "0#")
        Else
            GoTo e
        End If
        dDate = Format(mm & "/" & dd & "/" & "2000", "mm/dd/yyyy")
        Format_ddmm = dd & "/" & mm
    End If
    Exit Function
e:
    Format_ddmm = ""
End Function

'format a day/month/year string as dd/mm/yyyy
'if not able to format, out: vbnullstring
'if able, out a dd/mm string
Public Function Format_ddmmyyyy(str As String) As String
    Dim dd As String, mm As String, yyyy As String, dDate As Date
    
  If str <> "" Or Len(str) > 0 Then
    On Error GoTo e
    dd = Left(str, InStr(str, "/") - 1)
    mm = Mid(str, 4, 2)
    yyyy = Right("0000" & str, 4)
 
    
        If Val(dd) >= 1 And Val(dd) <= 31 Then
            dd = Format(dd, "0#")
        Else
            GoTo e
        End If
        
        If Val(mm) >= 1 And Val(mm) <= 12 Then
            mm = Format(mm, "0#")
        Else
            GoTo e
        End If
        
        If Val(yyyy) >= 0 And Val(yyyy) <= 9999 Then
            
            If Val(yyyy) >= 0 And Val(yyyy) <= 999 Then yyyy = CStr(2000 + Val(yyyy))
            If Val(yyyy) < 1900 Then GoTo e
            yyyy = Format(yyyy, "####")
        Else
            GoTo e
        End If
        
        dDate = Format(mm & "/" & dd & "/" & yyyy, "mm/dd/yyyy")
        'Format_ddmm = dd & "/" & mm
        Format_ddmmyyyy = dd & "/" & mm & "/" & yyyy
    End If
    Exit Function
e:
    DisplayMessage "0071", msOKOnly, miCriticalError
    Format_ddmmyyyy = ""
End Function


Public Function Format_ddmmyyyy1(str As String) As String
    Dim dd As String, mm As String, yyyy As String, dDate As Date
    Dim arrDate() As String
  If str <> "" Or Len(str) > 0 Then
    On Error GoTo e
    arrDate = Split(str, "/")
    If UBound(arrDate) = 2 Then
'        dd = Left(str, InStr(str, "/") - 1)
'        mm = Mid(str, 4, 2)
'        yyyy = Right("0000" & str, 4)
        dd = arrDate(0)
        mm = arrDate(1)
        yyyy = arrDate(2)
    Else
        GoTo e
    End If
    
        If Val(dd) >= 1 And Val(dd) <= 31 Then
            dd = Format(dd, "0#")
        Else
            GoTo e
        End If
        
        If Val(mm) >= 1 And Val(mm) <= 12 Then
            mm = Format(mm, "0#")
        Else
            GoTo e
        End If
        
        If Val(yyyy) >= 0 And Val(yyyy) <= 9999 Then
            
            If Val(yyyy) >= 0 And Val(yyyy) <= 999 Then yyyy = CStr(2000 + Val(yyyy))
            If Val(yyyy) < 1900 Then GoTo e
            yyyy = Format(yyyy, "####")
        Else
            GoTo e
        End If
        
        dDate = Format(mm & "/" & dd & "/" & yyyy, "mm/dd/yyyy")
        'Format_ddmm = dd & "/" & mm
        Format_ddmmyyyy1 = dd & "/" & mm & "/" & yyyy
    End If
    Exit Function
e:
    Format_ddmmyyyy1 = ""
End Function

'format a day/month/year string as dd/mm/yyyy
'if not able to format, out: vbnullstring
'if able, out a dd/mm string
Public Function CheckFormat_ddmmyyyy(str As String) As String
    Dim dd As String, mm As String, yyyy As String, dDate As Date, nowDate As Date
    
  If str <> "" Or Len(str) > 0 Then
    On Error GoTo e
    dd = Left(str, InStr(str, "/") - 1)
    mm = Mid(str, 4, 2)
    yyyy = Right("0000" & str, 4)
 
    
 
    
        If Val(dd) >= 1 And Val(dd) <= 31 Then
            dd = Format(dd, "0#")
        Else
            GoTo e
        End If
        
        If Val(mm) >= 1 And Val(mm) <= 12 Then
            mm = Format(mm, "0#")
        Else
            GoTo e
        End If
        
        If Val(yyyy) >= 0 And Val(yyyy) <= 9999 Then
            
            If Val(yyyy) >= 0 And Val(yyyy) <= 999 Then yyyy = CStr(2000 + Val(yyyy))
            If Val(yyyy) < 1900 Then GoTo e
            yyyy = Format(yyyy, "####")
        Else
            GoTo e
        End If
        
        dDate = Format(mm & "/" & dd & "/" & yyyy, "mm/dd/yyyy")
        
        If DateSerial(yyyy, mm, dd) > DateSerial(Year(Date), Month(Date), Day(Date)) Then
            DisplayMessage "0135", msOKOnly, miCriticalError
            CheckFormat_ddmmyyyy = ""
            Exit Function
        End If
                
        'Format_ddmm = dd & "/" & mm
        CheckFormat_ddmmyyyy = dd & "/" & mm & "/" & yyyy
    End If
    Exit Function
e:
    DisplayMessage "0071", msOKOnly, miCriticalError
    CheckFormat_ddmmyyyy = ""
End Function


''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 fps      : fpspread that you want to handle
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Public Sub UpdateCell(fps As fpSpread, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_v1.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow))
    If Not xmlNodeCell Is Nothing Then
        SetAttribute xmlNodeCell, "Value", pValue
    End If
        
    Set xmlNodeCell = Nothing
End Sub
' 4/1/2010 dhdang them vao
''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 fps      : fpspread that you want to handle
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Public Sub UpdateCell_sheet(fps As fpSpread, ByVal pSheet As Long, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_v1.Data(pSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow))
    If Not xmlNodeCell Is Nothing Then
        SetAttribute xmlNodeCell, "Value", pValue
    End If
        
    Set xmlNodeCell = Nothing
End Sub

Public Sub UpdateKHBSCell(fps As fpSpread, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    GetCellSpan fps, pCol, pRow
    Set xmlNodeCell = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, pCol, pRow))
    
    If Not xmlNodeCell Is Nothing Then
        SetAttribute xmlNodeCell, "Value", pValue
    End If
    
    Set xmlNodeCell1 = Nothing
    Set xmlNodeCell = Nothing
End Sub

Public Sub UpdateLastKHBSCell(fps As fpSpread, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim xmlNodeCellData0 As MSXML.IXMLDOMNode
    Dim loaiKHBS11 As String
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCellData0 = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, pCol, pRow))
    
    If Not xmlNodeCellData0 Is Nothing Then
        SetAttribute xmlNodeCellData0, "Value", pValue
    End If
    
    Set xmlNodeCell = TAX_Utilities_v1.DataKHBS.nodeFromID(GetCellID(fps, pCol, pRow))
    loaiKHBS11 = GetAttribute(TAX_Utilities_v1.Data(TAX_Utilities_v1.NodeValidity.childNodes.length - 1).childNodes(2).firstChild, "loaiKHBS")
    If Not xmlNodeCell Is Nothing Then
        If loaiKHBS11 = "frmKHBS_TT" Then
             SetAttribute xmlNodeCell, "Value", pValue
        Else
             Dim varTemp As Variant
             varTemp = GetAttribute(xmlNodeCell, "Value")
             If IsNumeric(CStr(varTemp)) And IsNumeric(CStr(pValue)) Then
                SetAttribute xmlNodeCell, "Value", CStr(CDbl(pValue) + CDbl(varTemp))
             Else
                SetAttribute xmlNodeCell, "Value", Trim(varTemp)
             End If
        End If
             
    End If
   
    Set xmlNodeCell = Nothing
End Sub

Public Sub FormatTextNumber(fps As fpSpread, ByVal intSheet As Integer, ByVal lCol As Long, ByVal lRow As Long)
    fps.Sheet = intSheet
    fps.Col = lCol
    fps.Row = lRow
    fps.CellType = CellTypeEdit
    fps.TypeEditCharSet = TypeEditCharSetNumeric
    fps.TypeHAlign = TypeHAlignCenter
End Sub

Public Sub FormatText(fps As fpSpread, ByVal intSheet As Integer, ByVal lCol As Long, ByVal lRow As Long)
    fps.Sheet = intSheet
    fps.Col = lCol
    fps.Row = lRow
    fps.CellType = CellTypeEdit
    fps.TypeEditCharSet = TypeEditCharSetASCII
    fps.TypeHAlign = TypeHAlignRight
    fps.TypeVAlign = TypeHAlignCenter
End Sub


Public Sub FormatTextPercent(fps As fpSpread, ByVal intSheet As Integer, ByVal lCol As Long, ByVal lRow As Long, ByVal tfView As Boolean)
    Dim positionDecimalSymbol As Integer
    Dim tempValue As String
    Dim xmlNode As MSXML.IXMLDOMNode
    
    
    fps.Sheet = intSheet
    fps.Row = lRow
    fps.Col = lCol
    fps.CellType = CellTypeNumber
    ' Set the characters to right
    'sua loi sai fomat khi thay doi dau phan cach '.' va ','
'    If tfView Then fps.value = Val(fps.value) / 1000
    positionDecimalSymbol = 0
    If tfView Then
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "12" Then
            Set xmlNode = TAX_Utilities_v1.Data(0).nodeFromID("K_47")  'J_42: thue suat uu dai
            tempValue = GetAttribute(xmlNode, "Value")
            fps.value = tempValue
        ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "11" Then
            Set xmlNode = TAX_Utilities_v1.Data(0).nodeFromID("K_34")  'J_42: thue suat uu dai
            tempValue = GetAttribute(xmlNode, "Value")
            fps.value = tempValue
        End If
    End If
    
    fps.TypeHAlign = TypeHAlignRight
    fps.TypeVAlign = TypeHAlignCenter
    fps.TypeEditCharSet = TypeEditCharSetNumeric
    fps.TypeNumberMin = 0
    fps.TypeNumberMax = 100
    fps.TypeNumberDecimal = ","
    fps.TypeNumberDecPlaces = 3
    fps.TypePicDefaultText = "..,..."
    fps.TypePicMask = "99,999"
    
End Sub


''' GetCellID description
''' Get CellID of current cell
''' Parameter1 pGrid    : the current fpSpread grid (input value)
''' Parameter2 pCol     : the current column (input value)
''' Parameter3 pRow     : the current row (input value)
Public Function GetCellID(pGrid As fpSpread, ByVal pCol As Long, ByVal pRow As Long) As String
    GetCellID = pGrid.ColNumberToLetter(pCol) & "_" & CStr(pRow)
End Function

''' SetAttribute description
''' Set an attribute value to xmlNode
''' Parameter1 xmlNodeCell      : xmlNode the node need set attribute value
''' Parameter2 pAttributeName   : attribute name
''' Parameter3 pAttributeName   : attribute value
Public Sub SetAttribute(xmlNodeCell As MSXML.IXMLDOMNode, pAttributeName As String, pValue As String)
    On Error Resume Next
    xmlNodeCell.Attributes.getNamedItem(pAttributeName).nodeValue = pValue
End Sub


''' Compare Date Value
''' Parameter1 strMY (dd/mm/yyyy)     : xmlNode the node need set attribute value
''' out:
'''     true: equal
'''     false: not equal
Public Function MYCompare(StrMY As String) As Boolean
    Dim d() As String
    Dim MY1 As Date, MY2 As Date
    d = Split(StrMY, "/")
    On Error GoTo e
    MY1 = DateSerial(Val(d(2)), Val(d(1)), Val(d(0)))
    MY2 = DateSerial(TAX_Utilities_v1.Year, TAX_Utilities_v1.Month, 1)
    If MY1 <= MY2 Then
        MYCompare = True
    Else
        MYCompare = False
    End If
    
    Exit Function
e:
    'MsgBox "Unvalid Date String"
    MYCompare = False
End Function

Function IsEmail(Email As String) As Boolean
    Dim P1 As String, P2 As String, StrEmail As String
    Dim i As Integer, n As Integer
    StrEmail = Email
    IsEmail = True
    If InStr(StrEmail, "@") > 0 Then
        IsEmail = True
        
        
        
'        P1 = Left(StrEmail, InStr(StrEmail, "@") - 1)
'        n = Len(P1)
'        For i = 0 To n - 1
'            Select Case LCase(Left(P1, 1))
'                Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "x", "y", "z", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", "_", "-"
'
'                Case Else
'                    IsEmail = False
'                    Exit Function
'            End Select
'            P1 = Right(P1, Len(P1) - 1)
'        Next
'        P2 = Right(StrEmail, Len(StrEmail) - InStr(StrEmail, "@"))
'        If InStr(P2, ".") > 1 And InStr(P2, ".") < Len(P2) Then
'            IsEmail = True
'        Else
'            IsEmail = False
'        End If
    Else
        IsEmail = False
    End If
End Function

''' GetAttribute description
''' Get an attribute value of xmlNode
''' Parameter1 xmlNodeCell      : xmlNode the node need get attribute value
''' Parameter2 pAttributeName   : attribute name
''' Output                      : attribute value
Public Function GetAttribute(xmlNodeCell As MSXML.IXMLDOMNode, pAttributeName As String) As String
    On Error Resume Next
    GetAttribute = xmlNodeCell.Attributes.getNamedItem(pAttributeName).nodeValue
End Function

'''Check TaxCode if it's valid
'''Parameter: 13 number in Taxcode
'''Output   : a string with ### format with 0,1
'''0: rite ; 1: wrong
Public Function CheckTaxCode(ms1 As Variant, ms2 As Variant, ms3 As Variant, _
    ms4 As Variant, ms5 As Variant, ms6 As Variant, ms7 As Variant, _
    ms8 As Variant, ms9 As Variant, ms10 As Variant, ms11 As Variant, _
    ms12 As Variant, ms13 As Variant) As String
    Dim a As Long
    Dim i As Integer, ii As Integer, iii As Integer
    i = 0
    ii = 0
    iii = 0
    On Error GoTo ErrorHandle
    If CStr(ms1) = "" Or CStr(ms2) = "" Or CStr(ms3) = "" Or CStr(ms4) = "" _
            Or CStr(ms5) = "" Or CStr(ms6) = "" Or CStr(ms7) = "" Or CStr(ms8) = "" Or CStr(ms9) = "" Or CStr(ms10) = "" Then
            i = 1
        Else
            a = 31 * Val(ms1) + 29 * Val(ms2) + 23 * Val(ms3) + 19 * Val(ms4) + 17 * Val(ms5) + 13 * Val(ms6) + 7 * Val(ms7) + 5 * Val(ms8) + 3 * Val(ms9)
            If ms10 <> 10 - (a Mod 11) Then ii = 1
        End If

        If CStr(ms11) = "" And CStr(ms12) = "" And CStr(ms13) = "" Then
            iii = 0 'rite
        ElseIf CStr(ms11) = "" Or CStr(ms12) = "" Or CStr(ms13) = "" Then
            iii = 1 'not rite
        ElseIf CStr(ms11) <> "" And CStr(ms12) <> "" And CStr(ms13) <> "" Then
            iii = 0 'rite
        End If
        CheckTaxCode = Trim(str(i)) & Trim(str(ii)) & Trim(str(iii))
        Exit Function
ErrorHandle:
    CheckTaxCode = "000"
    SaveErrorLog "mdlFunctions", "CheckTaxCode", Err.number, Err.Description
End Function

Public Function IsNullValue(ByVal strValue As String) As Boolean
    Dim lCtrl As Long
    
    strValue = Replace(strValue, "0", "")
    strValue = Replace(strValue, ".", "")
    strValue = Replace(strValue, " ", "")
    
    If strValue = vbNullString Then
        IsNullValue = True
    Else
        IsNullValue = False
    End If
    
End Function
Public Function IsNullValue_ac(ByVal strValue As String) As Boolean
    Dim lCtrl As Long
    
    'strValue = Replace(strValue, "0", "")
    strValue = Replace(strValue, ".", "")
    strValue = Replace(strValue, " ", "")
    
    If strValue = vbNullString Then
        IsNullValue_ac = True
    Else
        IsNullValue_ac = False
    End If
    
End Function

''' ParserCellID description
''' Parser CellID string to column and row value
''' Parameter1 pGrid    : the current fpSpread grid
''' Parameter2 pCellID  : the CellID value of the xmlNode need parser
''' Parameter2 pCol     : Column value of cell (Output value)
''' Parameter2 pRow     : Row number of cell (Output value)
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
    SaveErrorLog "mdlFunctions", "ParserCellID", Err.number, Err.Description
End Sub

'Date format

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
        Case LCase(yyyy)
            FpSpd.TypePicDefaultText = "...."
            FpSpd.TypePicMask = "9999"
    End Select
End Sub

' Ham tinh so luong HD
' a-b: khoang
' a;b;c liet ke
' strResult = -1 loi
'Public Function GetSoLuongHD(ByVal str As String) As String
'    Dim strArr1() As String
'    Dim strArr2() As String
'    Dim strArr3() As String
'    Dim flag1 As Boolean
'    Dim flag2 As Boolean
'    Dim count As Integer
'    Dim strResult As String, j As Integer, Tuso1 As Integer, Tuso2 As Integer, denso1 As Integer, denso2 As Integer
'    Dim i As Integer
'    fla1 = False
'    fla1 = False
'
'    If InStr(1, str, ";") > 0 Then
'        flag1 = True  ' co danh sach liet ke
'    End If
'    If InStr(1, str, "-") > 0 Then
'        flag2 = True  ' co danh sach khoang
'    End If
'    ' khi nhap cac so chi liet ke
'    If flag1 = True And flag2 = False Then
'        strArr1 = Split(str, ";")
'        For i = 0 To UBound(strArr1)
'            If IsNullValue(strArr1(i)) Then
'                GetSoLuongHD = -1
'                Exit Function
'            End If
'        Next i
'        strResult = "" & UBound(strArr1) + 1
'    End If
'    ' khi nhap cac so chi co khoang
'    If flag1 = False And flag2 = True Then
'        strArr2 = Split(str, "-")
'        For j = 0 To UBound(strArr2)
'            If IsNullValue(strArr2(j)) Then
'                GetSoLuongHD = -1
'                Exit Function
'            End If
'        Next j
'        If UBound(strArr2) <> 1 Then
'            GetSoLuongHD = -1
'            Exit Function
'        Else
'            strResult = IIf(Val(Trim(strArr2(1))) - Val(Trim(strArr2(0))) < 0, -1, Val(Trim(strArr2(1))) - Val(Trim(strArr2(0))) + 1)
'        End If
'    End If
'    ' khi nhap so khong co ; va khong co -
'    If flag1 = False And flag2 = False Then
'        If Not IsNullValue(Trim(str)) Then
'            GetSoLuongHD = 1
'            Exit Function
'        End If
'    End If
'
'    ' khi nhap so co ca khoang va liet ke
'    If flag1 = True And flag2 = True Then
'        ' Lay ve mang cac khoang
'        strArr1 = Split(str, ";")
'        count = 0
'        For i = 0 To UBound(strArr1)
'            If IsNullValue(strArr1(i)) Then
'                GetSoLuongHD = -1
'                Exit Function
'            Else
'                ' Kiem tra dang a-b
'                strArr2 = Split(strArr1(i), "-")
'                ' kiem tr
'
'                If UBound(strArr2) <> 1 Then
'                    GetSoLuongHD = -1
'                    Exit Function
'                Else
'                    If IsNullValue(Trim(strArr2(1))) Or IsNullValue(Trim(strArr2(0))) Then
'                        GetSoLuongHD = -1
'                        Exit Function
'                    Else
'                        count = count + 1 ' dem cac khoang hop le
'                        If Val(Trim(strArr2(1))) <= Val(Trim(strArr2(0))) Then
'                            GetSoLuongHD = -1
'                            Exit Function
'                        Else
'                            strResult = Val(strResult) + Val(Trim(strArr2(1))) - Val(Trim(strArr2(0))) + 1
'                        End If
'                    End If
'                End If
'
'
'            End If
'        Next i
'        ' Kiem tra khoang giao nhau
'        If UBound(strArr1) + 1 = count Then
'            For i = 0 To UBound(strArr1)
'                strArr2 = Split(strArr1(i), "-")
'                Tuso1 = Val(Trim(strArr2(0)))
'                denso1 = Val(Trim(strArr2(1)))
'                For j = i + 1 To UBound(strArr1)
'                    strArr3 = Split(strArr1(j), "-")
'                    Tuso2 = Val(Trim(strArr3(0)))
'                    denso2 = Val(Trim(strArr3(1)))
'                     If (((Tuso1 - Tuso2) * (denso1 - Tuso2)) <= 0 Or ((Tuso1 - denso2) * (denso1 - denso2)) <= 0 Or ((Tuso1 > Tuso2) And (denso1 < denso2))) Then
'                        GetSoLuongHD = -1
'                        Exit Function
'                     End If
'                Next j
'            Next i
'        Else
'            ' co khoang khong hop le
'            GetSoLuongHD = -1
'            Exit Function
'        End If
'
'    End If
'
'
'    GetSoLuongHD = strResult
'End Function

' Ham tra ve gia tri cua tu so den so
' Ham tra ve -1 loi
' cac tham so truyen theo thu tu
Public Function getTusoDenso(Tuso1 As String, denso1 As String, Tuso2 As String, denso2 As String) As String()
    ' phan tu 0: tu so, -1 loi khoang khong lien tiep , -2 loi khac
    ' phan tu 1: den so
    ' phan tu 2: so luong
    ' phan tu 3: TH  1-> Cho co so ton dau ky, 2 -> chi co so phat hanh, 3 -> co ca ton va phat hanh
    ' phan tu 4: luu loi cua tung loai 1-> tso1 trang, 2-> dso1 trang, 3 -> tso1>dso1, 4 -> tso2 trang, 5-> dso2 trang, 6-> tso2>dso2
    ' phan tu 5: luu loi cua tung loai 1-> tso1 trang, 2-> dso1 trang, 3 -> tso1>dso1, 4 -> tso2 trang, 5-> dso2 trang, 6-> tso2>dso2
    Dim str(6) As String
    Dim tSo1 As Double, dSo1 As Double, tSo2 As Double, dSo2 As Double
    'dhdang sua loi nhap so 0
    'ngay 13-05
    Dim isEmty1 As Boolean
    If Tuso1 = "" Then
        isEmty1 = True
    End If
    
     Dim isEmty2 As Boolean
    If denso1 = "" Then
        isEmty2 = True
    End If
    
     Dim isEmty3 As Boolean
    If Tuso2 = "" Then
        isEmty3 = True
    End If
    
     Dim isEmty4 As Boolean
    If denso2 = "" Then
        isEmty4 = True
    End If
    
    tSo1 = Val(Tuso1)
    dSo1 = Val(denso1)
    tSo2 = Val(Tuso2)
    dSo2 = Val(denso2)
    ' kiem tra xem cac khoang thoa man dk tuso <= den so
    If (dSo1 - tSo1) < 0 Or (dSo2 - tSo2 < 0) Then
        str(0) = -1
        getTusoDenso = str
        Exit Function
    End If
    ' truong hop 1 chi co tu so 1, den so 1
    If (tSo1 >= 0 Or dSo1 >= 0) And Tuso2 = "" And denso2 = "" Then
        If tSo1 >= 0 And isEmty1 = False And dSo1 = 0 And isEmty2 = True Then
            str(0) = tSo1
            str(1) = ""
            str(1) = ""
            str(4) = "2"
        ElseIf tSo1 = 0 And isEmty1 = True And dSo1 >= 0 And isEmty2 = False Then
            str(0) = ""
            str(1) = dSo1
            str(1) = ""
            str(4) = "1"
        Else
            If dSo1 - tSo1 >= 0 And isEmty1 = False And isEmty2 = False Then
                str(0) = tSo1
                str(1) = dSo1
                str(2) = dSo1 - tSo1 + 1
                str(4) = ""
            Else
                str(0) = tSo1
                str(1) = dSo1
                str(2) = "0"
                str(4) = "3"
            End If
        End If
        str(3) = "1"  ' chi co so ton dau ky
    ElseIf (tSo2 >= 0 Or dSo2 >= 0) And Tuso1 = "" And denso1 = "" Then
    ' truong hop chi co tuso2, den so 2
        If tSo2 = 0 And isEmty3 = True And dSo2 >= 0 And isEmty4 = False Then
            str(0) = ""
            str(1) = dSo2
            str(2) = "0"
            str(4) = "4"
        ElseIf tSo2 >= 0 And isEmty3 = False And dSo2 = 0 And isEmty4 = True Then
            str(0) = tSo2
            str(1) = ""
            str(2) = "0"
            str(4) = "5"
        Else
            If dSo2 - tSo2 >= 0 Then
                str(0) = tSo2
                str(1) = dSo2
                str(2) = dSo2 - tSo2 + 1
                str(4) = ""
            Else
                str(0) = tSo2
                str(1) = dSo2
                str(2) = "0"
                str(4) = "6"
            End If
        End If
        str(3) = "2" ' chi co so phat hanh
    ElseIf (tSo1 >= 0 Or dSo1 >= 0) And (tSo2 >= 0 Or dSo2 >= 0) Then
    ' truong hop co ca 4 so
        ' kiem tra so ton
        If tSo1 >= 0 And isEmty1 = False And denso1 = "" And isEmty2 = True Then
            str(0) = tSo1
            str(1) = ""
            str(2) = "0"
            str(4) = "1"
        ElseIf tSo1 = 0 And isEmty1 = True And dSo1 >= 0 And isEmty2 = False Then
            str(0) = tSo1
            str(1) = ""
            str(2) = "0"
            str(4) = "2"
        Else
            If dSo1 - tSo1 >= 0 Then
                str(0) = tSo1
                str(1) = dSo1
                str(2) = dSo1 - tSo1 + 1
                str(4) = ""
            Else
                str(0) = tSo1
                str(1) = dSo1
                str(2) = "0"
                str(4) = "3"
            End If
        End If
        ' kiem tra so phat hanh
        If tSo1 >= 0 And isEmty1 = False And denso1 = "" And isEmty2 = True Then
            str(0) = tSo1
            str(1) = ""
            str(2) = "0"
            str(4) = "1"
        ElseIf Tuso1 = "" And isEmty1 = True And dSo1 >= 0 And isEmty2 = False Then
            str(0) = tSo1
            str(1) = ""
            str(2) = "0"
            str(4) = "2"
        Else
            If dSo1 - tSo1 >= 0 Then
                str(0) = tSo1
                str(1) = dSo1
                str(2) = dSo1 - tSo1 + 1
                str(4) = ""
            Else
                str(0) = tSo1
                str(1) = dSo1
                str(2) = "0"
                str(4) = "3"
            End If
        End If

        If tSo2 - dSo1 <> 1 Then
            str(0) = -1
        Else
            str(0) = tSo1
            str(1) = dSo2
            str(2) = dSo2 - tSo1 + 1
            str(3) = "3" ' co ca so ton va so phat hanh
        End If
    Else
        str(0) = -1
    End If
    getTusoDenso = str
End Function

' Ham check trung khoang giua cac dong trong sheet
' tra ve mang cac col, row bi trung
Public Function checkTrung(ByRef fps As fpSpread, startRow As Integer, charTuSo As String, charDenSo As String, charSTT As String, charTenHD As String, charMso As String, charKH As String, charMST As String, isMST As Boolean, tSheet As Integer) As String()
    ' Kiem tra trung nhau tu so den so
    ' bat dau
    Dim i As Integer
    Dim Tuso2 As Variant, denso2 As Variant, Tenloaihd2 As Variant, mauso2 As Variant, kyhieu2 As Variant
    Dim tSo2 As Double, dSo2 As Double, tSo1 As Double, dSo1 As Double
    Dim vMST2 As Variant, Tuso1 As Variant, denso1 As Variant
    Dim colTuSo As Integer, colDenSo As Integer, colSTT As Integer, colTenHD As Integer, colMso As Integer, colKH As Integer, colMST As Integer
    Dim checkMST As Boolean
    Dim result As String
    With fps
        .Sheet = tSheet
        colTuSo = .ColLetterToNumber(charTuSo)
        colDenSo = .ColLetterToNumber(charDenSo)
        colSTT = .ColLetterToNumber(charSTT)
        colTenHD = .ColLetterToNumber(charTenHD)
        colMso = .ColLetterToNumber(charMso)
        colKH = .ColLetterToNumber(charKH)
        If (isMST = True) Then
            colMST = .ColLetterToNumber(charMST)
        End If
        
        j = startRow
            Do
                .Row = j
                .GetText colTuSo, j, Tuso1
                .GetText colDenSo, j, denso1
                If isMST = True Then
                    .GetText colMST, j, vMST2
                End If
                tSo1 = Val(Tuso1)
                dSo1 = Val(denso1)
                .GetText colTenHD, j, Tenloaihd1
                .GetText colMso, j, mauso1
                .GetText colKH, j, kyhieu1
                
                Do
                    i = .Row + 1
                    .GetText colTuSo, i, Tuso2
                    .GetText colDenSo, i, denso2
                    tSo2 = Val(Tuso2)
                    dSo2 = Val(denso2)
                    .GetText colTenHD, i, Tenloaihd2
                    .GetText colMso, i, mauso2
                    .GetText colKH, i, kyhieu2
                    If isMST = True Then
                        .GetText colMST, .Row, vMST2
                    End If
                        ' kiem tra check MST
                        If isMST = False Then
                            checkMST = True
                        Else
                            If (vMST = vMST2) Then
                                checkMST = True
                            Else
                                checkMST = False
                            End If
                        End If
                        'dhdang sua loi check trung ca dong trang
                        'date 18-05-2011
                        If (((tSo1 - tSo2) * (dSo1 - tSo2)) <= 0 Or ((tSo1 - dSo2) * (dSo1 - dSo2)) <= 0 Or ((tSo1 > tSo2) And (dSo1 < dSo2))) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2) And checkMST = True) And Trim(Tuso2) <> "" And Trim(denso2) <> "" And Trim(Tuso1) <> "" And Trim(denso1) <> "" Then
                           ' to do
                           If InStr(1, result, j, vbTextCompare) = 0 Then
                                result = result & "~" & j
                           End If
                           If InStr(1, result, i, vbTextCompare) = 0 Then
                                result = result & "~" & i
                           End If
                        End If
                    .Row = i
                    .Col = colSTT
                    ' Check cho den het cac dong co du lieu thi thoi
                Loop Until UCase(.Text) = "AA"
            j = j + 1
            .Row = j
            .Col = colSTT
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
     End With
     If Len(result) > 2 Then
        result = Mid$(result, 2)
     End If
     checkTrung = Split(result, "~")
End Function
' Ham check ton dau ky PL2 BC26
' tra ve mang cac col, row ngoai khoang ton dau ky
Public Function checkTon(ByRef fps As fpSpread, startRow As Integer, charTuSo As String, charDenSo As String, charSTT As String, charTenHD As String, charMso As String, charKH As String, charMST As String, isMST As Boolean, tSheet As Integer) As String
    ' Kiem tra trung nhau tu so den so
    ' bat dau
    Dim i As Integer
    Dim Tuso2 As Variant, denso2 As Variant, Tenloaihd2 As Variant, mauso2 As Variant, kyhieu2 As Variant
    Dim tSo2 As Double, dSo2 As Double, tSo1 As Double, dSo1 As Double
    Dim vMST2 As Variant, Tuso1 As Variant, denso1 As Variant
    Dim colTuSo As Integer, colDenSo As Integer, colSTT As Integer, colTenHD As Integer, colMso As Integer, colKH As Integer, colMST As Integer
    Dim checkMST As Boolean
    Dim result As String
    With fps
        .Sheet = tSheet
        colTuSo = .ColLetterToNumber(charTuSo)
        colDenSo = .ColLetterToNumber(charDenSo)
        colSTT = .ColLetterToNumber(charSTT)
        colTenHD = .ColLetterToNumber(charTenHD)
        colMso = .ColLetterToNumber(charMso)
        colKH = .ColLetterToNumber(charKH)
        If (isMST = True) Then
            colMST = .ColLetterToNumber(charMST)
        End If
        
        j = startRow
            Do
                .Sheet = tSheet
                .Row = j
                .GetText colTuSo, j, Tuso1
                .GetText colDenSo, j, denso1
                If isMST = True Then
                    .GetText colMST, j, vMST2
                End If
                tSo1 = Val(Tuso1)
                dSo1 = Val(denso1)
                .GetText colTenHD, j, Tenloaihd1
                .GetText colMso, j, mauso1
                .GetText colKH, j, kyhieu1
                
                .Sheet = 1
                i = 22
                .Row = i
                Do
                    i = .Row
                    .GetText .ColLetterToNumber("X"), i, Tuso2
                    .GetText .ColLetterToNumber("Y"), i, denso2
                    tSo2 = Val(Tuso2)
                    dSo2 = Val(denso2)
                    .GetText .ColLetterToNumber("D"), i, Tenloaihd2
                    .GetText .ColLetterToNumber("E"), i, mauso2
                    .GetText .ColLetterToNumber("F"), i, kyhieu2
                    If isMST = True Then
                        .GetText colMST, .Row, vMST2
                    End If
                        ' kiem tra check MST
                        If isMST = False Then
                            checkMST = True
                        Else
                            If (vMST = vMST2) Then
                                checkMST = True
                            Else
                                checkMST = False
                            End If
                        End If
'                        If (Not ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2) And checkMST = True)) Then
'                                  If InStr(1, result, j, vbTextCompare) = 0 Then
'                                        result = result & "~" & j
'                                  End If
                        If (((tSo1 < tSo2) Or (dSo1 > dSo2) Or (tSo1 > dSo2) Or (dSo1 < tSo2)) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2) And checkMST = True)) Then
                                'If (((tSo1 >= tSo2) And (dSo1 <= dSo2)) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2) And checkMST = True)) Then
                                   ' to do
                                   If InStr(1, result, j, vbTextCompare) = 0 Then
                                        result = result & "~" & j
                                   End If
                                   If InStr(1, result, i, vbTextCompare) = 0 Then
                                        'result = result & "~" & j
                                   End If
                                'End If
                        End If
                    i = i + 1
                    .Row = i
                    .Col = colSTT
                    ' Check cho den het cac dong co du lieu thi thoi
                Loop Until UCase(.Text) = "AA"
            .Sheet = tSheet
            j = j + 1
            .Row = j
            .Col = colSTT
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
     End With
     If Len(result) > 2 Then
        result = Mid$(result, 2)
     End If
     checkTon = result
End Function
' tra ve mang cac col, row bi trung BC26 PL1
Public Function checkTrung_01(ByRef fps As fpSpread, startRow As Integer, charTuSo As String, charDenSo As String, charSTT As String, charTenHD As String, charMso As String, charKH As String, charMST As String, isMST As Boolean, tSheet As Integer) As String()
    ' Kiem tra trung nhau tu so den so
    ' bat dau
    Dim i As Integer
    Dim Tuso2 As Variant, denso2 As Variant, Tenloaihd2 As Variant, mauso2 As Variant, kyhieu2 As Variant
    Dim tSo2 As Double, dSo2 As Double, tSo1 As Double, dSo1 As Double
    Dim vMST2 As Variant, Tuso1 As Variant, denso1 As Variant
    Dim colTuSo As Integer, colDenSo As Integer, colSTT As Integer, colTenHD As Integer, colMso As Integer, colKH As Integer, colMST As Integer
    Dim checkMST As Boolean
    Dim result As String
    With fps
        .Sheet = tSheet
        colTuSo = .ColLetterToNumber(charTuSo)
        colDenSo = .ColLetterToNumber(charDenSo)
        colSTT = .ColLetterToNumber(charSTT)
        colTenHD = .ColLetterToNumber(charTenHD)
        colMso = .ColLetterToNumber(charMso)
        colKH = .ColLetterToNumber(charKH)
        If (isMST = True) Then
            colMST = .ColLetterToNumber(charMST)
        End If
        
        j = startRow
            Do
                .Row = j
                .GetText colTuSo, j, Tuso1
                .GetText colDenSo, j, denso1
                If isMST = True Then
                    .GetText colMST, j, vMST2
                End If
                tSo1 = Val(Tuso1)
                dSo1 = Val(denso1)
                .GetText colTenHD, j, Tenloaihd1
                .GetText colMso, j, mauso1
                .GetText colKH, j, kyhieu1
                
                Do
                    i = .Row + 1
                    .GetText colTuSo, i, Tuso2
                    .GetText colDenSo, i, denso2
                    tSo2 = Val(Tuso2)
                    dSo2 = Val(denso2)
                    .GetText colTenHD, i, Tenloaihd2
                    .GetText colMso, i, mauso2
                    .GetText colKH, i, kyhieu2
                    If isMST = True Then
                        .GetText colMST, .Row, vMST2
                    End If
                        ' kiem tra check MST
                        If isMST = False Then
                            checkMST = True
                        Else
                            If (vMST = vMST2) Then
                                checkMST = True
                            Else
                                checkMST = False
                            End If
                        End If
                        If (((tSo1 - tSo2) * (dSo1 - tSo2)) <= 0 Or ((tSo1 - dSo2) * (dSo1 - dSo2)) <= 0 Or ((tSo1 > tSo2) And (dSo1 < dSo2))) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2) And checkMST = True) And Tuso2 <> "" And denso2 <> "" Then
                           ' to do
                           If InStr(1, result, j, vbTextCompare) = 0 Then
                                result = result & "~" & j
                           End If
                           If InStr(1, result, i, vbTextCompare) = 0 Then
                                result = result & "~" & i
                           End If
                        End If
                    .Row = i
                    .Col = colSTT
                    ' Check cho den het cac dong co du lieu thi thoi
                Loop Until UCase(.Text) = "BB"
            j = j + 1
            .Row = j
            .Col = colSTT
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "BB"
     End With
     If Len(result) > 2 Then
        result = Mid$(result, 2)
     End If
     checkTrung_01 = Split(result, "~")
End Function


'ky hieu HD format
Public Sub SetKyHieuHDFormat(FpSpd As fpSpread, SheetNumber As Integer, RowNumber As Long, ColNumber As Long)
'    FpSpd.Sheet = SheetNumber
'    FpSpd.Row = RowNumber
'    FpSpd.Col = ColNumber
'    FpSpd.CellType = CellTypePic
    ' Set the characters to center
    FpSpd.TypeHAlign = TypeHAlignCenter
    FpSpd.TypeVAlign = TypeHAlignCenter
'    FpSpd.TypePicDefaultText = "../..."
'    FpSpd.TypePicMask = "xx//99x"
End Sub

' Dinh dang lai cau truc so HD tu -> den
Public Function FormatSoHD(str As String) As String
  Dim strTemp As String
  strTemp = str
  Dim i As Integer
  If str <> "" Or Len(str) > 0 Then
        For i = 1 To 7 - Len(strTemp)
            strTemp = "0" + strTemp
        Next
        FormatSoHD = strTemp
   End If
   Exit Function
End Function

' Ham format dinh dang ky hieu hoa don
Public Function FormatKyHieu(str As String) As String
  Dim strTemp As String
  strTemp = UCase(str)
  If Len(str) = 5 Then
      FormatKyHieu = Left$(strTemp, 2) & "/" & Right$(strTemp, 3)
  ElseIf Len(str) = 7 Then
        If IsNumeric(Left$(strTemp, 2)) Then
            FormatKyHieu = Left$(strTemp, 4) & "/" & Right$(strTemp, 3)
        Else
            FormatKyHieu = Left$(strTemp, 2) & "/" & Right$(strTemp, 5)
        End If
  Else
       FormatKyHieu = strTemp
  End If
End Function

' Ham format dinh dang ky hieu bien lai phi
Public Function FormatKyHieuBLP(str As String) As String
  Dim strTemp As String
  strTemp = UCase(str)
  If Len(str) = 5 Then
      FormatKyHieuBLP = Left$(strTemp, 2) & "-" & Right$(strTemp, 3)
  ElseIf Len(str) = 7 Then
        If IsNumeric(Left$(strTemp, 2)) Then
            FormatKyHieuBLP = Left$(strTemp, 4) & "-" & Right$(strTemp, 3)
        Else
            FormatKyHieuBLP = Left$(strTemp, 2) & "-" & Right$(strTemp, 5)
        End If
  Else
       FormatKyHieuBLP = strTemp
  End If
End Function

' Kiem tra cau truc ky hieu HD
' CheckSoHD = 1 sai cau truc
' CheckSoHD = 2 sai length
Public Function CheckSoHD(str As String, strLoai As Variant) As String
  Dim result As String
  Dim str1 As String
  Dim str2 As String
  Dim str3 As String
  Dim strTmpKH As String
  Dim strLoaiIn As String
  Dim strLoaiIn120 As String
  Dim i As Integer
  
  strLoaiIn = "ETP"
  strLoaiIn120 = "BNT"
  
  strLoaiInNNXC = "TP"
  
  strTmpKH = "ABCDEGHKLMNPQRSTUVXY"
  result = "0"
  ' kiem tra length 6 ky tu
  If strLoai = "3" Then
  Else
    If Len(Trim(str)) <> 6 And Len(Trim(str)) <> 8 Then
      result = "2"
      CheckSoHD = result
      Exit Function
    End If
  End If
  If strLoai = "3" Then
            ' 2 ky tu dau la ca ky tu chu cai "ABCDEGHKLMNPQRSTUVXY"
            ' ky tu cuoi thuoc chuoi "ETP"
            'dhdang sua truong hop 8 ky tu
            
            'TT39 khong bat cau truc ky hieu voi cac hoa don chon loai TT120
            
'            If Len(Trim(str)) = 8 Then
'                  str1 = Mid$(str, 4, 4)
'                  str2 = Right$(str, 3)
'                  str3 = Left$(str, 3)
'            End If
'            If InStr(strTmpKH, UCase(Left$(str3, 1))) > 0 And InStr(strTmpKH, UCase(Mid$(str3, 2, 1))) > 0 And IsNumeric(str1) = True Then
'               If InStr(str3, "/") > 0 Then
'                  If InStr(strLoaiIn120, UCase(Right$(str2, 1))) > 0 Then
'                       result = "0"
'                  Else
'                       result = "1"
'                  End If
'               Else
'                  result = "1"
'               End If
'            Else
'                 result = "1"
'            End If
  Else
            ' 2 ky tu dau la ca ky tu chu cai "ABCDEGHKLMNPQRSTUVXY"
            ' ky tu cuoi thuoc chuoi "ETP"
            'dhdang sua truong hop 8 ky tu
            If Len(Trim(str)) = 8 Then
                  str1 = Left$(str, 2)
                  str = Right$(str, 6)
            Else
                  str1 = "01"
            End If
            If InStr(strTmpKH, UCase(Left$(str, 1))) > 0 And InStr(strTmpKH, UCase(Mid$(str, 2, 1))) > 0 And IsNumeric(str1) = True And IsNumeric(Mid$(str, 4, 2)) = True And (0 < Val(str1)) And (Val(str1) < 65) Then
               If InStr(str, "/") > 0 Then
                  If InStr(strLoaiIn, UCase(Right$(str, 1))) > 0 Then
                       result = "0"
                  Else
                       result = "1"
                  End If
               Else
                  result = "1"
               End If
            Else
                'kiem tra ne 01GTKT thi cho phep nhap them NN/XC
                If Left$(Trim$(CStr(strMauSoHD_01GTKT)), 6) = "01GTKT" Then
                    If InStr(str, "/") > 0 And str1 = "NN" And Mid$(str, 2, 2) = "XC" Then
                       If InStr(strLoaiInNNXC, UCase(Right$(str, 1))) > 0 Then
                            result = "0"
                       Else
                            result = "1"
                       End If
                    Else
                       result = "1"
                    End If
                Else
                    result = "1"
                End If
            End If
    End If
  CheckSoHD = result
End Function


' Kiem tra cau truc ky hieu BLP
' CheckSoHD = 1 sai cau truc
' CheckSoHD = 2 sai length

Public Function CheckSoBLP(str As String, strMaLoaiBLP As Variant) As String
  Dim result As String
  Dim str1 As String
  Dim str2 As String
  Dim str3 As String
  Dim strTmpKH As String
  Dim strLoaiIn As String
  Dim i As Integer
  
  strLoaiIn = "TP"
  strTmpKH = "ABCDEGHKLMNPQRSTUVXY"
  result = "0"
  
  If strMaLoaiBLP = Str_01BLP Then
        ' kiem tra length 6 ky tu hoac 8 ky tu
        If Len(Trim(str)) <> 6 And Len(Trim(str)) <> 8 Then
          result = "2"
          CheckSoBLP = result
          Exit Function
        End If
  Else
       ' kiem tra length 6 ky tu
       If Len(Trim(str)) <> 6 Then
          result = "2"
          CheckSoBLP = result
          Exit Function
        End If
  End If

    ' 2 ky tu dau la ca ky tu chu cai "ABCDEGHKLMNPQRSTUVXY"
    ' ky tu thu 3 "-"
    ' 2 ky tu tiep theo la nam in BLP
    ' Ky tu cuoi la "TP"
    
    If Len(Trim(str)) = 8 Then
          str1 = Left$(str, 2)
          str = Right$(str, 6)
    Else
          str1 = "01"
    End If
    
    If strMaLoaiBLP = Str_01BLP Then
        ' check cau truc co do dai 6 hoac 8 ky tu
        If InStr(strTmpKH, UCase(Left$(str, 1))) > 0 And InStr(strTmpKH, UCase(Mid$(str, 2, 1))) > 0 And IsNumeric(str1) = True And (0 < Val(str1)) And (Val(str1) < 65) Then
           If InStr(str, "-") > 0 And IsNumeric(Mid$(str, 4, 2)) = True Then
              If InStr(strLoaiIn, UCase(Right$(str, 1))) > 0 Then
                   result = "0"
              Else
                   result = "1"
              End If
           Else
              result = "1"
           End If
        Else
             result = "1"
        End If
    Else
        If InStr(strTmpKH, UCase(Left$(str, 1))) > 0 And InStr(strTmpKH, UCase(Mid$(str, 2, 1))) > 0 Then
           If InStr(str, "-") > 0 And IsNumeric(Mid$(str, 4, 2)) = True Then
              If InStr(strLoaiIn, UCase(Right$(str, 1))) > 0 Then
                   result = "0"
              Else
                   result = "1"
              End If
           Else
              result = "1"
           End If
        Else
             result = "1"
        End If
    End If

  CheckSoBLP = result
End Function

'Kiem tra cau truc mau so
'str: chuoi mau so
'strLoai: 0,1,2,3
'strTemp: cac ky tu tien to cua loai HD
'CheckMauSoHD = 1 sai cau truc
'CheckMauSoHD = 2 sai so Lien


Public Function CheckMauSoHD(str As String, strLoai As String, strTemp As String) As String
    Dim result As String
    Dim soLien As String
    Dim kyTuNganCach As String
    Dim strSoTT As Variant
    Dim strBD As Variant
    result = "0"
    'str = Left$(Trim(str), 11)
    If strLoai = "0" Then
        If Len(Trim(str)) <> 11 And Len(Trim(str)) <> 13 Then
            result = "1"
            CheckMauSoHD = result
            Exit Function
        Else
            If Left$(Trim(str), 6) <> Trim(strTemp) Then
                result = "1"
                CheckMauSoHD = result
                Exit Function
            Else
                soLien = Mid$(Trim(str), 7, 1)
                kyTuNganCach = Mid$(Trim(str), 8, 1)
                strSoTT = Mid$(Trim(str), 9, 3)
                strBD = Mid$(Trim(str), 12, 2)
                ' so lien phai nam trong khoang 2->9
                If IsNumeric(soLien) Then
                    If strTemp = "01BHDT" Then
                        If Val(soLien) < 0 Or Val(soLien) > 9 Then
                            result = "2"
                            CheckMauSoHD = result
                            Exit Function
                        End If
                    Else
                        If Val(soLien) < 2 Or Val(soLien) > 9 Then
                            result = "2"
                            CheckMauSoHD = result
                            Exit Function
                        End If
                    End If
                Else
                    result = "2"
                    CheckMauSoHD = result
                    Exit Function
                End If
           
                ' ky tu so 8 phai la "/"
                If kyTuNganCach <> "/" Then
                    result = "1"
                    CheckMauSoHD = result
                    Exit Function
                End If
                ' 3 ky tu tiep theo la so thu tu cua HD
                If Not IsNumeric(strSoTT) Or Val(strSoTT) < 0 Then
                    result = "1"
                    CheckMauSoHD = result
                    Exit Function
                End If
                ' 2 ky tu cuoi hoac la rong hoac la BD hoac la IV
                If Trim(strBD) = "" Or Trim(strBD) = "BD" Or Trim(strBD) = "IV" Then
                    result = "0"
                Else
                    result = "1"
                    CheckMauSoHD = result
                    Exit Function
                End If
                
            End If
        End If
    ElseIf strLoai = "3" Then
        result = "0"
    ElseIf strLoai = "4" Then
        If UCase(Left$(Trim(str), 6)) <> Trim(strTemp) Then
              result = "1"   ' Loi khac voi chuoi ky tu chuan
              CheckMauSoHD = result
              Exit Function
        Else
              result = "0"
        End If
    Else
        ' 3 ky tu dau phai theo mau
        If Left$(Trim(str), 3) <> strTemp Then
            result = "1"
            CheckMauSoHD = result
            Exit Function
        End If
        ' kiem tra do dai khong duoc qua 11 ky tu
        If Len(Trim(str)) > 20 Or Len(Trim(str)) < 3 Then
            result = "1"
            CheckMauSoHD = result
            Exit Function
        End If
    End If
    CheckMauSoHD = result
End Function


'Kiem tra cau truc mau so
'str: chuoi mau so
'strLoai: 0,1,2,3
'strTemp: cac ky tu tien to cua loai HD
'CheckMauSoHD = 1 sai cau truc
'CheckMauSoHD = 2 sai so Lien

Public Function CheckMauSoBLP(str As String, strLoai As String, strTemp As String) As String
    Dim result As String
    Dim soLien As String
    Dim kyTuNganCach As String
    Dim strSoTT As Variant
    Dim strBD As Variant
    result = "0"
    'str = Left$(Trim(str), 11)
    If strLoai = "0" Then
        If Len(Trim(str)) <> 10 Then
            result = "1"
            CheckMauSoBLP = result
            Exit Function
        Else
            If Left$(Trim(str), 5) <> Trim(strTemp) Then
                result = "1"
                CheckMauSoBLP = result
                Exit Function
            Else
                soLien = Mid$(Trim(str), 6, 1)
                kyTuNganCach = Mid$(Trim(str), 7, 1)
                strSoTT = Mid$(Trim(str), 8, 3)
                ' so lien phai nam trong khoang 1->9
                If IsNumeric(soLien) Then
                    If Val(soLien) < 1 Or Val(soLien) > 9 Then
                        result = "2"
                        CheckMauSoBLP = result
                        Exit Function
                    End If
                Else
                    result = "2"
                    CheckMauSoBLP = result
                    Exit Function
                End If
                
                ' ky tu so 7 phai la "-"
                If kyTuNganCach <> "-" Then
                    result = "1"
                    CheckMauSoBLP = result
                    Exit Function
                End If
                ' 3 ky tu tiep theo la so thu tu cua HD
                If Not IsNumeric(strSoTT) Or Val(strSoTT) < 0 Then
                    result = "1"
                    CheckMauSoBLP = result
                    Exit Function
                End If
            End If
        End If
    End If
    CheckMauSoBLP = result
End Function

Public Sub ValidateDate(FpSpd As fpSpread, SheetNumber As Long, RowNumber As Long, ColNumber As Long, strFormat As String)
    FpSpd.Sheet = SheetNumber
    mCurrentSheet = SheetNumber
    FpSpd.Row = RowNumber
    FpSpd.Col = ColNumber
    If Trim(FpSpd.Text) = vbNullString Then
        FpSpd.TypePicDefaultText = ""
        UpdateCell FpSpd, FpSpd.Col, FpSpd.Row, FpSpd.Text
    Else
        If ToDate(FpSpd.Text, strFormat) <> "" Then
            FpSpd.Text = ToString(ToDate(FpSpd.Text, strFormat), strFormat)
        Else
            FpSpd.Text = ""
        End If
        UpdateCell FpSpd, FpSpd.Col, FpSpd.Row, FpSpd.Text
    End If
     Select Case LCase(strFormat)
        Case LCase(DDMMYYYY)
            FpSpd.TypePicDefaultText = "../../...."
            FpSpd.TypePicMask = "99//99//9999"
        Case LCase(DDMM)
            FpSpd.TypePicDefaultText = "../.."
            FpSpd.TypePicMask = "99//99"
        Case LCase(MMYYYY)
            FpSpd.TypePicDefaultText = "../...."
            FpSpd.TypePicMask = "99//9999"
        Case LCase(yyyy)
            FpSpd.TypePicDefaultText = "...."
            FpSpd.TypePicMask = "9999"
    End Select
End Sub

Public Sub ValidateDateError(FpSpd As fpSpread, SheetNumber As Long, RowNumber As Long, ColNumber As Long, strFormat As String)
    FpSpd.Sheet = SheetNumber
    mCurrentSheet = SheetNumber
    FpSpd.Row = RowNumber
    FpSpd.Col = ColNumber
    If Trim(FpSpd.Text) = vbNullString Then
        FpSpd.TypePicDefaultText = ""
        UpdateCell FpSpd, FpSpd.Col, FpSpd.Row, FpSpd.Text
    Else
        If ToDate(FpSpd.Text, strFormat) <> "" Then
            FpSpd.Text = ToString(ToDate(FpSpd.Text, strFormat), strFormat)
            FpSpd.CellNote = ""
            FpSpd.BackColor = mNonErrorColor
        Else
            arrErrorCells.Add SheetNumber & "_" & GetCellID(FpSpd, ColNumber, RowNumber), FpSpd.BackColor
            
            FpSpd.Text = ""
            'FpSpd.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            FpSpd.BackColor = mAlertColor
            Select Case LCase(strFormat)
                Case LCase(DDMMYYYY)
                    FpSpd.CellNote = GetAttribute(GetMessageCellById("0081"), "Msg")
                Case LCase(DDMM)
                    FpSpd.CellNote = GetAttribute(GetMessageCellById("0082"), "Msg")
                Case LCase(MMYYYY)
                    FpSpd.CellNote = GetAttribute(GetMessageCellById("0083"), "Msg")
                Case LCase(yyyy)
                    FpSpd.CellNote = GetAttribute(GetMessageCellById("0084"), "Msg")
            End Select
        End If
        UpdateCell FpSpd, FpSpd.Col, FpSpd.Row, FpSpd.Text
    End If
     Select Case LCase(strFormat)
        Case LCase(DDMMYYYY)
            FpSpd.TypePicDefaultText = "../../...."
            FpSpd.TypePicMask = "99//99//9999"
        Case LCase(DDMM)
            FpSpd.TypePicDefaultText = "../.."
            FpSpd.TypePicMask = "99//99"
        Case LCase(MMYYYY)
            FpSpd.TypePicDefaultText = "../...."
            FpSpd.TypePicMask = "99//9999"
        Case LCase(yyyy)
            FpSpd.TypePicDefaultText = "...."
            FpSpd.TypePicMask = "9999"
    End Select
End Sub


Public Sub ValidateNumberFormat(FpSpd As fpSpread, SheetNumber As Long, RowNumber As Long, ColNumber As Long)
    FpSpd.Sheet = SheetNumber
    mCurrentSheet = SheetNumber
    FpSpd.Row = RowNumber
    FpSpd.Col = ColNumber
    If Trim(FpSpd.Text) = vbNullString Then
        FpSpd.Text = ""
        UpdateCell FpSpd, FpSpd.Col, FpSpd.Row, FpSpd.Text
    Else
        FpSpd.Text = IsNumber(FpSpd.Text)
        UpdateCell FpSpd, FpSpd.Col, FpSpd.Row, FpSpd.Text
    End If
     
End Sub

Public Function IsNumber(strNumber As String) As String
    Dim strNumberReturn As String
    
    Dim d As String
    Dim m As String
    Dim Y As String
    Dim i As Integer
    
    If strNumber = vbNullString Then
        IsNumber = ""
        Exit Function
    End If
    
    strNumberReturn = ""
    If (InStr(1, strNumber, "-") <= 0) And (InStr(1, strNumber, ",") <= 0) And (InStr(1, strNumber, ".") <= 0) Then
        strNumberReturn = strNumber
    End If
    
    IsNumber = strNumberReturn
End Function
Private Function IsYear(strYear As String) As String
    Dim strYearReturn As String
    
    strYearReturn = ""
    strYear = Replace(strYear, ".", "", 1)
    If IsNumeric(strYear) And Val(strYear) >= 0 And Val(strYear) <= 9999 Then
        strYearReturn = strYear
        If Val(strYearReturn) <= 9 Then
            strYearReturn = CStr(2000 + Val(strYearReturn))
        End If
        If Val(strYearReturn) <= 99 Then
            strYearReturn = CStr(2000 + Val(strYearReturn))
        End If
        If Val(strYearReturn) <= 999 Then
            strYearReturn = CStr(2000 + Val(strYearReturn))
        End If
    End If

    If Val(strYearReturn) > 1900 Then
        IsYear = strYearReturn
    Else
        IsYear = ""
    End If
End Function

Private Function IsDate(Y As Integer, m As Integer, d As Integer) As Variant
    Dim dDateTemp As Date
    
    IsDate = ""
    
    dDateTemp = DateSerial(Y, m, d)
    If d = Day(dDateTemp) And m = Month(dDateTemp) And Y = Year(dDateTemp) Then
        IsDate = dDateTemp
    End If
End Function

Public Function ToDate(strDate As String, strFormat As String) As Variant
    Dim arrDateUnit() As String
    Dim d As Integer
    Dim m As Integer
    Dim Y As Integer
    Dim i As Integer
    
    ToDate = ""
    strDate = Replace(strDate, ".", "", 1)
    arrDateUnit = Split(strDate, "/")
    For i = 0 To UBound(arrDateUnit)
        arrDateUnit(i) = Trim(arrDateUnit(i))
    Next
    
    Select Case LCase(strFormat)
        Case LCase(DDMMYYYY)
            If UBound(arrDateUnit) = 2 Then
                d = Val(arrDateUnit(0))
                m = Val(arrDateUnit(1))
                If IsYear(arrDateUnit(2)) <> "" Then
                    Y = Val(IsYear(arrDateUnit(2)))
                    ToDate = IsDate(Y, m, d)
                End If
            End If
        Case LCase(DDMM)
            If UBound(arrDateUnit) = 1 Then
                d = Val(arrDateUnit(0))
                m = Val(arrDateUnit(1))
                Y = Year(Now)
                ToDate = IsDate(Y, m, d)
            End If
        Case LCase(MMYYYY)
            If UBound(arrDateUnit) = 1 Then
                d = 1
                m = Val(arrDateUnit(0))
                If IsYear(arrDateUnit(1)) <> "" Then
                    Y = Val(IsYear(arrDateUnit(1)))
                    ToDate = IsDate(Y, m, d)
                End If
            End If
        Case LCase(yyyy)
            If UBound(arrDateUnit) = 0 Then
                d = 1
                m = 1
                If IsYear(arrDateUnit(0)) <> "" Then
                    Y = Val(IsYear(arrDateUnit(0)))
                    ToDate = IsDate(Y, m, d)
                End If
            End If
        Case LCase(dd)
            If UBound(arrDateUnit) = 0 Then
                d = Val(arrDateUnit(0))
                m = Month(Now)
                Y = Year(Now)
                ToDate = IsDate(Y, m, d)
            End If
        Case LCase(mm)
            If UBound(arrDateUnit) = 0 Then
                d = 1
                m = Val(arrDateUnit(0))
                Y = Year(Now)
                ToDate = IsDate(Y, m, d)
            End If
    End Select
End Function

Public Function ToString(dDate As Variant, strFormat) As Variant
    Dim dUnit As String
    Dim mUnit As String
    Dim yUnit As String
    
    ToString = ""
    dUnit = Trim(str(Day(dDate)))
    mUnit = Trim(str(Month(dDate)))
    yUnit = Trim(str(Year(dDate)))
    
    If Len(dUnit) = 1 Then
        dUnit = "0" & dUnit
    End If
    If Len(mUnit) = 1 Then
        mUnit = "0" & mUnit
    End If
    
    Select Case LCase(strFormat)
        Case LCase(DDMMYYYY)
            ToString = dUnit & "/" & mUnit & "/" & yUnit
        Case LCase(DDMM)
            ToString = dUnit & "/" & mUnit
        Case LCase(MMYYYY)
            ToString = mUnit & "/" & yUnit
        Case LCase(yyyy)
            ToString = yUnit
        Case LCase(dd)
            ToString = dUnit
        Case LCase(mm)
            ToString = mUnit
    End Select
End Function

Public Function GetMessageCellById(ByVal strId As String) As MSXML.IXMLDOMNode
    Dim xmlInforNode As MSXML.IXMLDOMNode
    
    For Each xmlInforNode In TAX_Utilities_v1.NodeMessage
        If GetAttribute(xmlInforNode, "ID") = strId Then
            Set GetMessageCellById = xmlInforNode
            Exit Function
        End If
    Next
End Function

Public Function CPab(ByVal str As String, ByVal number As Integer) As String
    CPab = str & Space(number - Len(str))
End Function

Public Function GetCatalogueFileName(ByRef blnFollowCheck As Boolean, Optional lSheet As Long = 1) As String
    Dim strReturn As String
    Dim dValidDate As Date
    Dim strCatalogueName As String, strCatalogueID As String
    Dim fso As New FileSystemObject
    Dim xmlCatalogeValidNode As MSXML.IXMLDOMNode
    
    blnFollowCheck = False

    'Get valid catalogue node
    Set xmlCatalogeValidNode = GetValidityNode("100_7", TAX_Utilities_v1.Month, TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year)
    
    If Not xmlCatalogeValidNode.nextSibling Is Nothing Then
        dValidDate = Format(GetAttribute(xmlCatalogeValidNode.nextSibling, "StartDate"), "DD/MM/YYYY")
    End If
    
    'Kiem tra ky hieu luc hien hanh
    If xmlCatalogeValidNode.nextSibling Is Nothing Then
        blnFollowCheck = True
    ElseIf dValidDate > Date Then
        blnFollowCheck = True
    End If
    'Get catalogue ID
    strCatalogueID = GetAttribute(TAX_Utilities_v1.NodeValidity, "CatalogueID")
    'Get catalogue pattern name
    strCatalogueName = GetCatalogueName(xmlCatalogeValidNode, strCatalogueID)
    
    'Get catalogue file name
    strReturn = GetAbsolutePath(TAX_Utilities_v1.DataFolder & _
        strCatalogueName & ".xml")
    If fso.FileExists(strReturn) Then
        GetCatalogueFileName = strReturn
        Set fso = Nothing
        Exit Function
    End If
    
    'Get catalogue template file name
    strReturn = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(lSheet - 1), "TemplateFolder") & _
                strCatalogueName & ".xml")

    GetCatalogueFileName = strReturn
    
    Set fso = Nothing
    
End Function

Public Function GetCatalogueTemplateFile(Optional lSheet As Long = 1) As String
    Dim strReturn As String
    Dim strCatalogueName As String, strCatalogueID As String
    Dim xmlCatalogeValidNode As MSXML.IXMLDOMNode
    
    'Get valid catalogue node
    Set xmlCatalogeValidNode = GetValidityNode("100_7", TAX_Utilities_v1.Month, TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year)
       
    'Get catalogue ID
    strCatalogueID = GetAttribute(TAX_Utilities_v1.NodeValidity, "CatalogueID")
    
    'Get catalogue pattern name
    strCatalogueName = GetCatalogueName(xmlCatalogeValidNode, strCatalogueID)
    
    strReturn = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(lSheet - 1), "TemplateFolder") & _
        strCatalogueName & ".xml"
    GetCatalogueTemplateFile = GetAbsolutePath(strReturn)
End Function
Public Function GetValidityNode(ID As String, Optional strMonth As String, Optional strThreeMonths As String, Optional strYear As String) As MSXML.IXMLDOMNode
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
    
    If GetAttribute(TAX_Utilities_v1.NodeMenu, "FinanceYear") = "1" Then
        strNgayTaiChinh = GetNgayBatDauNamTaiChinh
        iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
        iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
    Else
        iNgayTaiChinh = 1
        iThangTaiChinh = 1
    End If
    
    If strMonth <> "" Then
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
        
    ElseIf strThreeMonths <> "" Then
        Select Case strThreeMonths
            Case "1", "2", "3", "4"
                ValidityDate = GetNgayCuoiQuy(CInt(strThreeMonths), _
                            CInt(strYear), iNgayTaiChinh, iThangTaiChinh)
        End Select
    '*******************************************
    ' ThanhDX modified
    ' Date: 04/04/06
    ' ElseIf GetAttribute(TAX_Utilities_v1.NodeMenu, "Day") = "1" Then
    ElseIf strYear <> "" Then
    '*******************************************
       ValidityDate = NgayCuoiNamTaiChinh(CInt(strYear), iThangTaiChinh, iNgayTaiChinh)
    Else
        ValidityDate = Date
    End If
    
    xmlDomMenu.Load GetAbsolutePath("menu.xml")
    
    Set xmlNodeListMenu = xmlDomMenu.getElementsByTagName("Root").Item(0).childNodes
    For Each xmlNode In xmlNodeListMenu
        If ID = GetAttribute(xmlNode, "ID") Then
            Set xmlNodeListValidity = xmlNode.selectNodes("Validity")
            Exit For
        End If
    Next
    'Set xmlNodeListValidity = xmlDomMenu.selectNodes("Validity")
    'Set xmlNodeListValidity = TAX_Utilities_v1.NodeMenu.selectNodes("Validity")
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
    SaveErrorLog "mdlFunctions", "GetValidityNode", Err.number, Err.Description
End Function
Public Function GetNgayCuoiQuy(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
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
Public Function NgayCuoiNamTaiChinh(Y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer) As Date
    Dim dNgayTC As Date
    
    dNgayTC = DateSerial(Y, dThangTaiChinh, dNgayTaiChinh)
    NgayCuoiNamTaiChinh = DateAdd("M", 12, dNgayTC)
    NgayCuoiNamTaiChinh = DateAdd("d", -1, NgayCuoiNamTaiChinh)
End Function



Public Function GetNgayBatDauNamTaiChinh() As String
    
    Dim xmlDomHeader As New MSXML.DOMDocument
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & "Header_01.xml")
        GetNgayBatDauNamTaiChinh = GetAttribute(xmlDomHeader.getElementsByTagName("Cell")(23), "Value")
    Set xmlDomHeader = Nothing
End Function
Public Function GetThangTaiChinh(strDate As String) As Integer
    Dim arrDateUnit() As String
    Dim i As Integer
    
    GetThangTaiChinh = -1
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        arrDateUnit(1) = Trim(arrDateUnit(1))
        GetThangTaiChinh = Val(arrDateUnit(1))
    End If
End Function
Public Function GetNgayTaiChinh(strDate As String) As Integer
    Dim arrDateUnit() As String
    Dim i As Integer
    
    GetNgayTaiChinh = -1
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        arrDateUnit(0) = Trim(arrDateUnit(0))
        GetNgayTaiChinh = Val(arrDateUnit(0))
    End If
End Function

Public Static Function CRound( _
    ByVal dblNumber As Double, _
    Optional ByVal numDecimalPlaces As Long = 0 _
  ) As Double
' by Donald, donald@xbeat.net, 20020419
' modification of Round10 inspired by Jost's Round13 (Variant = CDec!)
  Dim fInit As Boolean
  Dim numDecimalPlacesPrev As Long
  Dim vFac As Variant
  Dim dFacInv As Double
  
  ' calc factor once for this depth of rounding
  If Not fInit Or numDecimalPlacesPrev <> numDecimalPlaces Then
    vFac = CDec(10 ^ numDecimalPlaces)
    dFacInv = 10 ^ -numDecimalPlaces
    numDecimalPlacesPrev = numDecimalPlaces
    fInit = True
  End If
  
  On Error GoTo Err_CRound

  If dblNumber > 0 Then
    CRound = Int(dblNumber * vFac + 0.5)
    CRound = CRound * dFacInv
  Else
    CRound = -Int(-dblNumber * vFac + 0.5)
    CRound = CRound * dFacInv
  End If
  
  Exit Function
  
Err_CRound:
  CRound = dblNumber
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

Public Function cReadNum(ByVal Num As Double) As String
  Dim i As Byte
  Dim Hang, Donvi
  Dim Sochia As Double
  Dim luu As Double, t1 As Byte, t2 As Byte, t3 As Byte
  Dim St As String
  Dim result As Double
  Hang = Array("", " mét", " hai", " ba", " bèn", " n¨m", " s¸u", " b¶y", " t¸m", " chÝn")
  Donvi = Array(" tû", " triÖu", " ngh×n", "")
  Sochia = 1000000000
  St = "" 'Thirdpart number
  For i = 0 To 3
        luu = Num
        luu = luu \ Sochia Mod 1000
        t1 = luu \ 100
        t2 = luu \ 10 Mod 10
        t3 = luu Mod 10
        If t1 = 0 And t2 = 0 And t3 = 0 Then GoTo SKIP
            If t1 = 0 And St <> "" Then St = St & " kh«ng"
                    St = St & Hang(t1)
            If St <> "" Then St = St & " tr¨m"
            If t2 <> 0 Then
            If t2 = 1 Then
                    St = St & " m­êi"
            If t3 <> 5 Then St = St & Hang(t3)
            If t3 = 5 Then St = St & " l¨m"
            Else
                St = St & Hang(t2) & " m­¬i"
                If t3 <> 1 And t3 <> 5 Then St = St & Hang(t3)
                If t3 = 1 And t2 > 1 Then St = St & " mèt"
                If t3 = 5 Then St = St & " l¨m"
            End If
        Else
            If t3 <> 0 And St <> "" Then St = St & " linh"
            St = St & Hang(t3)
            End If
            If St <> "" Then St = St & Donvi(i)
SKIP:
        Sochia = Sochia \ 1000
Next
        St = Trim(St)
        If St <> "" Then
                St = St
        Else
                St = "kh«ng !"
        End If
        cReadNum = UCase(Left(St, 1)) & Right(St, Len(St) - 1)
        cReadNum = cReadNum & " ®ång ch½n"
End Function
Public Function GetQuyNamTaiChinh(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer, dType As Integer) As Integer
   ' q Quy ke khai
   ' y nam ke khai
   ' dNgayTaiChinh ngay tai chinh lay tren man hinh HTKK
   ' dThangTaiChinh thang tai chinh tren phan thong tin chung HTKK
   ' dType: 0 tra ve quy, 1 tra ve nam
    Dim intYear As Integer, intDay As Integer, intMonth As Integer, result As Integer
   
    intDay = dNgayTaiChinh
    intMonth = (q - 1) * 3 + dThangTaiChinh
    intYear = Y
    If intMonth > 12 Then
        intMonth = intMonth - 12
        intYear = Y + 1
    End If
    If dType = 0 Then
       result = DatePart("Q", DateSerial(intYear, intMonth, intDay))
    Else
       result = Year(DateSerial(intYear, intMonth, intDay))
    End If
    GetQuyNamTaiChinh = result
End Function
'******************************
'Set MST tren phan TT chung
Sub UpdateMST(fps As fpSpread, ByVal Col As String, ByVal Row As Long) 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim MSTDN As Variant
    Dim xmlDomHeader As New MSXML.DOMDocument
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & "Header_01.xml")
        MSTDN = GetAttribute(xmlDomHeader.getElementsByTagName("Cell")(32), "Value")
    Set xmlDomHeader = Nothing
    fps.Sheet = 1
    fps.SetText fps.ColLetterToNumber(Col), Row, MSTDN
    'UpdateCell_sheet fps, 1, fps.ColLetterToNumber("D"), 7, mstDN
End Sub

'kiem tra neu nguoi su dung check co ke khai ma so thue dl thi cac thong tin mstdl, ten mstdl, diachi mstdl, quan huyen, thanh pho,.... khong dc de trong

Public Sub checkErrorHeader(pGrid As fpSpread, chuoiTTHeader As String, viTriSetErr As String, chuoiMessage As String, strMSTTK As Variant)
    Dim i As Integer
    Dim arrMangGTHeader() As String
    Dim arrErrMangSetErr() As String
    Dim arrmangMessage() As String
    Dim varTemp As Variant, vMessage As Variant
    Dim lCol As Long, lRow As Long

    
    arrMangGTHeader = Split(chuoiTTHeader, "~")
    arrErrMangSetErr = Split(viTriSetErr, "~")
    arrmangMessage = Split(chuoiMessage, "~")
    
    With pGrid
        
        If isCheckTTDLT() Then
        
            For i = 0 To UBound(arrMangGTHeader)
                .Sheet = 1
                'lay vi tri cua thong tin tren header
                ParserCellID pGrid, arrMangGTHeader(i), lCol, lRow
                'lay gia tri cua tung thong tin herder
                .GetText lCol, lRow, varTemp
    
                'set lai gia tri header neu co loi
                If Trim(varTemp) = vbNullString Then
                    ' set lai header va lay message
                    .Sheet = .SheetCount
                    ParserCellID pGrid, arrErrMangSetErr(i), lCol, lRow
                    .SetText lCol, lRow, "0"
                    ParserCellID pGrid, arrmangMessage(i), lCol, lRow
                    .GetText lCol, lRow, vMessage
                    'set color va set message
                    .Sheet = 1
                    ParserCellID pGrid, arrMangGTHeader(i), lCol, lRow
                    .Col = lCol
                    .Row = lRow
                    .CellNote = vMessage
                    .BackColor = mErrorColor
                    'lay vi tri tren sheet header de danh dau loi
                Else
                    .Sheet = 1
                    ParserCellID pGrid, arrMangGTHeader(i), lCol, lRow
                    .Col = lCol
                    .Row = lRow
                    .CellNote = ""
                    If i = UBound(arrMangGTHeader) Or i = UBound(arrMangGTHeader) - 1 Then
                        .BackColor = mNonErrorColor
                    Else
                        .BackColor = mFormColor
                    End If
                    
                    .Sheet = .SheetCount
                    ParserCellID pGrid, arrErrMangSetErr(i), lCol, lRow
                    .SetText lCol, lRow, "1"
                End If
            Next
            
        Else
                Exit Sub
        End If
    End With
End Sub

'ham kiem tra nguoi su dung co check vao phan ke khai thong tin mst khong

Public Function isCheckTTDLT() As Boolean
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlDom2 As New MSXML.DOMDocument
    Dim value As String
    
        'check xem NNT co check vao o check nhap thong tin header k, neu co thi ktra header k duoc bo trong, k check thi exit
        xmlDom2.Load TAX_Utilities_v1.DataFolder & "\Header_01.xml"
        Set xmlNodeList = xmlDom2.getElementsByTagName("Cell")
        Set xmlNode = xmlNodeList.Item(31)
        value = GetAttribute(xmlNode, "Value")
        isCheckTTDLT = False
        If UCase(Trim(value)) = "X" Or UCase(Trim(value)) = "1" Then
            isCheckTTDLT = True
        End If
End Function

' ham kiem tra neu nguoi dung k luu lai thong tin mstdl trong lan su dung truoc nhung bo sung them thong tin mstdl
' vao lan su dung sau thi cap nhan lai thong tin mstdl

Public Sub updateMSTDL(pGrid As fpSpread, CellTTHeader As String)
    Dim lRow As Long, lCol As Long
    Dim mstDL, tenNVDLT, chungChiHN As Variant
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlDom2 As New MSXML.DOMDocument
    Dim value As String
    Dim arrMangTTHeader() As String
    Dim vTMSTDL As String, vTTenNNDLT As String, vTCCHN As String
    
    arrMangTTHeader = Split(CellTTHeader, "~")
    vTMSTDL = arrMangTTHeader(0)
    vTTenNNDLT = arrMangTTHeader(1)
    vTCCHN = arrMangTTHeader(2)
    With pGrid
    
        .Sheet = 1
        If isCheckTTDLT = True Then
            'check MST DaiLy, ten NVDLT, chungchi hanh nghe da ke khai tu lan truoc chua
            xmlDom2.Load TAX_Utilities_v1.DataFolder & "\Header_01.xml"
            Set xmlNodeList = xmlDom2.getElementsByTagName("Cell")
            
            'get MSTDL vu
            ParserCellID pGrid, vTMSTDL, lCol, lRow
            .GetText lCol, lRow, mstDL
            If Trim(mstDL) = vbNullString Then
                 Set xmlNode = xmlNodeList.Item(32)
                value = GetAttribute(xmlNode, "Value")
                .SetText lCol, lRow, value
                UpdateCell pGrid, lCol, lRow, value
            End If
             'get tenNVDLT
            ParserCellID pGrid, vTTenNNDLT, lCol, lRow
            .GetText lCol, lRow, tenNVDLT
            If Trim(tenNVDLT) = vbNullString Then
                Set xmlNode = xmlNodeList.Item(42)
                value = GetAttribute(xmlNode, "Value")
                .SetText lCol, lRow, value
                UpdateCell pGrid, lCol, lRow, value
            End If
            'get CCHN
            ParserCellID pGrid, vTCCHN, lCol, lRow
            .GetText lCol, lRow, chungChiHN
            If Trim(chungChiHN) = vbNullString Then
                Set xmlNode = xmlNodeList.Item(43)
                value = GetAttribute(xmlNode, "Value")
                .SetText lCol, lRow, value
                UpdateCell pGrid, lCol, lRow, value
            End If

        Else
            Exit Sub
        End If
    End With
End Sub


' Ham tinh lai han nop doi voi cac to khai GH
Public Sub TinhHanNop_KHBS(pGrid As fpSpread)
    On Error GoTo ErrorHandle
    
    Dim songaynopcham As Long
    Dim hannop As String
    Dim ngayKHBS  As Variant
        ' To khai 01/GTGT gia han thang 4,5,6 nam 2012 -> tinh lai han nop
         If GetAttribute(TAX_Utilities_v1.NodeMenu, "ID") = "01" Then
             If (TAX_Utilities_v1.Month = 4 Or TAX_Utilities_v1.Month = 5 Or TAX_Utilities_v1.Month = 6) And TAX_Utilities_v1.Year = 2012 And TAX_Utilities_v1.CheckToKhaiGH = True Then
                 If TAX_Utilities_v1.Month = 4 Then
                     hannop = "20/" & "11" & "/" & TAX_Utilities_v1.Year
                 ElseIf TAX_Utilities_v1.Month = 5 Then
                     hannop = "20/" & "12" & "/" & TAX_Utilities_v1.Year
                 ElseIf TAX_Utilities_v1.Month = 6 Then
                     hannop = "21/" & "01" & "/" & TAX_Utilities_v1.Year + 1
                 End If
             Else
                 ' cac ky ke khai khac van tinh han nop binh thuong
                 If TAX_Utilities_v1.Month = 12 Then
                     hannop = "20/" & "01" & "/" & TAX_Utilities_v1.Year + 1
                 ElseIf TAX_Utilities_v1.Month = 4 Then
                     hannop = "02/" & "05" & "/" & TAX_Utilities_v1.Year
                 Else
                     hannop = "20/" & Right("00" & TAX_Utilities_v1.Month + 1, 2) & "/" & TAX_Utilities_v1.Year
                 End If
             End If
        End If
    
        'Neu vao ngay thu 7 thi cong them 2 ngay,  ngay CN thi cong them mot ngay
        If Weekday(CDate(hannop)) = 7 Then
            hannop = DateAdd("D", 2, CDate(hannop))
            hannop = Format(hannop, "dd/mm/yyyy")
        ElseIf Weekday(CDate(hannop)) = 1 Then
            hannop = DateAdd("D", 1, CDate(hannop))
            hannop = Format(hannop, "dd/mm/yyyy")
        End If
        
        With pGrid
            .Sheet = .SheetCount - 1
            .GetText .ColLetterToNumber("BG"), 5, ngayKHBS
            ' Tinh so ngay nop cham
             songaynopcham = numberb2d(hannop, CStr(ngayKHBS))
            .SetText .ColLetterToNumber("E"), 24, hannop
            .SetText .ColLetterToNumber("BD"), 5, songaynopcham
        End With
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "TinhHanNop_KHBS", Err.number, Err.Description
End Sub


Private Function numberb2d(fd As String, td As String) As Integer
    numberb2d = DateDiff("d", s2d(fd), s2d(td))
    If numberb2d <= 0 Then numberb2d = 0
End Function


Private Function s2d(d As String) As Date
   Dim strFormat() As String
    strFormat = Split(d, "/")
    s2d = DateSerial(strFormat(2), strFormat(1), strFormat(0))
    
End Function



' set co quan thue ra quyet dinh hoan thue
Public Sub setCQTQuanLyHoanThue(pGrid As fpSpread)
        Dim xmlDomData As New MSXML.DOMDocument
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim arrDanhsach() As String
        Dim tenCucThue As String
        Dim maCucThue As String
        Dim tenChiCucThue As String
        Dim maChiCucThue As String
        
        Dim loaiCq As String
        Dim maLoaiCq As String
        
        
        
        
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    
                    If arrDanhsach(0) = "1" Then
                        tenCucThue = tenCucThue + arrDanhsach(3) + Chr$(9)
                        maCucThue = maCucThue + arrDanhsach(1) + Chr$(9)
                    End If
                    'If arrDanhsach(0) = "0" Then
                        tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                        maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                    'End If
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        ' setloai
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_DM_LoaiCq.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    loaiCq = loaiCq + arrDanhsach(1) + Chr$(9)
                    maLoaiCq = maLoaiCq + arrDanhsach(0) + Chr$(9)
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        With pGrid
            .Sheet = .SheetCount - 1 'set thong tin CQL
            '.EventEnabled(EventAllEvents) = False
            
            .Row = 33
            .Col = .ColLetterToNumber("BE")
            .TypeComboBoxList = tenCucThue
            
            .Row = 33
            .Col = .ColLetterToNumber("BI")
            .TypeComboBoxList = maCucThue
            
            .Row = 35
            .Col = .ColLetterToNumber("BE")
            .TypeComboBoxList = tenChiCucThue
    
            .Row = 35
            .Col = .ColLetterToNumber("BI")
            .TypeComboBoxList = maChiCucThue
                                
            '.EventEnabled(EventAllEvents) = True
        End With
End Sub


Public Function CheckNgayInKyKK(ngay As String, thangQuy As Integer, nam As Integer, isQuy As Boolean) As Boolean
    Dim strDate() As String
    Dim vdate As Date
    strDate = Split(ngay, "/")
    vdate = DateSerial(CInt(strDate(2)), CInt(strDate(1)), CInt(strDate(0)))
    ' kiem tra thuoc thang/quy ke khai va thang quy truoc lien ke
    ' y/c chi Thuy sau dao tao HTKK
    If isQuy = True Then
        If (DatePart("Q", vdate) = thangQuy And DatePart("YYYY", vdate) = nam) Then
            CheckNgayInKyKK = True
        Else
            CheckNgayInKyKK = False
        End If
    Else
        If (DatePart("M", vdate) = thangQuy And DatePart("YYYY", vdate) = nam) Then
            CheckNgayInKyKK = True
        Else
            CheckNgayInKyKK = False
        End If
    End If
End Function


Public Function CheckNgayTruocKyKK(ngay As String, thangQuy As Integer, nam As Integer, isQuy As Boolean) As Boolean
    Dim strDate() As String
    Dim vdate As Date
    strDate = Split(ngay, "/")
    vdate = DateSerial(CInt(strDate(2)), CInt(strDate(1)), CInt(strDate(0)))
    ' kiem tra thuoc thang/quy ke khai va thang quy truoc lien ke
    ' y/c chi Thuy sau dao tao HTKK
    If isQuy = True Then
        If (DatePart("Q", vdate) = thangQuy And DatePart("YYYY", vdate) = nam) Or (DatePart("Q", vdate) <= thangQuy And DatePart("YYYY", vdate) = nam) Or (DatePart("YYYY", vdate) < nam) Then
            CheckNgayTruocKyKK = True
        Else
            CheckNgayTruocKyKK = False
        End If
    Else
        If (DatePart("M", vdate) = thangQuy And DatePart("YYYY", vdate) = nam) Or (DatePart("M", vdate) <= thangQuy And DatePart("YYYY", vdate) = nam) Or (DatePart("YYYY", vdate) < nam) Then
            CheckNgayTruocKyKK = True
        Else
            CheckNgayTruocKyKK = False
        End If
    End If
End Function

' Kiem tra ngay nho hon ngay cuoi ky ke khai
Public Function getNgayCuoiKyKK(thangQuy As Integer, nam As Integer, isQuy As Boolean) As Date
    Dim vdate As Date
    Dim thangCuoiQuy As Integer
    Dim tmpDate As Date
     If isQuy = True Then
       thangCuoiQuy = (thangQuy - 1) * 3 + 3
       tmpDate = DateSerial(nam, thangCuoiQuy, 1)
       ' ngay cuoi quy
       vdate = DateAdd("M", 1, tmpDate)
       vdate = DateAdd("D", -1, vdate)
    Else
       tmpDate = DateSerial(nam, thangQuy, 1)
       ' ngay cuoi quy
       vdate = DateAdd("M", 1, tmpDate)
       vdate = DateAdd("D", -1, vdate)
    End If
    getNgayCuoiKyKK = vdate
End Function

' get loai hoa don
Public Function getLoaiHD(maHoaDon As String) As String
    Dim strDataFileName As String
    Dim arrDanhsach() As String
    Dim strComboHien As String
    Dim strCombo As String
    Dim MSTDN As String
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strTemp As String
    
    
    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
    ' Lay danh muc loai hoa don
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If maHoaDon = Trim$(arrDanhsach(1)) Then
                    strTemp = Trim$(arrDanhsach(0))
                    Exit For
                End If
            End If
        Next
    End If
    getLoaiHD = strTemp
End Function





'Kiem tra cau truc mau so
'str: chuoi mau so
'strLoai: 0,1,2,3
'strTemp: cac ky tu tien to cua loai HD
'strKyHieuHD: ky hieu hoa don

'CheckSoLienHDDT = 3 sai so Lien hoa don dien tu


Public Function CheckSoLienHDDT(str As String, strLoai As String, strTemp As String, strKyHieuHD As String) As String
    Dim result As String
    Dim soLien As String
    Dim kyTuNganCach As String
    Dim strSoTT As Variant
    Dim strBD As Variant
    result = "0"
    If strLoai = "0" Then
        If Len(Trim(str)) <> 11 And Len(Trim(str)) <> 13 Then
        Else
            If Left$(Trim(str), 6) = Trim(strTemp) Then
                soLien = Mid$(Trim(str), 7, 1)
                kyTuNganCach = Mid$(Trim(str), 8, 1)
                strSoTT = Mid$(Trim(str), 9, 3)
                strBD = Mid$(Trim(str), 12, 2)
                ' so lien phai phai bang 0
                If IsNumeric(soLien) Then
                    If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If Val(soLien) <> 0 Then
                            result = "3"
                            CheckSoLienHDDT = result
                            Exit Function
                        End If
                    Else
                        result = "0"
                        CheckSoLienHDDT = result
                        Exit Function
                    End If
                Else
                    result = "3"
                    CheckSoLienHDDT = result
                    Exit Function
                End If
            End If
        End If
    End If
    CheckSoLienHDDT = result
End Function



' Lay ten CQT theo ma
Public Sub GetCQTByMaCQT(ByVal maCQT As String, Optional ByRef TenCQT As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = "..\InterfaceIni\Catalogue_Tinh_Thanh.xml"
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maCQT = arrDanhsach(1) And arrDanhsach(0) = 0 Then
                            TenCQT = arrDanhsach(3)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


' Lay ma hoa don theo ten
Public Sub GetMaHoaDon(ByVal maHD As String, Optional ByRef tenHD As String)
     Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
       tenHD = ""
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If InStr(1, maHD, arrDanhsach(1)) > 0 Then
                            tenHD = arrDanhsach(1) & "###" & strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


' Lay ma bien lai phi, le phi theo ten
Public Sub GetMaBLP(ByVal maBLP As String, Optional ByRef tenBLP As String)
     Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_BLP.xml")
       tenBLP = ""
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If InStr(1, maBLP, arrDanhsach(1)) > 0 Then
                            tenBLP = arrDanhsach(1) & "###" & strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub
