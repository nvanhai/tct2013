VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeader_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used for general information
'this Class is belong to TAX_Business_v2 project which will be compline to DLL
'Coder: PhuocHK
Option Explicit

Private Const S1MST1Row = 2
Private Const S1MST1Col = "C"
Private Const S1MST2Row = 2
Private Const S1MST2Col = "D"
Private Const S1MST3Row = 2
Private Const S1MST3Col = "E"
Private Const S1MST4Row = 2
Private Const S1MST4Col = "F"
Private Const S1MST5Row = 2
Private Const S1MST5Col = "G"
Private Const S1MST6Row = 2
Private Const S1MST6Col = "H"
Private Const S1MST7Row = 2
Private Const S1MST7Col = "I"
Private Const S1MST8Row = 2
Private Const S1MST8Col = "J"
Private Const S1MST9Row = 2
Private Const S1MST9Col = "K"
Private Const S1MST10Row = 2
Private Const S1MST10Col = "L"
Private Const S1MST11Row = 2
Private Const S1MST11Col = "N"
Private Const S1MST12Row = 2
Private Const S1MST12Col = "O"
Private Const S1MST13Row = 2
Private Const S1MST13Col = "P"

Private Const S1PhoneRow = 12
Private Const S1PhoneCol = "C"

Private Const S1FaxRow = 14
Private Const S1FaxCol = "C"

Private Const S1EmailRow = 16
Private Const S1EmailCol = "C"

Private Const S1AccountRow = 20
Private Const S1AccountCol = "C"

Private Const S1StartDateRow = 24
Private Const S1StartDateCol = "C"

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Dim NgayBatDau As Variant

'****************************************
'Description: Prepare1 procedure set type of C_24 Cell
'Author: ThanhDX
'Date:09/01/2006
'****************************************
Public Sub Prepare1()
    Dim bl As Boolean
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim arrDanhsach() As String
    Dim tenCucThue As String
    Dim maCucThue As String
    Dim tenChiCucThue As String
    Dim maChiCucThue As String
    

    
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
    
    
    With fps
        .Sheet = 1 'Thong tin doanh nghiep
        'Format mask to C_24 Cell
        .EventEnabled(EventAllEvents) = False
        .Row = 24
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypePic
        .TypePicMask = "99//99"
        .TypePicDefaultText = ""
        '.Text = ""
        
        SetDateFormat fps, 1, 57, .ColLetterToNumber("C"), DDMMYYYY
        .Row = 57
        .Col = .ColLetterToNumber("C")
        .TypeHAlign = TypeHAlignLeft
        
        
        .Row = 28
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = tenCucThue
        .Row = 28
        .Col = .ColLetterToNumber("Q")
        .TypeComboBoxList = maCucThue
        
        .Row = 30
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = tenChiCucThue

        .Row = 30
        .Col = .ColLetterToNumber("Q")
        .TypeComboBoxList = maChiCucThue

        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    'Dim DataText As Variant
    Dim varIdCucThue As Variant
    Dim vCheck As Variant
    
    With fps
        .Sheet = 1
        ' Lay Id cua cuc thue vua chon o tren
        .GetText .ColLetterToNumber("Q"), 28, varIdCucThue

        If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
            ' Lay thong tin cho chi cuc thue
            Dim xmlDomData As New MSXML.DOMDocument
            Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
            Dim xmlNode As MSXML.IXMLDOMNode
            Dim arrDanhsach() As String
            Dim tenChiCucThue As String
            Dim maChiCucThue As String

            varIdCucThue = Left$(varIdCucThue, 3)

            If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                For Each xmlNode In xmlNodeListCell
                    If GetAttribute(xmlNode, "Value") <> "" Then
                        arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")

                        If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                            tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                            maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                        End If
                    End If
                Next
                Set xmlDomData = Nothing
                Set xmlNodeListCell = Nothing
                Set xmlNode = Nothing
            End If

            .Row = 30
            .Col = .ColLetterToNumber("C")
            .TypeComboBoxList = tenChiCucThue

            .Row = 30
            .Col = .ColLetterToNumber("Q")
            .TypeComboBoxList = maChiCucThue
            
         ' Thong tin dai ly thue
        .GetText .ColLetterToNumber("C"), 37, vCheck
        If vCheck = "1" Or UCase(vCheck) = "X" Then
            .Col = .ColLetterToNumber("C")
            .Row = 39
            .Lock = False
            .Row = 41
            .Lock = False
            .Row = 43
            .Lock = False
            .Row = 45
            .Lock = False
            .Row = 47
            .Lock = False
            .Row = 49
            .Lock = False
            .Row = 51
            .Lock = False
            .Row = 53
            .Lock = False
            .Row = 55
            .Lock = False
            .Row = 57
            .Lock = False
            .Row = 59
            .Lock = False
            .Row = 61
            .Lock = False
        End If

         ' Thong tin co quan chi tra
        .GetText .ColLetterToNumber("P"), 32, vCheck
        If vCheck = "1" Or UCase(vCheck) = "X" Then
            .Col = .ColLetterToNumber("C")
            .Row = 32
            .Lock = False
            .Row = 34
            .Lock = False
        End If



        End If
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
        Dim Phone As Variant, Fax As Variant, Email As Variant, Account As Variant, StartDate As Variant
    Dim Err1 As Integer, Err2 As Integer, Err3 As Integer
    Dim Err4 As Integer, Err5 As Integer, Err6 As Integer
    Dim Err7 As Integer, ErrStartDate As Integer
    Dim strCheck As String
    ' vttoan: them phan thong tin don vi chu quan
    Dim strMSTDVCQ As Variant
    Dim vErrMsgMST As Variant
    Dim iFlagEmty As Variant
    
    Dim iFlagTrungMST As Integer
    
    ' vttoan them thong tin dai ly thue
    Dim vErrmstDaiLy As Variant
    Dim vErrtenDaiLy As Variant
    Dim vErrdiachiDaiLy As Variant
    Dim vErrqhuyenDaiLy As Variant
    Dim vErrtphoDaiLy As Variant
    Dim vErrTenNhanVienDLT As Variant
    Dim vErrChungChiSo As Variant
    Dim vErrNgayHDDaiLy As Variant
    Dim vErrSoHDDaiLy As Variant
    Dim vErrNgayHDDaiLyEmpty As Variant
    
    
    Dim mstDL As Variant
    Dim tenDL As Variant
    Dim diachiDL As Variant
    Dim qhuyenDL As Variant
    Dim tphoDL As Variant
    Dim checkDL As Variant
    Dim strMST1, vTempValue As Variant
    Dim vTenNhanVienDLT As Variant
    Dim vChungChiSo As Variant
    Dim vNgayHDDaiLy As Variant
    Dim arrNgay() As String
    
    Dim vSoHDDaiLy As Variant
    
    
    Dim vErrorTrungMST As Variant
    Dim chkDVCQ As Variant
    Dim vMErrTenDVCQ As Variant, vTenDVCQ As Variant, iFlagTenDVCQEmpty As Integer
    Dim vMErrMSTDVCQ As Variant, iFlagMSTDVCQEmpty As Integer
    ' end dvcq
    Err1 = 0
    Err2 = 0
    Err3 = 0
    Err4 = 0
    Err5 = 0
    Err6 = 0
    Err7 = 0
    ErrStartDate = 0
    iFlagTrungMST = 0
    Dim vValue
    With fps

        .Sheet = .SheetCount
        ' Lay thong bao ve sai ma so thue trong sheet header
            .GetText .ColLetterToNumber("E"), 23, vErrMsgMST
        ' lay thong bao loi ten dvcq bat buoc phai nhap
            .GetText .ColLetterToNumber("E"), 36, vMErrTenDVCQ
        ' lay thong bao loi mst dvcq bat buoc phai nhap
            .GetText .ColLetterToNumber("E"), 37, vMErrMSTDVCQ
        .Sheet = 1
        ' Lay ve MSt
        .GetText .ColLetterToNumber("C"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("D"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("E"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("F"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("G"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("H"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("I"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("J"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("K"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("L"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("N"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("O"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        .GetText .ColLetterToNumber("P"), 2, vTempValue
        strMST1 = strMST1 & Trim(vTempValue)
        
        'get chk Thong tin don vi chu quan
        .GetText .ColLetterToNumber("P"), 32, chkDVCQ
        If Trim(chkDVCQ) = "1" Or UCase(Trim(chkDVCQ)) = "X" Then
            
            .GetText .ColLetterToNumber("C"), 34, strMSTDVCQ
            .Row = 34
            If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
                If Len(strMSTDVCQ) = 14 Then
                    strMSTDVCQ = Left$(strMSTDVCQ, 10) & Right$(strMSTDVCQ, 3)
                End If
            End If
       ' vttoan: check cau truc mst don vi chu quan
            'dntai set mst valid truoc khi check
                .Col = .ColLetterToNumber("C")
                .CellNote = ""
                .BackColor = mNonErrorColor

                If Trim(strMSTDVCQ) <> vbNullString Then
                       strCheck = ""
                       If Len(strMSTDVCQ) = 10 Or Len(strMSTDVCQ) = 13 Then
                         strCheck = CheckTaxCode(Mid(strMSTDVCQ, 1, 1), Mid(strMSTDVCQ, 2, 1), Mid(strMSTDVCQ, 3, 1), Mid(strMSTDVCQ, 4, 1), Mid(strMSTDVCQ, 5, 1), Mid(strMSTDVCQ, 6, 1), Mid(strMSTDVCQ, 7, 1), Mid(strMSTDVCQ, 8, 1), Mid(strMSTDVCQ, 9, 1), Mid(strMSTDVCQ, 10, 1), Mid(strMSTDVCQ, 11, 1), Mid(strMSTDVCQ, 12, 1), Mid(strMSTDVCQ, 13, 1))
                            If Len(strMSTDVCQ) = 10 And strCheck <> "000" Then
                                .Col = .ColLetterToNumber("C")
                                .CellNote = vErrMsgMST
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            ElseIf Len(strMSTDVCQ) = 13 And (strCheck <> "000" Or IsNumeric(strMSTDVCQ) = False) Then
                                .Col = .ColLetterToNumber("C")
                                .CellNote = vErrMsgMST
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            End If
                       ElseIf Len(strMSTDVCQ) <> 10 Or Len(strMSTDVCQ) <> 13 Then
                            .Col = .ColLetterToNumber("C")
                            .CellNote = vErrMsgMST
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                       End If
                       Else
                    .Col = .ColLetterToNumber("C")
                    .CellNote = vMErrMSTDVCQ
                    .BackColor = mErrorColor
                    iFlagMSTDVCQEmpty = 1
                End If
                        
            'get ten dvcq
            .GetText .ColLetterToNumber("C"), 32, vTenDVCQ
            If Trim(vTenDVCQ) = vbNullString Then
                .Col = .ColLetterToNumber("C")
                .Row = 32
                .CellNote = vMErrTenDVCQ
                .BackColor = mErrorColor
                iFlagTenDVCQEmpty = 1
            Else
                .Col = .ColLetterToNumber("C")
                .Row = 32
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Else
            .Col = .ColLetterToNumber("C")
            .Row = 34
            .CellNote = ""
            .BackColor = mNonErrorColor
                                    
            .Row = 32
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        ' vttoan them thong tin dai ly thue "thong tu 28"
        .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 24, vErrmstDaiLy
            .GetText .ColLetterToNumber("E"), 25, vErrtenDaiLy
            .GetText .ColLetterToNumber("E"), 26, vErrdiachiDaiLy
            .GetText .ColLetterToNumber("E"), 27, vErrqhuyenDaiLy
            .GetText .ColLetterToNumber("E"), 28, vErrtphoDaiLy
            .GetText .ColLetterToNumber("E"), 29, vErrorTrungMST
            
            .GetText .ColLetterToNumber("E"), 31, vErrTenNhanVienDLT
            .GetText .ColLetterToNumber("E"), 32, vErrChungChiSo
            .GetText .ColLetterToNumber("E"), 33, vErrNgayHDDaiLy
            
            .GetText .ColLetterToNumber("E"), 34, vErrSoHDDaiLy
            .GetText .ColLetterToNumber("E"), 35, vErrNgayHDDaiLyEmpty
            
        .Sheet = 1
            .GetText .ColLetterToNumber("C"), 39, mstDL
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
            If Len(mstDL) = 14 Then
                mstDL = Left$(mstDL, 10) & Right$(mstDL, 3)
            End If
        End If
            .GetText .ColLetterToNumber("C"), 41, tenDL
            .GetText .ColLetterToNumber("C"), 43, diachiDL
            .GetText .ColLetterToNumber("C"), 45, qhuyenDL
            .GetText .ColLetterToNumber("C"), 47, tphoDL
            .GetText .ColLetterToNumber("C"), 37, checkDL
            
            .GetText .ColLetterToNumber("C"), 59, vTenNhanVienDLT
            .GetText .ColLetterToNumber("C"), 61, vChungChiSo
            
            .GetText .ColLetterToNumber("C"), 55, vSoHDDaiLy
            .GetText .ColLetterToNumber("C"), 57, vNgayHDDaiLy
                    If Trim(checkDL) = 1 Then
                            If Trim(mstDL) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 39
                                .CellNote = vErrmstDaiLy
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            ElseIf Trim(mstDL) <> vbNullString Then
                                strCheck = ""
                                If Len(mstDL) = 10 Or Len(mstDL) = 13 Then
                                  strCheck = CheckTaxCode(Mid(mstDL, 1, 1), Mid(mstDL, 2, 1), Mid(mstDL, 3, 1), Mid(mstDL, 4, 1), Mid(mstDL, 5, 1), Mid(mstDL, 6, 1), Mid(mstDL, 7, 1), Mid(mstDL, 8, 1), Mid(mstDL, 9, 1), Mid(mstDL, 10, 1), Mid(mstDL, 11, 1), Mid(mstDL, 12, 1), Mid(mstDL, 13, 1))
                                     If Len(mstDL) = 10 And strCheck <> "000" Then
                                         .Col = .ColLetterToNumber("C")
                                         .Row = 39
                                         .CellNote = vErrMsgMST
                                         .BackColor = mErrorColor
                                         iFlagEmty = 1
                                     ElseIf Len(mstDL) = 13 And (strCheck <> "000" Or IsNumeric(mstDL) = False) Then
                                         .Col = .ColLetterToNumber("C")
                                         .Row = 39
                                         .CellNote = vErrMsgMST
                                         .BackColor = mErrorColor
                                         iFlagEmty = 1
                                     ElseIf Trim(mstDL) = Trim(strMST1) Then
                                         .Col = .ColLetterToNumber("C")
                                         .Row = 39
                                         .CellNote = vErrorTrungMST
                                         .BackColor = mAlertColor
                                         iFlagTrungMST = 1
                                     Else
                                         .Col = .ColLetterToNumber("C")
                                         .Row = 39
                                         .CellNote = ""
                                         .BackColor = mNonErrorColor
                                     End If
                                ElseIf Len(mstDL) <> 10 Or Len(mstDL) <> 13 Then
                                     .Col = .ColLetterToNumber("C")
                                     .Row = 39
                                     .CellNote = vErrMsgMST
                                     .BackColor = mErrorColor
                                     iFlagEmty = 1
                                Else
                                     .Col = .ColLetterToNumber("C")
                                     .Row = 39
                                     .CellNote = ""
                                     .BackColor = mNonErrorColor
                                End If
                            
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 39
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            If Trim(tenDL) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 41
                                .CellNote = vErrtenDaiLy
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 41
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            If Trim(diachiDL) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 43
                                .CellNote = vErrdiachiDaiLy
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 43
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            If Trim(qhuyenDL) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 45
                                .CellNote = vErrqhuyenDaiLy
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 45
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            If Trim(tphoDL) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 47
                                .CellNote = vErrtphoDaiLy
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 47
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            ' Kiem tra ten nhan vien dai ly
                            If Trim(vTenNhanVienDLT) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 59
                                .CellNote = vErrTenNhanVienDLT
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 59
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            ' Kiem tra chung chi so
                            If Trim(vChungChiSo) = "" Then
                                .Col = .ColLetterToNumber("C")
                                .Row = 61
                                .CellNote = vErrChungChiSo
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 61
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                            ' Kiem tra ngay HD khong duoc lon hon ngay hien tai
                            If Trim(vNgayHDDaiLy) <> "" And Trim$(vNgayHDDaiLy) <> "../../...." Then
                                arrNgay = Split(vNgayHDDaiLy, "/")
                                If DateDiff("D", DateSerial(CInt(arrNgay(2)), CInt(arrNgay(1)), CInt(arrNgay(0))), Date) < 0 Then
                                    .Col = .ColLetterToNumber("C")
                                    .Row = 57
                                    .CellNote = vErrNgayHDDaiLy
                                    .BackColor = mErrorColor
                                    iFlagEmty = 1
                                Else
                                    .Col = .ColLetterToNumber("C")
                                    .Row = 57
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            Else
                                .Col = .ColLetterToNumber("C")
                                .Row = 57
                                .CellNote = vErrNgayHDDaiLyEmpty
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            End If
                            ' Kiem tra so HDDL
                            If Trim$(vSoHDDaiLy) = "" Or vSoHDDaiLy = vbNullString Then
                                 .Col = .ColLetterToNumber("C")
                                .Row = 55
                                .CellNote = vErrSoHDDaiLy
                                .BackColor = mErrorColor
                                iFlagEmty = 1
                            Else
                                 .Col = .ColLetterToNumber("C")
                                .Row = 55
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                    Else
                        .Col = .ColLetterToNumber("C")
                        .Row = 39
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 41
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 43
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 45
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 47
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 59
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 61
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber("C")
                        .Row = 55
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        .Row = 57
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        
                        
                    End If
'----------------------------------------------------------------------------
        .Sheet = 1
         mCurrentSheet = .Sheet
        .Row = 2
        .GetText .ColLetterToNumber(S1StartDateCol), S1StartDateRow, StartDate
        If Format_ddmm(CStr(StartDate)) = "" Then
            ErrStartDate = 1
        Else
            .SetText .ColLetterToNumber(S1StartDateCol), S1StartDateRow, Format_ddmm(CStr(StartDate))
            UpdateCell fps, .ColLetterToNumber(S1StartDateCol), S1StartDateRow, Format_ddmm(CStr(StartDate))
        End If
        
        'Check MST
        .GetText .ColLetterToNumber(S1MST1Col), S1MST1Row, MST1
        .GetText .ColLetterToNumber(S1MST2Col), S1MST2Row, MST2
        .GetText .ColLetterToNumber(S1MST3Col), S1MST3Row, MST3
        .GetText .ColLetterToNumber(S1MST4Col), S1MST4Row, MST4
        .GetText .ColLetterToNumber(S1MST5Col), S1MST5Row, MST5
        .GetText .ColLetterToNumber(S1MST6Col), S1MST6Row, MST6
        .GetText .ColLetterToNumber(S1MST7Col), S1MST7Row, MST7
        .GetText .ColLetterToNumber(S1MST8Col), S1MST8Row, MST8
        .GetText .ColLetterToNumber(S1MST9Col), S1MST9Row, MST9
        .GetText .ColLetterToNumber(S1MST10Col), S1MST10Row, MST10
        .GetText .ColLetterToNumber(S1MST11Col), S1MST11Row, MST11
        .GetText .ColLetterToNumber(S1MST12Col), S1MST12Row, MST12
        .GetText .ColLetterToNumber(S1MST13Col), S1MST13Row, MST13
    
        strCheck = CheckTaxCode(MST1, MST2, MST3, MST4, MST5, MST6, MST7, MST8, MST9, MST10, MST11, MST12, MST13)
        Err1 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)
        Err2 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)
        Err3 = CInt(strCheck)
        
        
        .Sheet = .SheetCount
        'set formula for MST
        .Col = .ColLetterToNumber("B")
        .Row = 12
        .Formula = IIf(Err1 = 1, "0", "1")
        .Col = .ColLetterToNumber("B")
        .Row = 13
        .Formula = IIf(Err2 = 1, "0", "1")
        .Col = .ColLetterToNumber("B")
        .Row = 14
        .Formula = IIf(Err3 = 1, "0", "1")
        .Row = 19
        .Formula = IIf(ErrStartDate = 1, "0", "1")
        
        .Col = 1
        .Row = 1
        
        '.ReDraw = True
        '.Visible = True
         'set lai trang thai header
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 23, IIf(iFlagEmty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 29, IIf(iFlagTrungMST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 36, IIf(iFlagTenDVCQEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 37, IIf(iFlagMSTDVCQEmpty = 1, "0", "1")
    End With
End Sub

'This function is ued to support report
'some special data that we want to save to data files
'Public Sub SaveExceptionData()
'    With fps
'        .Sheet = 1
'        mCurrentSheet = .Sheet
'        .Row = 22
'        .Col = .ColLetterToNumber("C")
'        If .TypeComboBoxCurSel <> 0 Then
'            UpdateCell fps, .ColLetterToNumber("C"), 22, "101_" & .TypeComboBoxCurSel
'        Else
'            UpdateCell fps, .ColLetterToNumber("C"), 22, ""
'        End If
'    End With
'End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    fps.EventEnabled(EventAllEvents) = False
    ASheet = fps.ActiveSheet
    SSheet = fps.Sheet
    
    CheckDynamicError
    
    fps.ActiveSheet = ASheet
    fps.Sheet = SSheet
    fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim checkDL As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        .Visible = True
        If .ActiveSheet = 1 Then
            ' Xu ly thong tin dai ly thue
            If Row = 37 And Col = .ColLetterToNumber("C") Then
                .GetText .ColLetterToNumber("C"), 37, checkDL
                If checkDL = "0" Then
                        ' MST dai ly
                        .SetText .ColLetterToNumber("C"), 39, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 39, ""
                        ' Ten dai ly
                        .SetText .ColLetterToNumber("C"), 41, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 41, ""
                        ' Dia chi DL thu
                        .SetText .ColLetterToNumber("C"), 43, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 43, ""
                        ' Quan huyen DL
                        .SetText .ColLetterToNumber("C"), 45, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 45, ""
                        ' Tinh /thanh DL
                        .SetText .ColLetterToNumber("C"), 47, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 47, ""
                        ' Dien thoai
                        .SetText .ColLetterToNumber("C"), 49, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 49, ""
                        ' Fax
                        .SetText .ColLetterToNumber("C"), 51, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 51, ""
                        ' Email
                        .SetText .ColLetterToNumber("C"), 53, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 53, ""
                        ' So HD
                        .SetText .ColLetterToNumber("C"), 55, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 55, ""
                        ' Ngay hop dong
                        .SetText .ColLetterToNumber("C"), 57, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 57, ""
                        ' Ho ten nhan vien dai ly
                        .SetText .ColLetterToNumber("C"), 59, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 59, ""
                        ' Chung chi hanh nghe so
                        .SetText .ColLetterToNumber("C"), 61, ""
                        UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 61, ""
                        .Col = .ColLetterToNumber("C")
                        .Row = 39
                        .Lock = True
                        .Row = 41
                        .Lock = True
                        .Row = 43
                        .Lock = True
                        .Row = 45
                        .Lock = True
                        .Row = 47
                        .Lock = True
                        .Row = 49
                        .Lock = True
                        .Row = 51
                        .Lock = True
                        .Row = 53
                        .Lock = True
                        .Row = 55
                        .Lock = True
                        .Row = 57
                        .Lock = True
                        .Row = 59
                        .Lock = True
                        .Row = 61
                        .Lock = True
                Else
                        .Col = .ColLetterToNumber("C")
                        .Row = 39
                        .Lock = False
                        .Row = 41
                        .Lock = False
                        .Row = 43
                        .Lock = False
                        .Row = 45
                        .Lock = False
                        .Row = 47
                        .Lock = False
                        .Row = 49
                        .Lock = False
                        .Row = 51
                        .Lock = False
                        .Row = 53
                        .Lock = False
                        .Row = 55
                        .Lock = False
                        .Row = 57
                        .Lock = False
                        .Row = 59
                        .Lock = False
                        .Row = 61
                        .Lock = False
                End If
            ElseIf Row = 32 And Col = .ColLetterToNumber("P") Then
                .GetText .ColLetterToNumber("P"), 32, checkDL
                If Trim(checkDL) = vbNullString Or Trim(checkDL) = "0" Then
                    .SetText .ColLetterToNumber("C"), 32, ""
                    UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 32, ""
                    .SetText .ColLetterToNumber("C"), 34, ""
                    UpdateCell_sheet fps, 1, .ColLetterToNumber("C"), 34, ""
                    .Col = .ColLetterToNumber("C")
                    .Row = 32
                    .Lock = True
                    .Row = 34
                    .Lock = True
                Else
                    .Col = .ColLetterToNumber("C")
                    .Row = 32
                    .Lock = False
                    .Row = 34
                    .Lock = False
                End If
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim varTemp As Variant
    Dim intIndexCuc As Integer
    Dim intIndexChiCuc As Integer
    Dim varIdCucThue As Variant
    
  '  Dim checkDL As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        .Visible = True
        If .ActiveSheet = 1 Then
            .Sheet = 1
            mCurrentSheet = 1
            If (Col = .ColLetterToNumber("C")) And (Row = 24) Then
                .GetText Col, Row, varTemp
                .SetText Col, Row, Format_ddmm(CStr(varTemp))
                UpdateCell fps, Col, Row, .Text
                NgayBatDau = GetNgayBatDau()
                If (NgayBatDau <> .Text) And Trim(NgayBatDau) <> "" Then
                    DisplayMessage "0073", msOKOnly, miWarning
                    fps.SetFocus
                End If
                'NgayBatDau = .Text
            End If
            
            If Col = .ColLetterToNumber("C") And Row = 28 Then 'Khi chon Combo
                
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("Q")
                .Row = 28
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("Q")
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                fps.GetText .Col, .Row, varIdCucThue
                
                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
                    ' Lay thong tin cho chi cuc thue
                    Dim xmlDomData As New MSXML.DOMDocument
                    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
                    Dim xmlNode As MSXML.IXMLDOMNode
                    Dim arrDanhsach() As String
                    Dim tenChiCucThue As String
                    Dim maChiCucThue As String
                                        
                    varIdCucThue = Left$(varIdCucThue, 3)
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = 30
                    .Col = .ColLetterToNumber("C")
                    .TypeComboBoxList = tenChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = 30
                    .Col = .ColLetterToNumber("Q")
                    .TypeComboBoxList = maChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                End If
            End If
            
            ' Neu chon combo thuoc chi cuc thi hien thi va cap nhat ma chi cuc
            If Col = .ColLetterToNumber("C") And Row = 30 Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("Q")
                .Row = 30
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("Q")
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim varTaxId As Variant
    Static strValue As String
    
    
    'Dim lRow As Long, lCol As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        GetCellSpan fps, Col, Row
        If Row = 2 Then
            Select Case .ColNumberToLetter(Col)
                Case "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "N", "O", "P"
                    .Sheet = .ActiveSheet
                    .GetText Col, Row, varTaxId
                    If Len(CStr(varTaxId)) > 0 And Not IsNumeric(CStr(varTaxId)) Then
                        .SetText Col, Row, strValue
                    Else
                        strValue = CStr(varTaxId)
                    End If
            End Select
        End If
      
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_KeyPress(KeyAscii As Integer)
    Dim r As Long, c As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        r = .ActiveRow
        c = .ActiveCol
        GetCellSpan fps, c, r
        If r = 2 Then
            If Chr$(KeyAscii) = "0" Or Chr$(KeyAscii) = "1" Or Chr$(KeyAscii) = "2" Or Chr$(KeyAscii) = "3" _
                Or Chr(KeyAscii) = "4" Or Chr$(KeyAscii) = "5" Or Chr$(KeyAscii) = "6" Or Chr$(KeyAscii) = "7" _
                Or Chr(KeyAscii) = "8" Or Chr$(KeyAscii) = "9" Then
                Select Case .ColNumberToLetter(c)
                    Case "C", "D", "E", "F", "G", "H", "I", "J", "K", "N", "O"
                        .SetText c, r, Val(Chr$(KeyAscii))
                        UpdateCell fps, c, r, Val(Chr$(KeyAscii))
                        .SetActiveCell c + 1, r
                    Case "L"
                        .SetText c, r, Val(Chr$(KeyAscii))
                        UpdateCell fps, c, r, Val(Chr$(KeyAscii))
                        .SetActiveCell c + 2, r
                    Case "P"

                End Select
                TAX_Utilities_v2.AdjustData(.ActiveSheet - 1) = True
            ElseIf KeyAscii = vbKeyBack Then
                Select Case .ColNumberToLetter(c)
                    Case "D", "E", "F", "G", "H", "I", "J", "K", "L", "O", "P"
                        .SetText c, r, ""
                        UpdateCell fps, c, r, ""
                        .SetActiveCell c - 1, r
                    Case "C"
                        .SetText c, r, ""
                        UpdateCell fps, c, r, ""
                    Case "N"
                        .SetText c, r, ""
                        UpdateCell fps, c, r, ""
                        .SetActiveCell c - 2, r
                End Select
                TAX_Utilities_v2.AdjustData(.ActiveSheet - 1) = True
            Else
                KeyAscii = 0
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    
End Sub

Public Sub ResetErrorCells()
    
End Sub

Public Sub ResetData()
    
End Sub

Private Function GetNgayBatDau() As String
     
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim str43Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    GetNgayBatDau = ""
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
       
        
        strDataFileName = TAX_Utilities_v2.DataFolder & "Header_01.xml"
        
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("C_24")
            GetNgayBatDau = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
        Else
            GetNgayBatDau = ""
        End If
        
    End With
    
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "clsHeader", "GetNgayBatDau", Err.number, Err.Description
End Function

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
   Dim varTemp As Variant
   With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
        ' check ngay bat dau su dung
            If Col = .ColLetterToNumber("C") And Row = 57 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        .SetFocus
                        .SetActiveCell Col, Row
                    End If
                Else
                 .SetText Col, Row, ""
                End If
                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
            
End Sub
