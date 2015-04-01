VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TTDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL
'Coder: PhuocHK
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1


Dim strComboHienHanghoa As String, strComboHanghoa As String, strDonvitinh As String
Dim strComboHienDichvu As String, strComboDichvu As String
Dim strComboHien As String, strCombo As String
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim strTenDieuChinhCu As Variant, strKyDieuChinhCu As Variant, blDieuChinhCu As Boolean, blKyDieuChinhCu As Boolean, RowDieuChinhCu As Long
Dim arrIdHangHoa As New Scripting.Dictionary
Dim dblTongChenhLech As Double
Dim blKiemTraDieuChinhCu As Boolean, strTenToKhaiCu As Variant, blKiemTraDieuChinhCuDelete As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim arrIdHHDVGoc As New Scripting.Dictionary, arrIdHHDV As New Scripting.Dictionary
Dim arrIdHHDVBGTGoc As New Scripting.Dictionary, arrIdHHDVBGT As New Scripting.Dictionary
Dim mblKeyF5 As Boolean
Public loaiKHBS As String
Public FlagThaydoi As String
Dim arrDanhSachKHBS() As String
Dim arrData() As String
Dim mblKeyF6 As Boolean
Public StrTKThang_PS As Variant

Public Strloaitk As String
Public strLoaiNNKD As String
Public StrSolanBosung As String

Public coutDyn As Integer

Private Const Row_MSTDN = 7
Private Const Col_MSTDN = "M"

Private Const sheetKHBS = 2


' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String


'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter

'Chuc nang: Dinh dang lai Luoi truoc khi nhap du lieu
'- Lay ten danh muc file danh muc theo Menu
'- Ghi du lieu danh muc vao chuoi
'- Tao mang hang hoa dich vu tu danh sach
'- Chuyen du lieu vao Combo
'- Lay ten danh muc dau du file danh muc theo Menu
'- Tao mang hang hoa dich vu toan bo cho nhap file text
'- Tao dinh dang cho Ky ke khai thang, nam
Public Sub Prepare1()
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim strDDKTTN As String
    Dim arrDanhsach() As String, blChon As Boolean
    Dim arrDVT(15) As String
    Dim i, j, count As Long
    blChon = False
    
    'Lay ten danh muc file danh muc theo Menu
    strDataFileName = GetCatalogueFileName(blChon)
    
    arrIdHHDV.RemoveAll
    'Ghi du lieu danh muc vao chuoi
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(7) = "1") Or blChon = False Then
                    strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
                    'Tao mang hang hoa dich vu tu danh sach
                    arrIdHHDV.Add arrDanhsach(1), CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
                    arrIdHHDVBGT.Add arrDanhsach(1), CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
                    
                    strComboHien = strComboHien + arrDanhsach(2) + Chr$(9)
                    If Left(Trim(arrDanhsach(1)), 1) = "1" Then
                        strComboHanghoa = strComboHanghoa + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
                        ' Them thue suat vao combo ten hang hoa 20/9/2010
                        strComboHienHanghoa = strComboHienHanghoa + arrDanhsach(2) + " (" + arrDanhsach(4) + " %)" + Chr$(9)
                        
                    End If
                    If Left(Trim(arrDanhsach(1)), 1) = "2" Then
                        strComboDichvu = strComboDichvu + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
                        ' Them thue suat vao combo ten hang hoa 20/9/2010
                        strComboHienDichvu = strComboHienDichvu + arrDanhsach(2) + " (" + arrDanhsach(4) + " %)" + Chr$(9)
                    End If
                End If
            End If
        Next
        
        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode1 As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_29122007.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode1 In xmlNodeListMap
            strDonvitinh = strDonvitinh + GetAttribute(xmlNode1, "value") + Chr$(9)
        Next
    End If
    
    'Lay ten danh muc dau du file danh muc theo Menu
    strDataFileName = GetCatalogueFileName(blChon)
    
    arrIdHHDVGoc.RemoveAll
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                'Tao mang hang hoa dich vu toan bo cho nhap file text
                arrIdHHDVGoc.Add arrDanhsach(1), CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
                arrIdHHDVBGTGoc.Add arrDanhsach(1), CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1) + Chr$(9)
            End If
        Next
    End If
    
    'Chuyen du lieu vao Combo
    With fps
        '.ActiveSheet = 1
        .Sheet = 1
        mCurrentSheet = .Sheet
        SetDateFormat fps, 1, 61, .ColLetterToNumber("AI"), DDMMYYYY
        .Row = 61
        .Col = .ColLetterToNumber("AI")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        .Col = .ColLetterToNumber("F")
        .Row = 38
        .TypeComboBoxList = strComboHanghoa
        .Row = 42
        .TypeComboBoxList = strComboDichvu
        .Row = 47
        .TypeComboBoxList = strComboHanghoa
        .Row = 51
        .TypeComboBoxList = strComboHanghoa
        .Row = 55
        .TypeComboBoxList = strComboHanghoa
        'Combo Hien
        .Col = .ColLetterToNumber("G")
        .Row = 38
        .TypeComboBoxList = strComboHienHanghoa '
        .Row = 42
        .TypeComboBoxList = strComboHienDichvu '
        .Row = 47
        .TypeComboBoxList = strComboHienHanghoa '
        .Row = 51
        .TypeComboBoxList = strComboHienHanghoa '
        .Row = 55
        .TypeComboBoxList = strComboHienHanghoa '
        .Col = .ColLetterToNumber("P")
        .Row = 38
        .TypeComboBoxList = strDonvitinh '
        .Row = 47
        .TypeComboBoxList = strDonvitinh '
        .Row = 51
        .TypeComboBoxList = strDonvitinh '
        .Row = 55
        .TypeComboBoxList = strDonvitinh '
'       'ActiveSheet = 2
'        .Sheet = 2
'        .Col = .ColLetterToNumber("K")
'        .Row = 37
'        .TypeComboBoxList = strComboHanghoa + strComboDichvu
'        'Combo Hien
'        .Col = .ColLetterToNumber("L")
'        .Row = 37
'        .TypeComboBoxList = strComboHienHanghoa + strComboHienDichvu '
'        'setdate
'        SetDateFormat fps, 2, 37, .ColLetterToNumber("E"), DDMMYYYY
'        SetDateFormat fps, 3, 10, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = sheetKHBS
            mCurrentSheet = .Sheet
            SetDateFormat fps, sheetKHBS, 31, .ColLetterToNumber("BE"), DDMMYYYY
            SetDateFormat fps, sheetKHBS, 46, .ColLetterToNumber("BG"), DDMMYYYY
            
            .Row = 46
            .Col = .ColLetterToNumber("BG")
            .Text = Format(Date, "dd/mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
    End With
    Set xmlNodeListCell = Nothing
    mblKeyF5 = False
    
    
                ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
'Chuc nang: Format lai luoi sao khi chuyen du lieu
'Doi voi to khai:
'- Khoi tao lai theo dinh dang du lieu cu cua to khai
    '- Khoi tao DVT voi mat hang nhom 2
'- Them du lieu cho Combo trong truong hop danh muc da bi xoa tren to khai
'- Them du lieu cho Combo trong truong hop nhap tu file text
    '- Truong hop danh muc da co hang hoa
    '- Truong hop danh muc chua co hang hoa
'- Lay gia tri khoi dau cho bien strGiaTriCu
'- Xoa cong thuc cua o Chu ky
'Doi voi Ban giai trinh:
'- Lay gia tri tong Cot 6 Chenh lech
'- Them du lieu cho Combo trong truong hop danh muc da bi xoa
'- Them du lieu cho Combo trong truong hop nhap tu file text
    '- Truong hop danh muc da co hang hoa
    '- Truong hop danh muc chua co hang hoa
'- Khoa mot so chi tieu doi voi mat hang la nhom 3
    '- Lay gia tri Id
    '- Kiem tra gia tri Id thuoc nhom nao tren to khai
    '- Neu nhom 3 thi khoa, Khong thi mo khoa
'- Nhap lai DVT doi voi nhom 2

Public Sub Prepare2()

Dim i As Integer, intIndexCombo As Integer, iSection As Integer
Dim strLTN As Variant, chkTTDB As Variant, strLTNCu As Variant, strId As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long
Dim tCol As Long, tRow As Long, t As Integer
Dim aGiaTriId As Variant, aGiaTriDVT As Variant, aSection As Integer, currentSheet2 As Integer
Dim GTSua As Variant
Dim count As Long
Dim ngay As Variant
Dim kythang As Variant
Dim kyngay As Variant
    With fps
            'Doi voi to khai
            .Sheet = 1
            currentSheet2 = mCurrentSheet
            mCurrentSheet = .Sheet
            'set to khai chinh thuc hay bao xung
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("Y")
                .Row = 12
                .Text = "1"
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("AA")
                .Row = 12
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("AC")
                .Row = 12
                .Text = "0"
                UpdateCell fps, .Col, .Row, .Text
             ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("Y")
                .Row = 12
                .Text = "0"
                UpdateCell fps, .Col, .Row, .Text
                  
                .Col = .ColLetterToNumber("AA")
                .Row = 12
                .Text = "1"
                UpdateCell fps, .Col, .Row, .Text
                  
                .Col = .ColLetterToNumber("AC")
                .Row = 12
                .Text = StrSolanBosung
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
            
            
            
            'ma NNKD
            If strLoaiNNKD = "0" Then
                strLoaiNNKD = "0000"
            End If
            .Col = .ColLetterToNumber("L")
            .Row = 14
            .Text = strLoaiNNKD
            UpdateCell fps, .Col, .Row, .Text
            'end
            mCurrentSheet = currentSheet2
            mCurrentSheet = .Sheet
            'xu ly ngay thang
            ngay = TAX_Utilities_v1.Day
            .SetText .ColLetterToNumber("S"), 10, ngay
            UpdateCell fps, .ColLetterToNumber("S"), 10, ngay
            
            .SetText .ColLetterToNumber("T"), 10, "x"
            UpdateCell fps, .ColLetterToNumber("T"), 10, "x"
            
            .GetText .ColLetterToNumber("V"), 10, kythang
            .GetText .ColLetterToNumber("AD"), 10, kyngay

            If StrTKThang_PS = "TK_THANG" Then
    '            .Col = .ColLetterToNumber("AT")
    '            .Row = 5
    '            .Text = 1
    '            UpdateCell fps, .Col, .Row, .Text
                .SetText .ColLetterToNumber("P"), 2, kythang
    '            .Col = .ColLetterToNumber("BC")
    '            .Row = 5
    '            .Text = ""
    '            UpdateCell fps, .Col, .Row, .Text
                 .SetText .ColLetterToNumber("T"), 10, ""
                UpdateCell fps, .ColLetterToNumber("T"), 10, ""
                
                .SetText .ColLetterToNumber("AN"), 10, ""
                UpdateCell fps, .ColLetterToNumber("AN"), 10, ""
            ElseIf StrTKThang_PS = "TK_LANPS" Then
    '            .Col = .ColLetterToNumber("BC")
    '            .Row = 5
    '            .Text = 1
    '            UpdateCell fps, .Col, .Row, .Text
                .SetText .ColLetterToNumber("T"), 10, "x"
                UpdateCell fps, .ColLetterToNumber("T"), 10, "x"
                .SetText .ColLetterToNumber("P"), 2, kyngay
    '            .Col = .ColLetterToNumber("AT")
    '            .Row = 5
    '            .Text = ""
    '            UpdateCell fps, .Col, .Row, .Text
                .SetText .ColLetterToNumber("AN"), 10, TAX_Utilities_v1.Day & "/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
                UpdateCell fps, .ColLetterToNumber("AN"), 10, TAX_Utilities_v1.Day & "/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
            End If
      
            'end
            mblKhoiDauToKhai = True
            'Khoi tao lai theo dinh dang du lieu cu cua to khai
            iSection = 1
            Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
            Set xmlCellsNode = xmlCellNode.parentNode
            Do
            If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
            
                i = 1
                .Col = .ColLetterToNumber("C")
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                .Row = iRow
                tRow = iRow
                Do
                    .GetText .ColLetterToNumber("F"), .Row, strLTN
                    .GetText .ColLetterToNumber("N"), 34, chkTTDB
                    .GetText .ColLetterToNumber("E"), .Row, strId
                    .GetText .ColLetterToNumber("R"), .Row, GTSua
                    If chkTTDB = "1" Then
                        If iSection < 3 Then
                         .Col = .ColLetterToNumber("P")
                         .Row = tRow
'                         .Text = ""
'                         .Lock = True
'                         UpdateCell fps, .Col, .Row, .Text
                   
                         .Col = .ColLetterToNumber("S")
                         .Row = tRow
                         .value = 0
                         .Lock = True
                         UpdateCell fps, .Col, .Row, .value
                    
                         .Col = .ColLetterToNumber("V")
                         .Row = tRow
                         .value = 0
                         .Lock = True
                         UpdateCell fps, .Col, .Row, .value
                
                         .Col = .ColLetterToNumber("AA")
                         .Row = tRow
                         .value = 0
                         .Lock = True
                         UpdateCell fps, .Col, .Row, .value
                        Else
                            If GTSua = 1 Then
                                .Col = .ColLetterToNumber("P")
                                .Row = tRow
                                .Lock = False
                            End If
                        End If
                        
                        If iSection < 3 Then
                            .Col = .ColLetterToNumber("AD")
                            .Row = tRow
                            .value = 0
                            UpdateCell fps, .Col, .Row, .value
                        Else
                            .Col = .ColLetterToNumber("AD")
                            .Row = tRow
                            .Text = ""
                            UpdateCell fps, .Col, .Row, 0
                         End If
                    Else
                        .Col = .ColLetterToNumber("P")
                        If Trim(Mid(strLTN, 200, 15)) <> "" Then
'                            .Text = Trim(Mid(strLTN, 210, 10))
'                            UpdateCell fps, .Col, .Row, .Text
                            .Lock = False
                        End If
                        'Khoi tao DVT voi mat hang nhom 2
                        If Trim(Mid(strLTN, 200, 15)) = "" And (iSection = 2 Or GTSua = 1) Then
                            .Lock = False
                        End If
                        If iSection < 3 Then
                            .Col = .ColLetterToNumber("AD")
                            If Trim(Mid(strLTN, 225, 5)) = "" Then
                                .value = 0
                                UpdateCell fps, .Col, .Row, .value
                            Else
                                .value = Trim(Mid(strLTN, 225, 5))
                                UpdateCell fps, .Col, .Row, .value
                            End If
                        Else
                            .Col = .ColLetterToNumber("AD")
                            .Text = ""
                            UpdateCell fps, .Col, .Row, 0
                        End If
                    End If
                    
                    .Row = tRow
                    .Col = iCol
                    'Them du lieu cho Combo trong truong hop danh muc da bi xoa tren to khai
                    .GetText .ColLetterToNumber("O"), .Row, strLTNCu
                    If strLTN = "" And strLTNCu <> "" Then
                        If Not arrIdHHDV.Exists(strId) Then
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            .TypeComboBoxString = strLTNCu
                            .Text = strLTNCu
                            UpdateCell fps, .Col, .Row, .Text
        
                            'Combo Hien
                            .Col = .ColLetterToNumber("G")
                            .TypeComboBoxString = Trim(Mid(strLTNCu, 1, 200))
                            .Text = Trim(Mid(strLTNCu, 1, 200))
                            UpdateCell fps, .Col, .Row, .Text
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            
                            'Cap nhap lai du lieu theo hang hoa
                            .ActiveSheet = 1
                            fps_Change .Col, .Row
                        Else
                        'Truong hop danh muc da co hang hoa
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            .Text = arrIdHHDV.Item(strId)
                            UpdateCell fps, .Col, .Row, .Text
        
                            'Combo Hien
                            .Col = .ColLetterToNumber("G")
                            .Text = Trim(Mid(arrIdHHDV.Item(strId), 1, 200))
                            UpdateCell fps, .Col, .Row, .Text
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            'Cap nhap lai du lieu theo hang hoa
                            mblKhoiDauToKhai = True
                            .ActiveSheet = 1
                            fps_Change .Col, .Row
                        End If
                    End If
                    
                    'Them du lieu cho Combo trong truong hop nhap tu file text
                    .Row = tRow
                    .Col = iCol
                 
                    If (strLTN = "" And strLTNCu = "" And strId <> "") Or (strId <> "" And strId <> Trim(Mid(strLTNCu, 200, 15))) Then
                        If arrIdHHDV.Exists(strId) Then
                            'Truong hop danh muc da co hang hoa
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            .Text = arrIdHHDV.Item(strId)
                            UpdateCell fps, .Col, .Row, .Text
        
                            'Combo Hien
                            .Col = .ColLetterToNumber("G")
                            .Text = Trim(Mid(arrIdHHDV.Item(strId), 1, 200))
                            UpdateCell fps, .Col, .Row, .Text
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            'Cap nhap lai du lieu theo hang hoa
                            mblKhoiDauToKhai = True
                            .ActiveSheet = 1
                            fps_Change .Col, .Row
                        Else
                            'Truong hop danh muc chua co hang hoa
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            .TypeComboBoxString = arrIdHHDVGoc.Item(strId)
                            .Text = arrIdHHDV.Item(strId)
                            UpdateCell fps, .Col, .Row, .Text
        
                            'Combo Hien
                            .Col = .ColLetterToNumber("G")
                            .TypeComboBoxString = Trim(Mid(arrIdHHDVGoc.Item(strId), 1, 200))
                            .Text = Trim(Mid(arrIdHHDVGoc.Item(strId), 1, 200))
                            UpdateCell fps, .Col, .Row, .Text
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            'Cap nhap lai du lieu theo hang hoa nhap tu file text
                            mblKhoiDauToKhai = True
                            .ActiveSheet = 1
                            fps_Change .Col, .Row
                        End If
                    End If
                    .Col = .ColLetterToNumber("C")
                    .Row = i + iRow
                    tRow = .Row
                    i = i + 1
                 Loop Until .Text = "aa"
                 
                 iSection = iSection + 1
            End If
                If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                    Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
                Else
                    Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling
                End If
            Loop Until xmlCellsNode Is Nothing
            
            'Lay gia tri khoi dau cho bien strGiaTriCu
            .GetText .ColLetterToNumber("G"), 38, strGiaTriCu
            
            'Xoa cong thuc cua o Chu ky
            Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.firstChild
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            ' Ho va ten nhan vien dai ly
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            ' So chung chi hanh nghe
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            
            
            'Update lai cot tong
            .Sheet = 1
            .Row = 6
            .Col = .ColLetterToNumber("AN")
            UpdateCell fps, .Col, .Row, .value
            .Row = 11
             .Col = .ColLetterToNumber("V")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AA")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AG")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AJ")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AN")
            UpdateCell fps, .Col, .Row, .value
            
            Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
            Set txmlCellNode = txmlCellNode.parentNode
            Set txmlCellNode = txmlCellNode.parentNode
            Set txmlCellNode = txmlCellNode.nextSibling
            Set txmlCellNode = txmlCellNode.firstChild
            Set txmlCellNode = txmlCellNode.firstChild
            ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
            .Row = tRow
            .Col = .ColLetterToNumber("V")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AA")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AG")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AJ")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AN")
            UpdateCell fps, .Col, .Row, .value
            
    'tinh so dong
        .Col = .ColLetterToNumber("c")
        .Row = 33
        Do
            count = count + 1
            .Row = .Row + 1
        Loop Until .value = "D"
        
            .Sheet = 1
            .Col = .ColLetterToNumber("AI")
            .Row = 12 + count
            If .value = "" Then
'                SetDateFormat fps, 1, 12 + count, .ColLetterToNumber("AJ"), DDMMYYYY
                .Text = Format(Date, "dd/mm/yyyy")
                UpdateCell fps, .Col, .Row, .Text
            End If
            .Refresh
            
            
'             If Not TAX_Utilities_v1.DataKHBS Is Nothing Then
'                Set xmlNode = TAX_Utilities_v1.Data(.SheetCount - 2).childNodes(2).firstChild
'                SetAttribute xmlNode, "DateKHBS", TAX_Utilities_v1.DateKHBS
'                SetAttribute xmlNode, "loaiKHBS", loaiKHBS
'                If loaiKHBS = "frmKHBS_BS" Then
'                    .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0115"), "Msg") & _
'                    GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
'                Else
'                    .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0116"), "Msg") & _
'                    GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
'                End If
'                fillData
'                .EventEnabled(EventAllEvents) = False
'                .Sheet = 1
'                .Row = 10 + count
'                .Col = .ColLetterToNumber("AJ")
'                .Formula = "Header!T5"
'                UpdateCell fps, .Col, .Row, .Text
'                .Sheet = 1
'                .Row = 12 + count
'                .Col = .ColLetterToNumber("AI")
'                .Text = Format(Date, "dd/mm/yyyy")
'                UpdateCell fps, .Col, .Row, .Text
'                .EventEnabled(EventAllEvents) = True
'                .ActiveSheet = .SheetCount - 1
'            End If
'        'dhdang sua chuc nang tai bang ke
'        'ngay 15/02/2011
'        .Sheet = 2 'PL1
'        .ActiveSheet = .Sheet
'            .Row = 3
'            .Col = .ColLetterToNumber("C")
'            .CellType = CellTypeButton
'            .Lock = False
'            .TypeButtonTextColor = RGB(0, 0, 255)
'            .TypeButtonColor = &H8000000F
'            .TypeButtonAlign = TypeButtonAlignRight
'            .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
'        .Sheet = 3 'PL2
'         .ActiveSheet = .Sheet
'            .Row = 1
'            .Col = .ColLetterToNumber("C")
'            .CellType = CellTypeButton
'            .Lock = False
'            .TypeButtonTextColor = RGB(0, 0, 255)
'            .TypeButtonColor = &H8000000F
'            .TypeButtonAlign = TypeButtonAlignRight
'            .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
'
    End With
    mblKhoiDauToKhai = False
End Sub

Private Sub fps_BeforeEditMode(ByVal Col As Long, ByVal Row As Long, ByVal UserAction As FPUSpreadADO.BeforeEditModeActionConstants, CursorPos As Variant, Cancel As Variant)
Dim iRow As Long, iCol As Long
With fps
    If .Sheet = 1 Then
        iRow = .ActiveRow
        iCol = .ActiveCol
        GetCellSpan fps, iCol, iRow
        If iCol = fps.ColLetterToNumber("G") Then
            fps.Col = fps.ColLetterToNumber("G")
            fps.Row = iRow
            strGiaTriCu = fps.Text
        End If
    End If
End With
End Sub

'Chuc nang: thuc hien khi chon Checkbox
'Khi danh dau Checkbox
    'Hien thong bao co muon xoa du lieu cu
    'Khong muon xoa thi thoat ra
    'Khi chon xoa
        'Xoa du lieu va khoa tu cot 3 den 7
        'Khong co tac dung voi mat hang o muc 3
        'Cot DVT xoa trang va khoa
            'Doi voi mat hang muc 2
            'Xoa DVT tuong ung o ban Giai trinh
'Huy danh dau Checkbox
    'Khong co tac dung voi mat hang nhom 3
    'Mo khoa tu cong 3 den cot 7
    'Lay lai gia tri DVT va Thue suat

Public Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim lResult As Integer
    Dim blTemp As Boolean
    Dim iCol As Long, iRow As Long, tCol As Long, tRow As Long
    Dim aGiaTriDVT As Variant, aGiaTriId As Variant, t As Integer
    Dim i As Long, DataText As Variant, strId As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim iSection As Integer, GTSua As Variant
With fps
    .EventEnabled(EventAllEvents) = False
    mCurrentSheet = .ActiveSheet
    Select Case mCurrentSheet
    Case 1
        'thuc hien khi chon Checkbox
        If Col = .ColLetterToNumber("N") And Row = 34 Then
            .Sheet = .ActiveSheet
            .Row = Row
            .Col = Col
            
            If .value = 1 Then 'Khi danh dau Checkbox
                'Hien thong bao co muon xoa du lieu cu
                lResult = DisplayMessage("0034", msYesNo, miQuestion, , mrNo)
                If lResult = 3 Then 'Khong muon xoa thi thoat ra
                    .value = 0
                    UpdateCell fps, .Col, .Row, IIf(.value = 1, "x", vbNullString)
                    
                    .EventEnabled(EventAllEvents) = True
                    Exit Sub
                Else 'Khi chon xoa
                    'Xoa du lieu va khoa tu cot 3 den 7
                    
                    Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
                    Set xmlCellsNode = xmlCellNode.parentNode
                    iSection = 1
                     Do
                        If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                        If iSection < 3 Then 'Khong co tac dung voi mat hang o muc 3
                        Do
                            For i = 1 To xmlCellsNode.childNodes.length
                                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                                Select Case iCol
                                    Case .ColLetterToNumber("P") 'Cot DVT xoa trang va khoa
                                        .Row = iRow
                                        .Col = iCol
                                      '  .Text = ""
                                      '  .Lock = True
'                                        UpdateCell fps, iCol, iRow, .Text
                                        .GetText .ColLetterToNumber("R"), iRow, GTSua
                                        If iSection = 2 Or GTSua = 1 Then 'Doi voi mat hang muc 2
                                            .Row = iRow
                                            .Col = iCol
                                            .Col = .ColLetterToNumber("E")
                                            aGiaTriId = .Text
                                        End If
                                    Case .ColLetterToNumber("S")
                                        .Row = iRow
                                        .Col = iCol
                                        .value = 0
                                        .Lock = True
                                        UpdateCell fps, iCol, iRow, .value
                                    Case .ColLetterToNumber("V")
                                        .Row = iRow
                                        .Col = iCol
                                        .EventEnabled(EventAllEvents) = True
                                        .value = 0
                                        .EventEnabled(EventAllEvents) = False
                                        .Row = iRow
                                        .Col = iCol
                                        .Lock = True
                                        UpdateCell fps, iCol, iRow, .value
                                    Case .ColLetterToNumber("AA")
                                        .Row = iRow
                                        .Col = iCol
                                        .EventEnabled(EventAllEvents) = True
                                        .value = 0
                                        .EventEnabled(EventAllEvents) = False
                                        .Row = iRow
                                        .Col = iCol
                                        .Lock = True
                                        UpdateCell fps, iCol, iRow, .value
                                    Case .ColLetterToNumber("AD")
                                        .Row = iRow
                                        .Col = iCol
                                        .value = 0
                                        UpdateCell fps, iCol, iRow, .value
                                End Select
                            Next
                            If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                            Set xmlCellsNode = xmlCellsNode.nextSibling
                        Loop Until xmlCellsNode Is Nothing
                        End If
                        iSection = iSection + 1
                       End If
                        If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
                        Else
                            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling
                        End If
                    Loop Until xmlCellsNode Is Nothing
                End If
            ElseIf .value = 0 Then 'Huy danh dau Checkbox
                    
                    Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
                    Set xmlCellsNode = xmlCellNode.parentNode
                    iSection = 1
                     Do
                    If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                    If iSection < 3 Then 'Khong co tac dung voi mat hang nhom 3
                        'Mo khoa tu cong 3 den cot 7
                        'Lay lai gia tri DVT va Thue suat
                        Do
                            For i = 1 To xmlCellsNode.childNodes.length
                                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                                .GetText .ColLetterToNumber("F"), iRow, strId
                                Select Case iCol
                                    Case .ColLetterToNumber("P")
                                        .Row = iRow
                                        .Col = iCol
                                        .Lock = False
                                        .GetText .ColLetterToNumber("R"), iRow, GTSua
                                        If Trim(Mid(strId, 210, 10)) <> "" Then
                                            .Text = Trim(Mid(strId, 210, 10))
                                            UpdateCell fps, .Col, .Row, .Text
                                        ElseIf Trim(Mid(strId, 210, 10)) = "" And (iSection = 2 Or GTSua = 1) Then
                                            .Lock = False
                                        End If
                                    Case .ColLetterToNumber("S")
                                        .Row = iRow
                                        .Col = iCol
                                        .Lock = False
                                    Case .ColLetterToNumber("V")
                                        .Row = iRow
                                        .Col = iCol
                                        .Lock = False
                                    Case .ColLetterToNumber("AA")
                                        .Row = iRow
                                        .Col = iCol
                                        .Lock = False
                                    Case .ColLetterToNumber("AD")
                                        .Row = iRow
                                        .Col = iCol
                                        If Trim(Mid(strId, 220, 5)) <> "" Then
                                        .value = Trim(Mid(strId, 220, 5))
                                        UpdateCell fps, .Col, .Row, .value
                                        Else
                                         .value = 0
                                        UpdateCell fps, .Col, .Row, .value
                                       End If
                                End Select
                            Next
                            If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                            Set xmlCellsNode = xmlCellsNode.nextSibling
                            
                        Loop Until xmlCellsNode Is Nothing
                        End If
                        iSection = iSection + 1
                        End If
                        If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
                        Else
                            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling
                        End If
                    Loop Until xmlCellsNode Is Nothing
            End If
        End If
    TAX_Utilities_v1.AdjustData(.Sheet) = True
    End Select
    .EventEnabled(EventAllEvents) = True
End With
End Sub

'Chuc nang: thay doi du lieu tren luoi
'Doi voi to khai
'- Huy chon Combo to khai khong hop le va ra thong bao
'- Huy chon Combo ban giai trinh khong hop le va ra thong bao
'- Lay lai vi tri o duoc chon
'- Khi chon Combo
    '- Xac dinh nhom cua mat hang dang chon
    '- Chuyen sang du lieu tuong ung
    '- Lay gia tri Combo cot G Combo Hienthi
    '- Nhap lai Combo cot F: Combo F
    '- Nhap lai gia tri cot E: Id
    '- Nhap lai gia tri cong D: Ten tai nguyen
    '- Neu cot F khac trong thi chuyen gia tri cho cong O : Gia tri
    '- Lay gia tri Checkbox
    '- Neu Checkbox khong chon
        '- Nhap lai gia tri vao DVT nhom 1, 3
        '- Neu Nhom 2 khong co DVT thi cho nhap, co DVT thi khoa
        '- Hien thi thue suat voi nhom 1, 2
        '- Khong hien thi thue suat voi nhom 3
    '- Khi Checkbox duoc chon
        '- Xoa va khoa DVT nhom 1, 2
        '- Nhom 3 hien DVT
    '- Doi voi Nhom 2 khi nhap DVT
        '- Kiem tra Vitri O de lay aGiaTriDVT, aGiaTriId
        '- Cap nhap lai DVT cho ban Giai trinh
'Doi voi Ban giai trinh
'- Chon Combo va co gia tri
    '- Lay gia tri cua Combo hien thi
    '- Update combo F
    '- Lay gia tri Combo cua cot F
    '- Chuyen gia tri vao cot E: Id
    '- Chuyen gia tri vao cot D: Ten hang hoa
    '- Chuyen gia tri vao cot O: Gia tri
    '- Chuyen gia tri DVT vao cot S va V
                 '- S: So lieu da ke khai
                 '- V: So lieu dieu chinh
'- Kiem tra Id khong co trong to khai hien thong bao va tra lai gia tri cu
'- Chon Combo va co gia tri
    '- Lay gia tri cua Combo hien thi
    '- Update combo F
    '- Lay gia tri Combo cua cot F
    '- Chuyen gia tri vao cot E: Id
    '- Chuyen gia tri vao cot D: Ten hang hoa
    '- Chuyen gia tri vao cot O: Gia tri
    '- Chuyen gia tri DVT vao cot S va V
                 '- S: So lieu da ke khai
                 '- V: So lieu dieu chinh
'- Khoa mot so chi tieu doi voi mat hang la nhom 3
    '- Kiem tra Vitri O
    '- Lay gia tri Id
    '- Kiem tra gia tri Id thuoc nhom nao tren to khai va DVT
    '- Neu nhom 3 thi khoa, 'Khong thi mo khoa
    '- Nhap lai DVT doi voi nhom 2
'- Cap nhap lai cot 6 (Chenh lech)
'- Kiem tra gia tri ky ke khai co hop le MMYYYY
'- Kiem tra Ten hang va ky ke khai da ton tai chua
    '- Truong hop sua ky ke khai
    '- Neu da co, thi thong bao va tro ve gia tri cu
'- Tinh tong cot 6: chenh lech va hien thi len luoi

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant, strLTN As Variant
Dim intIndexCombo As Integer
Dim chkTTDB As Variant
Dim aCol As Long, aRow As Long, iCol As Long, iRow As Long
Dim i As Integer, iSection As Integer, aSection As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim tempTen As Variant, tempKy As Variant, itempTen As Variant, itempKy As Variant
Dim aGiaTriId As Variant, aGiaTriDVT As Variant
Dim GT4 As Variant, GT5 As Variant, GT6 As Variant, GT7 As Variant
Dim GTSua As Variant, GTGBB As Variant
Dim Cot4 As Variant, Cot5 As Variant, Cot6 As Variant
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tRow As Long, tCol As Long, mblLoi As Boolean
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant, varTemp, valTemp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim xmlNodeH15 As MSXML.IXMLDOMNode

fps.EventEnabled(EventAllEvents) = False
    
    With fps
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        Select Case .Sheet
        Case 1
            mblLoi = False
            'Huy chon Combo to khai khong hop le va ra thong bao
            If blGiaTriCu And Col = .ColLetterToNumber("G") Then 'And Not mblKeyF5 Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0086", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0086", msOKOnly, miWarning
                End If
                
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
                blGiaTriCu = False
                mblLoi = True
               
            End If
            'Huy chon Combo ban giai trinh khong hop le va ra thong bao
            If blKiemTraDieuChinhCu And Col = .ColLetterToNumber("G") Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0089", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0089", msOKOnly, miWarning
                End If
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
                blKiemTraDieuChinhCu = False
                mblLoi = True
            End If
            
            'Lay vi tri o duoc chon
            aCol = Col
            aRow = Row
            GetCellSpan fps, aCol, aRow
            .Col = aCol
            .Row = aRow
            
            'Xac dinh nhom cua mat hang dang chon
            iSection = 1
            aSection = 0
            Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
            Set xmlCellsNode = xmlCellNode.parentNode
            Do
                Do
                    For i = 1 To xmlCellsNode.childNodes.length
                        ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                        If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                            If iCol = aCol And iRow = aRow And aSection = 0 Then
                                aSection = iSection
                                Exit For
                            End If
                        End If
                    Next
                    If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                    If aSection > 0 Then Exit Do
                    Set xmlCellsNode = xmlCellsNode.nextSibling
                Loop Until xmlCellsNode Is Nothing
                If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                    Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
                End If
                If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                    iSection = iSection + 1
                End If
            Loop Until xmlCellsNode.parentNode.nextSibling Is Nothing
            
            .Col = aCol
            .Row = aRow
            
            If Col = .ColLetterToNumber("G") Then 'Khi chon Combo
                
                'Chuyen sang du lieu tuong ung
                'Lay gia tri Combo cot G Combo Hienthi
                intIndexCombo = .typeComboboxCurSel
                .Col = .ColLetterToNumber("F")
                .typeComboboxCurSel = intIndexCombo
                'Nhap lai Combo cot F: Combo F
                
                .Col = .ColLetterToNumber("F")
                UpdateCell fps, .Col, .Row, .Text
    
                'Nhap lai gia tri cot E: Id
                .Col = .ColLetterToNumber("E")
                .GetText .ColLetterToNumber("F"), Row, strId
                .Text = Trim(Mid(strId, 200, 15))
                'lay ID cua loai hang
                 valTemp = .Text
                'end
                 UpdateCell fps, .Col, .Row, .Text
                 
                 'Nhap lai gia tri cong D: Ten tai nguyen
                 .Col = .ColLetterToNumber("D")
                 .Text = Trim(Left(strId, 200))
                 UpdateCell fps, .Col, .Row, .Text
                 
                 If strId <> "" Then 'Neu cot F khac trong thi chuyen gia tri cho cong O : Gia tri
                    .Col = .ColLetterToNumber("O")
                    .Text = strId
                    UpdateCell fps, .Col, .Row, .Text
                 End If
                 
                'Nhap lai Gia bao bi
                .Col = .ColLetterToNumber("Q")
                .value = Trim(Mid(strId, 230, 14))
                UpdateCell fps, .Col, .Row, .value
                
                'Cho phep sua

                .Col = .ColLetterToNumber("R")
                .value = Trim(Mid(strId, 245, 1))
                UpdateCell fps, .Col, .Row, .value
                 
                 'Lay gia tri Cot sua
                 .GetText .ColLetterToNumber("R"), aRow, GTSua
                 
                 'Lay gia tri Checkbox
                .GetText .ColLetterToNumber("N"), 34, chkTTDB
                If chkTTDB = "0" Or chkTTDB = "" Then 'Neu Checkbox khong chon
                
                    'Nhap lai gia tri vao DVT nhom 1, 3
                    .Col = .ColLetterToNumber("P")
                    If (Trim(Mid(strId, 200, 15)) <> "" Or aSection = 1 Or aSection >= 3) Then 'And Not mblKeyF5 Then
                        If GTSua <> "1" Or Not mblKeyF5 Then
                            .Text = Trim(Mid(strId, 215, 10))
                            UpdateCell fps, .Col, .Row, .Text
                        End If
                    End If
'                    If CStr(valTemp) = "10700" Or CStr(valTemp) = "10800" Then
'                        .Col = .ColLetterToNumber("S")
'                        .Row = aRow
'                        .Lock = True
'                    End If
                    'Neu Nhom 2 khong co DVT thi cho nhap, co DVT thi khoa
                    If (Trim(Mid(strId, 215, 10)) = "" Or .Text = "") And (aSection = 2 Or GTSua = 1) Then
                        .Lock = False
                    Else
                        .Lock = False
                    End If
                    
                    'Khong tac dung voi nhom 3
                    'Hien thi thue suat
                    If aSection < 3 Then
                        .Col = .ColLetterToNumber("AD")
                        If Trim(Mid(strId, 225, 5)) <> "" Then
                            .value = Trim(Mid(strId, 225, 5))
                        Else
                            .value = 0
                        End If
                        UpdateCell fps, .Col, .Row, .value
                        
                        If Not mblKhoiDauToKhai And Not mblKeyF5 And Not mblLoi Then
                            fps.EventEnabled(EventAllEvents) = True
                            'Tinh lai Gia tinh thue TTDB
                            .GetText .ColLetterToNumber("S"), .Row, GT4
                            .GetText .ColLetterToNumber("V"), .Row, GT5
                            .GetText .ColLetterToNumber("AD"), .Row, GT7
                            .GetText .ColLetterToNumber("Q"), .Row, GTGBB

                            GT6 = CRound((GT5 - (GTGBB * GT4)) / (1 + GT7 / 100), 0)
                            If GT6 < 0 Then
                                GT6 = 0
                            End If
                            .Col = .ColLetterToNumber("AA")
                            .value = GT6
                            .Row = aRow
                            .Col = .ColLetterToNumber("AA")
                            UpdateCell fps, .Col, .Row, .value
                            fps.EventEnabled(EventAllEvents) = False
                        End If
                    Else 'Khong hien thi thue suat
                        .Col = .ColLetterToNumber("AD")
                        .Text = ""
                        UpdateCell fps, .Col, .Row, 0
                    End If

                Else 'Khi Checkbox duoc chon
                    'Khong tac dung voi nhom 3
                    'Xoa va khoa DVT
'                    If aSection < 3 Then
'                        .Col = .ColLetterToNumber("P")
'                        .Text = ""
'                        .Lock = True
'                        UpdateCell fps, .Col, .Row, .Text
'                    Else 'Nhom 3 hien DVT
                        .Col = .ColLetterToNumber("P")
                        If Trim(Mid(strId, 210, 10)) <> "" Or aSection = 1 Or aSection >= 3 Then
                        .Text = Trim(Mid(strId, 210, 10))
                        UpdateCell fps, .Col, .Row, .Text
                            If GTSua = 1 Then
                                .Lock = False
                            Else
                                .Lock = False
                                
                            End If
                        End If
                    'End If
                    
                    'Nhap lai Gia bao bi
                    .Col = .ColLetterToNumber("Q")
                    .value = 0
                    UpdateCell fps, .Col, .Row, .value
                    
                    'Cho phep sua
                    If aSection < 3 Then
                        .Col = .ColLetterToNumber("R")
                        .value = 0
                        UpdateCell fps, .Col, .Row, .value
                    End If
                    'Khong tac dung voi nhom 3
                    'Xoa thue sua
                    If aSection < 3 Then
                        .Col = .ColLetterToNumber("AD")
                        .value = 0
                        UpdateCell fps, .Col, .Row, .value
                    Else 'Nhom 3 thi Huy thue suat
                        .Col = .ColLetterToNumber("AD")
                        .Text = ""
                        UpdateCell fps, .Col, .Row, 0
                    End If
                    If aSection < 3 Then
                        .Col = .ColLetterToNumber("V")
                        .value = 0
                        .Lock = True
                        UpdateCell fps, .Col, .Row, .Text
                        
                        .Col = .ColLetterToNumber("AA")
                        .value = 0
                        .Lock = True
                        UpdateCell fps, .Col, .Row, .Text
                    End If
                End If
            End If
            
            'Doi voi nhom 1, 2 va cot 4, 5
            If (aCol = .ColLetterToNumber("S") Or aCol = .ColLetterToNumber("V")) And aSection < 3 And aSection > 0 Then
                If Not mblKhoiDauToKhai And Not mblKeyF5 And Not mblLoi Then
'                    Tinh lai Gia tinh thue TTDB
                    
                    fps.EventEnabled(EventAllEvents) = True

                    .GetText .ColLetterToNumber("S"), .Row, GT4
                    .GetText .ColLetterToNumber("V"), .Row, GT5
                    .GetText .ColLetterToNumber("AD"), .Row, GT7
                    .GetText .ColLetterToNumber("Q"), .Row, GTGBB

                    GT6 = CRound((GT5 - (GTGBB * GT4)) / (1 + GT7 / 100), 0)
                    If GT6 < 0 Then
                        GT6 = 0
                    End If
                    .Col = .ColLetterToNumber("AA")
                    .value = GT6
                    .Col = .ColLetterToNumber("AA")
                    .Row = aRow
                    UpdateCell fps, .Col, .Row, .value

                    fps.EventEnabled(EventAllEvents) = False
                End If
            End If
            
            'Doi voi Nhom 2 khi nhap DVT
            If aCol = .ColLetterToNumber("P") Then
                'Kiem tra Vitri O de lay aGiaTriDVT, aGiaTriId
                .Row = aRow
                aGiaTriDVT = .Text
                .Col = .ColLetterToNumber("E")
                aGiaTriId = .Text
                If aGiaTriId <> "" Then
                    .Sheet = 1
                    mCurrentSheet = 1
                End If
            End If
        
            'Tinh lai cot cong cho truong hop nhap tu file
            If mblKhoiDauToKhai = True Then
                .Row = aRow
                .Col = .ColLetterToNumber("AN")
                UpdateCell fps, .ColLetterToNumber("AN"), aRow, .value
                
                fps.EventEnabled(EventAllEvents) = True
                Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))

                Set txmlCellNode = txmlCellNode.parentNode
                Set txmlCellNode = txmlCellNode.parentNode
                Set txmlCellNode = txmlCellNode.parentNode.lastChild
                Set txmlCellNode = txmlCellNode.previousSibling
                Set txmlCellNode = txmlCellNode.firstChild
                Set txmlCellNode = txmlCellNode.firstChild
                ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
                .Row = tRow
                .Col = .ColLetterToNumber("V")
                UpdateCell fps, .Col, .Row, .value
                .Col = .ColLetterToNumber("AA")
                UpdateCell fps, .Col, .Row, .value
                .Col = .ColLetterToNumber("AG")
                UpdateCell fps, .Col, .Row, .value
                .Col = .ColLetterToNumber("AJ")
                UpdateCell fps, .Col, .Row, .value
                .Col = .ColLetterToNumber("AN")
                UpdateCell fps, .Col, .Row, .value
                fps.EventEnabled(EventAllEvents) = False
            End If
'      Case 2
'      .Col = Col
'      .Row = Row
'         If Col = .ColLetterToNumber("L") Then 'Khi chon Combo
''           Dim strId As Variant
'                    'Chuyen sang du lieu tuong ung
'                    'Lay gia tri Combo cot G Combo Hienthi
'                    intIndexCombo = .typeComboboxCurSel
'                    .Col = .ColLetterToNumber("K")
'                    .typeComboboxCurSel = intIndexCombo
'                    'Nhap lai Combo cot F: Combo F
'                    Trim (Mid(strId, 200, 10))
'                    .Col = .ColLetterToNumber("K")
'                    UpdateCell fps, .Col, .Row, .Text
'
'                    .Col = .ColLetterToNumber("I")
'                    .GetText .ColLetterToNumber("K"), Row, strId
'                    .Text = Trim(Mid(strId, 200, 10))
'                     UpdateCell fps, .Col, .Row, .Text
'          End If
'      Case 3
      
      
      'KHBS
'        Case .SheetCount - 1
'             Dim cValueOld As Variant
'             Dim cellid As String
'             Dim xmlCellNodeData As MSXML.IXMLDOMNode
'            ' Dim i As Integer
'                 fps.EventEnabled(EventAllEvents) = False
'                  cCol = .Col
'                  cRow = .Row
'                 .Col = .ColLetterToNumber("BD")
'                 .Row = .Row
'                 .GetText .ColLetterToNumber("BC"), .Row, varTemp
'
'           If varTemp <> "" Then
'              Dim arr() As String
'                For i = 0 To UBound(arrDanhSachKHBS)
'                   If varTemp = Left(arrDanhSachKHBS(i), InStr(arrDanhSachKHBS(i), "~") - 1) Then
'                       cellid = "AN_" & Right(arrDanhSachKHBS(i), Len(arrDanhSachKHBS(i)) - InStr(arrDanhSachKHBS(i), "~") - 2)
'                      For Each xmlCellNodeData In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
'                           If GetAttribute(xmlCellNodeData, "CellID") = cellid Then
'                               cValue = GetAttribute(xmlCellNodeData, "Value")
'                                Exit For
'                            End If
'                       Next
'                   End If
'                Next
'                .Col = .ColLetterToNumber("BF")
'                .Row = Row
'                cTagF = .CellTag
'               If cTagF <> "" Then
'                      For Each xmlCellNodeData In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
'                           If GetAttribute(xmlCellNodeData, "CellID") = Right(cTagF, Len(cTagF) - InStr(cTagF, "~")) Then
'                               cValueF = GetAttribute(xmlCellNodeData, "Value")
'                                Exit For
'                            End If
'                       Next
'                End If
'                  .Col = .ColLetterToNumber("BF")
'                  .Row = Row
'                  .CellTag = .Row & "~" & cellid
'                   .value = cValue
'                   UpdateCell fps, .Col, .Row, .value
'                   .Col = .ColLetterToNumber("BG")
'                   .Row = Row
'                   UpdateCell fps, .Col, .Row, .value
'                   .Col = .ColLetterToNumber("BH")
'                   .Row = Row
'                   UpdateCell fps, .Col, .Row, .value
'
'
'                .Sheet = .SheetCount - 1
'                Set xmlNodeH15 = TAX_Utilities_v1.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                     ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
'                    .Col = cCol
'                    .Row = cRow
'                    UpdateCell fps, .Col, .Row, .value
'                    If .value < 0 Then
'                        .Col = .ColLetterToNumber("BE")
'                        .Row = cRow + 3
'                        .value = 0
'                        UpdateCell fps, .Col, .Row, 0
'                     Else
'                        .Col = .ColLetterToNumber("BE")
'                        .Row = cRow + 3
'                        UpdateCell fps, .Col, .Row, .value
'                    End If
'                fps.EventEnabled(EventAllEvents) = True
'             End If
     End Select
      
            mblKeyF5 = False
            mblLoi = False
    End With
    fps.EventEnabled(EventAllEvents) = True
End Sub

'Chuc nang: thay doi gia tri cua Combo
'Doi voi to khai
'- Xac dinh nhom cua mat hang
'- Kiem tra mat hang da duoc chon chua tren to khai va ban dieu chinh
    '- kiem tra trung Ten TN trong to khai
    '- Neu co thi dat co blGiaTriCu = True
    '- kiem tra da khai trong ban dieu chinh
    '- neu trung thi dat co blKiemTraDieuChinhCu = True
'Doi voi ban diem chinh
'- kiem tra ten hang va ky da duoc nhap chua
    'neu da duoc nhap thi dat co blDieuChinhCu = True

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim intIndexCombo As Integer
Dim temp As Variant, itemp As Variant, iRow As Long, iCol As Long, iColRow As String
Dim tempTen As Variant, tempKy As Variant, itempTen As Variant, itempKy As Variant, itempId As Variant
Dim i As Integer, aSection As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim aCol As Long, aRow As Long, iSection As Integer

Dim intIndexCuc As Integer, intIndexChiCuc As Integer
Dim varIdCucThue As Variant
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode
Dim arrDanhsach() As String
Dim tenCucThue As String
Dim maCucThue As String
Dim tenChiCucThue As String
Dim maChiCucThue As String
With fps

    .Sheet = .ActiveSheet
    mCurrentSheet = .Sheet
    Select Case .Sheet
        Case 1
        'Lay vi tri O ban dau
        aCol = .ActiveCol
        aRow = .ActiveRow
        GetCellSpan fps, aCol, aRow
            
        'Xac dinh nhom cua mat hang
        iSection = 1
        aSection = 0
        Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
        Set xmlCellsNode = xmlCellNode.parentNode
        Do
            Do
                For i = 1 To xmlCellsNode.childNodes.length
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                    If iCol = aCol And iRow = aRow And aSection = 0 Then
                        aSection = iSection
                        Exit For
                    End If
                Next
                If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                If aSection > 0 Then Exit Do
                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing
            If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
            End If
            If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                iSection = iSection + 1
            End If
        Loop Until xmlCellsNode.parentNode.nextSibling Is Nothing
        
        'Kiem tra mat hang da duoc chon chua tren to khai va ban dieu chinh
        Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, Col, Row))
        Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        blGiaTriCu = False
        blKiemTraDieuChinhCu = False
        If Col = fps.ColLetterToNumber("G") Then
            'kiem tra trung Ten TN trong to khai
            fps.Col = Col
            fps.Row = Row
            temp = fps.Text
'            Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 13))
            Set xmlCellsNode = xmlCellNode.parentNode
'            Do
'            If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
            
            i = 1
            .Col = .ColLetterToNumber("C")
            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
            .Row = iRow
          
            Do
                 .Col = .ColLetterToNumber("G")
                 itemp = .Text
                 If .Row <> Row Then
                    If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
                        RowGiaTriCu = Row
                        ColGiaTriCu = .ColLetterToNumber("G")
                        .Row = Row
                        .Col = .ColLetterToNumber("G")
                        blGiaTriCu = True
                        Exit Sub
                    End If
                 End If
                 .Col = .ColLetterToNumber("C")
                 .Row = i + iRow
                 i = i + 1
            Loop Until .Text = "aa"
            .Sheet = 1
            mCurrentSheet = 1
        End If
        Case 2
'         If Col = .ColLetterToNumber("L") Then 'Khi chon Combo
'           Dim strId As Variant
'                    'Chuyen sang du lieu tuong ung
'                    'Lay gia tri Combo cot G Combo Hienthi
'                    intIndexCombo = .TypeComboBoxCurSel
'                    .Col = .ColLetterToNumber("K")
'                    .TypeComboBoxCurSel = intIndexCombo
'                    'Nhap lai Combo cot F: Combo F
'                    Trim (Mid(strId, 200, 10))
'                    .Col = .ColLetterToNumber("K")
'                    UpdateCell fps, .Col, .Row, .Text
'
'                    .Col = .ColLetterToNumber("I")
'                    .GetText .ColLetterToNumber("K"), Row, strId
'                    .Text = Trim(Mid(strId, 200, 10))
'                     UpdateCell fps, .Col, .Row, .Text
'          End If
        Case 3
'KHBS
'        Case .SheetCount - 1
'                    'lay ve dong combobox
'            Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, .ColLetterToNumber("BI"), 5))
'            Set xmlCellNode = xmlCellNode.parentNode
'            Set xmlCellNode = xmlCellNode.parentNode
'            Set xmlCellNode = xmlCellNode.parentNode.lastChild
'            Set xmlCellNode = xmlCellNode.firstChild
'            Set xmlCellNode = xmlCellNode.firstChild
'            For i = 1 To 6
'                Set xmlCellNode = xmlCellNode.nextSibling
'            Next i
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'
'
'            .Sheet = .SheetCount - 1
'            mCurrentSheet = .SheetCount - 1
'
'            .Col = Col
'            .Row = Row
'
'            If Col = .ColLetterToNumber("BE") And Row = iRow Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber("BI")
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber("BI")
'                UpdateCell fps, .Col, .Row, .Text
'                ' Lay Id cua cuc thue vua chon o tren
'                .GetText .Col, .Row, varIdCucThue
'
'                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
'                    ' Lay thong tin cho chi cuc thue
'                    varIdCucThue = Left$(varIdCucThue, 3)
'
'                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
'                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'                        For Each xmlNode In xmlNodeListCell
'                            If GetAttribute(xmlNode, "Value") <> "" Then
'                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'
'                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
'                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
'                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
'                                End If
'                            End If
'                        Next
'                        Set xmlDomData = Nothing
'                        Set xmlNodeListCell = Nothing
'                        Set xmlNode = Nothing
'                    End If
'
'                    .Row = iRow + 2
'                    .Col = .ColLetterToNumber("BE")
'                    .TypeComboBoxList = tenChiCucThue
'
'                    .Row = iRow + 2
'                    .Col = .ColLetterToNumber("BI")
'                    .TypeComboBoxList = maChiCucThue
'                End If
'            End If
'
'            If Col = .ColLetterToNumber("BE") And Row = iRow + 2 Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexChiCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber("BI")
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexChiCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber("BI")
'                UpdateCell fps, .Col, .Row, .Text
'            End If
        End Select
End With
End Sub


'Chuc nang: thay doi du lieu khi ban phim xuong
'Doi voi to khai
'- Xoa ten hang trong Combo
    '- kiem tra Ten hang tren to khai da khai tren ban giai trinh
    '- xac dinh nhom cua ten hang
    '- neu da co thi hien thong bao va khong cho phep xoa
    '- neu chua co thi cho phep xoa
        '- xoa Combo hien thi
        '- xom Combo F
        '- xoa Id
        '- xoa Ten hang
        '- xoa DVT
        '- xoa Gia tri
        '- xoa Thue suat
'- Xoa hang nhan F6
    '- kiem tra Ten hang da khai tren Ban dieu chinh
    '- neu co thi thong bao va huy xoa
'Doi voi ban dieu chinh
'- xoa ten hang tren Combo
    '- xoa Combo hien thi
    '- xom Combo F
    '- xoa Id
    '- xoa Ten hang
    '- xoa hang DVT
    '- mo khoa cac hang duoc nhap
            'S: So lieu da ke khai
            'V: So lieu dieu chinh
            'AD: Ghi chu
    
Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
'Dim i As Integer, iCol As Long, iRow As Long, aCol As Long, aRow As Long, vitri As Long
'Dim oldCol As Long
'Dim iSection As Integer, aSection As Integer
'Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'Dim ActCol As Long, ActRow As Long, itempTen As Variant
'Dim strMa As Variant
'With fps
'    'xac dinh vi tri hien tai
'    ActCol = fps.ActiveCol
'    ActRow = fps.ActiveRow
'    GetCellSpan fps, ActCol, ActRow
'    .Sheet = .ActiveSheet
'    mCurrentSheet = .Sheet
'        iSection = 1
'        aSection = 0
'        Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 13))
'        Set xmlCellsNode = xmlCellNode.parentNode
'        Do
'            Do
'                For i = 1 To xmlCellsNode.childNodes.length
'                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
'                    If iCol = ActCol And iRow = ActRow And aSection = 0 Then
'                        aSection = iSection
'                        Exit For
'                    End If
'                Next
'                If xmlCellsNode.nextSibling Is Nothing Then Exit Do
'                If aSection > 0 Then Exit Do
'                Set xmlCellsNode = xmlCellsNode.nextSibling
'            Loop Until xmlCellsNode Is Nothing
'            If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
'                Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
'            End If
'            If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
'                iSection = iSection + 1
'            End If
'        Loop Until xmlCellsNode.parentNode.nextSibling Is Nothing
'
'        If Keycode = vbKeyDelete And ActCol = .ColLetterToNumber("G") Then
'        'Xoa ten hang trong Combo
'        'kiem tra Ten hang tren to khai da khai tren ban giai trinh
'            blKiemTraDieuChinhCuDelete = False
'            .Sheet = 2
'            mCurrentSheet = 2
'
'            If aSection < 3 Then
'                  Set xmlCellNode = TAX_Utilities_v1.Data(2 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 8))
'                  Set xmlCellsNode = xmlCellNode.parentNode
'
'                  Do
'                      ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(0), "CellID"), iCol, iRow
'                      .Row = iRow
'                      .Col = .ColLetterToNumber("G")
'                      itempTen = .Text
'                      If strGiaTriCu = itempTen And itempTen <> "" Then
'                          'neu da co thi hien thong bao va khong cho phep xoa
'                          .Sheet = 1
'                          mCurrentSheet = 1
'                           DisplayMessage "0089", msOKOnly, miWarning
'                          .SetFocus
'                          .Row = ActRow
'                          .Col = .ColLetterToNumber("G")
'                          Exit Function
'                      End If
'                      Set xmlCellsNode = xmlCellsNode.nextSibling
'                  Loop Until xmlCellsNode Is Nothing
'            End If
'        'neu chua co thi cho phep xoa
'        .Sheet = 1
'        mCurrentSheet = 1
'            fps.EventEnabled(EventAllEvents) = False
'                  'xoa Combo hien thi
'                  .Row = ActRow
'                  .Col = ActCol
'                  .Text = ""
'                  UpdateCell fps, .Col, .Row, .Text
'                  'xom Combo F
'                  .Row = ActRow
'                  .Col = .ColLetterToNumber("F")
'                  .Text = ""
'                  UpdateCell fps, .Col, .Row, .Text
'
'                  'Chon lai combo cua colF
'                  'xoa Id
'                  .Col = .ColLetterToNumber("E")
'                  .Row = ActRow
'                  .Text = ""
'                  UpdateCell fps, .Col, .Row, .Text
'                  'xoa Ten hang
'                  .Col = .ColLetterToNumber("D")
'                  .Row = ActRow
'                  .Text = ""
'                  UpdateCell fps, .Col, .Row, .Text
'                  'xoa DVT
'                  .Col = .ColLetterToNumber("P")
'                  .Row = ActRow
'                  .Text = ""
'                  .Col = .ColLetterToNumber("P")
'                  .Row = ActRow
'                  If aSection <> 2 Then
'                    .Lock = True
'                  End If
'                  UpdateCell fps, .Col, .Row, .Text
'                  'xoa Gia bao bi
'                  .Col = .ColLetterToNumber("Q")
'                  .Row = ActRow
'                  .Value = 0
'                  UpdateCell fps, .Col, .Row, .Text
'
'                   'Cho phep sua
'                  .Col = .ColLetterToNumber("R")
'                  .Row = ActRow
'                  .Value = 0
'                  UpdateCell fps, .Col, .Row, .Text
'                  'xoa Gia tri
'                  .Col = .ColLetterToNumber("O")
'                  .Row = ActRow
'                  .Text = ""
'                  UpdateCell fps, .Col, .Row, .Text
'
'                  .EventEnabled(EventChange) = True
'
'                  'xoa Thue suat
'                  .Row = ActRow
'                  .Col = .ColLetterToNumber("AD")
'                  UpdateCell fps, .Col, .Row, 0
'                  .Row = ActRow
'                  If aSection < 3 Then
'                    .Value = 0
'                  Else
'                    .Text = ""
'                  End If
'
'                  'tro ve vi tri Combo
'                  .Row = ActRow
'                  .Col = .ColLetterToNumber("G")
'            End If
'
'            If aSection < 3 Then
'                .Row = ActRow
'                .Col = .ColLetterToNumber("G")
'                strMa = .Text
'        End If
'    fps.EventEnabled(EventAllEvents) = True
'End With

DeleteEnable = True
End Function
Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

'Chuc nang: thay doi du lieu khi nhan phim len
'them hoac xoa hang
'Doi voi to khai
'- huy voi truong hop la Checkbox va Chu ky
'- xac dinh nhom cua ten hang
'- them hang
    '- Nhom 1: nhap lai Combo hang cho dong moi
    '- Nhom 2: nhap lai Combo dich vu cho dong moi
    '- Nhom 3: nhap lai Combo hang cho dong moi
'- tao lai STT cho nhom do
'Doi voi Ban giai trinh
'- them hang
    '- Lay gia tri Combo cho mat hang moi tu DOM
    '- nhap lai Combo hang cho dong moi
'- xoa hang
    '- Mo khoa lai format khi xoa hang dau tien
    '- Tinh lai gia tri hang tong so
    '- thuc hien su kien thay doi du lieu
    
Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i, j As Integer, iCol As Long, iRow As Long, aCol As Long, aRow As Long, vitri As Long
Dim oldCol As Long
Dim iSection As Integer, aSection As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim ActCol As Long, ActRow As Long, itempTen As Variant
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tRow As Long, tCol As Long

    ActCol = fps.ActiveCol
    ActRow = fps.ActiveRow
    GetCellSpan fps, ActCol, ActRow
    If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
    'them hoac xoa hang
        With fps
            .Sheet = .ActiveSheet
            mCurrentSheet = .Sheet
            Select Case .Sheet
                Case 1
                'Huy voi truong hop la Checkbox va Chu ky
                Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, ActCol, ActRow))
         
                Set xmlCellNode = xmlCellNode.parentNode
                Set xmlCellNode = xmlCellNode.parentNode
                If GetAttribute(xmlCellNode, "Dynamic") = "0" Then
                    .Sheet = 1
                    .Row = ActRow
                    .Col = ActCol
                    Exit Sub
                End If
                
                .Row = ActRow
                .Col = ActCol
                
                If KeyCode = vbKeyF5 Then
                    .Row = ActRow
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .ColLetterToNumber("AN"), ActRow, .value
                    If ActCol = .ColLetterToNumber("G") Or ActCol = .ColLetterToNumber("S") Or ActCol = .ColLetterToNumber("V") Then
                        mblKeyF5 = True
                    End If
                    fps_Change ActCol, ActRow
                End If
                
                'Dat lai gia tri cu cho cot G
                .Row = ActRow
                .Col = ActCol
                
                If KeyCode = vbKeyF6 Then
                    If ActCol = .ColLetterToNumber("G") Then
                        .Col = ActCol
                        .Row = ActRow
                        strGiaTriCu = .Text
                        blGiaTriCu = False
                        fps_Change ActCol, ActRow
                    End If
                End If
                
                .Row = ActRow
                .Col = ActCol
                
                aCol = .ColLetterToNumber("G")
                oldCol = .Col
                vitri = .Row
                aRow = .Row
               
                'xac dinh nhom cua ten hang
                iSection = 1
                aSection = 0
                Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
                Set xmlCellsNode = xmlCellNode.parentNode
                Do
                    Do
                        For i = 1 To xmlCellsNode.childNodes.length
                            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                            If iCol = aCol And iRow = aRow And aSection = 0 Then
                                aSection = iSection
                                Exit For
                            End If
                        Next
                        If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                        If aSection > 0 Then Exit Do
                        Set xmlCellsNode = xmlCellsNode.nextSibling
                    Loop Until xmlCellsNode Is Nothing
                    If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                        Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
                    End If
                    If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                        iSection = iSection + 1
                    End If
                Loop Until xmlCellsNode.parentNode.nextSibling Is Nothing
                
                'them hang
                If KeyCode = vbKeyF5 Then
                    Select Case aSection
                        Case 1 'Nhom 1: nhap lai Combo hang cho dong moi
                            .Col = .ColLetterToNumber("F")
                            .Row = aRow + 1
                            .TypeComboBoxList = strComboHanghoa
                            .Col = .ColLetterToNumber("G")
                            .TypeComboBoxList = strComboHienHanghoa
                        Case 2 'Nhom 2: nhap lai Combo dich vu cho dong moi
                            .Col = .ColLetterToNumber("F")
                            .Row = aRow + 1
                            .TypeComboBoxList = strComboDichvu
                            .Col = .ColLetterToNumber("G")
                            .TypeComboBoxList = strComboHienDichvu
                        Case 3, 4, 5 'Nhom 3: nhap lai Combo hang cho dong moi
                            .Col = .ColLetterToNumber("F")
                            .Row = aRow + 1
                            .TypeComboBoxList = strComboHanghoa
                            .Col = .ColLetterToNumber("G")
                            .TypeComboBoxList = strComboHienHanghoa '
                    End Select
                End If
                
                'Tinh lai gia tri hang tong so
                If KeyCode = vbKeyF6 Or KeyCode = vbKeyF5 Then
                    .Row = 6
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .Col, .Row, .value
                    .Row = 36
                     .Col = .ColLetterToNumber("V")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AA")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AG")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AJ")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .Col, .Row, .value
                    
                    Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
                    Set txmlCellNode = txmlCellNode.parentNode
                    Set txmlCellNode = txmlCellNode.parentNode
                    Set txmlCellNode = txmlCellNode.nextSibling
                    Set txmlCellNode = txmlCellNode.firstChild
                    Set txmlCellNode = txmlCellNode.firstChild
                    ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = .ColLetterToNumber("V")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AA")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AG")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AJ")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .Col, .Row, .value
                    
                    Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
                    Set txmlCellNode = txmlCellNode.parentNode
                    Set txmlCellNode = txmlCellNode.parentNode
                    Set txmlCellNode = txmlCellNode.nextSibling
                    Set txmlCellNode = txmlCellNode.nextSibling
                    Set txmlCellNode = txmlCellNode.nextSibling
                    Set txmlCellNode = txmlCellNode.firstChild
                    Set txmlCellNode = txmlCellNode.firstChild
                    ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = .ColLetterToNumber("V")
                    UpdateCell fps, .Col, .Row, .value
                
                    fps.EventEnabled(EventAllEvents) = True
                    Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
    
                    Set txmlCellNode = txmlCellNode.parentNode
                    Set txmlCellNode = txmlCellNode.parentNode
                    Set txmlCellNode = txmlCellNode.parentNode.lastChild
                    Set txmlCellNode = txmlCellNode.previousSibling
                    Set txmlCellNode = txmlCellNode.firstChild
                    Set txmlCellNode = txmlCellNode.firstChild
                    ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = .ColLetterToNumber("V")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AA")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AG")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AJ")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .Col, .Row, .value
                    fps.EventEnabled(EventAllEvents) = False
                End If
        
                .SetActiveCell iCol, iRow
                
                'tao lai STT cho nhom do
                Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, aCol, aRow))
                Set xmlCellsNode = xmlCellNode.parentNode
               
                Set xmlCellsNode = xmlCellsNode.parentNode.firstChild
                Set xmlCellNode = xmlCellsNode.firstChild
                ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
                i = 1
                .Col = .ColLetterToNumber("C")
                .Row = iRow
                Do
                    If aSection <= 5 Then
                        .Text = str(i)
                    End If
                    .Col = .ColLetterToNumber("C")
                    .Row = i + iRow
                    i = i + 1
                 Loop Until .Text = "aa"
                .SetActiveCell oldCol, vitri
'                .Refresh
'            Case 2
''Stt
'                i = 1
'                .Col = .ColLetterToNumber("B")
'                .Row = 37
'                Do
'                     .Text = str(i)
'                     .Col = .ColLetterToNumber("B")
'                     UpdateCell fps, .Col, .Row, .value
'                     .Row = i + 37
'                     i = i + 1
'                Loop Until .Text = "aa"
'                .SetActiveCell iCol, iRow
'            Case 3
' 'stt
'                i = 1
'                j = 1
'                .Col = .ColLetterToNumber("B")
'                .Row = 10
'                Do
'                     .Text = str(i)
'                     .Col = .ColLetterToNumber("B")
'                     UpdateCell fps, .Col, .Row, .value
'                     .Row = i + 10
'                     i = i + 1
'                Loop Until .Text = "aa"
'
'                .Row = 17 + i + j
'                Do
'                     .Text = str(j - 1)
'                     .Col = .ColLetterToNumber("B")
'                     UpdateCell fps, .Col, .Row, .value
'                     .Row = j + i + 17
'                     j = j + 1
'                Loop Until .Text = "aa"
'                .SetActiveCell iCol, iRow
'             Case .SheetCount - 1
'                 If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
'                 fps.EventEnabled(EventAllEvents) = False
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 34
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("BE")
'                         .Text = "10"
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 9
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    i = 1
'                   .Row = .Row + 3
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("BE")
'                         .Text = "10"
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "bb"
'                    .SetActiveCell iCol, iRow
'                    .Row = iRow
'                    .Col = iCol
'                    .SetActiveCell iCol, iRow
'                End If
'                 fps.EventEnabled(EventAllEvents) = True
            End Select
        End With
    End If
    
    
    
    fps.EventEnabled(EventAllEvents) = True
End Sub

'Chuc nang: chuyen con tro sang O khac
'Doi voi to khai
'- Lay gia tri Combo vua chuyen den
'Doi voi Ban giai trinh
'- Lay gia tri Combo va Ky ke khai vua chuyen den
'- Goi ham CellChange
    '- Kiem tra gia tri Ky ke khai theo dinh dang

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim varTemp As Variant
Dim iCol As Long, iRow As Long
Dim count As Long
With fps
    .Sheet = .ActiveSheet
    mCurrentSheet = .Sheet
   .EventEnabled(EventAllEvents) = False
        'Lay gia tri Combo vua chuyen den
        Select Case .Sheet
        Case 1
            iRow = NewRow
            iCol = NewCol
            GetCellSpan fps, iCol, iRow
            If iCol = fps.ColLetterToNumber("G") Then
                fps.Col = fps.ColLetterToNumber("G")
                fps.Row = iRow
                strGiaTriCu = fps.Text
            End If
    'tinh so dong
        .Col = .ColLetterToNumber("C")
        .Row = 33
        Do
            count = count + 1
            .Row = .Row + 1
        Loop Until .value = "D"
    'check date
            If Col = .ColLetterToNumber("AI") And Row = 16 + count Then
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
                UpdateCell fps, .Col, .Row, .Text
            End If
            
'        Case 2
'    'check date
'            If Col = .ColLetterToNumber("E") Then
'                .GetText Col, Row, varTemp
'                If varTemp <> "" And varTemp <> "../../...." Then
'                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    Else
'                        .SetFocus
'                        .SetActiveCell Col, Row
'                    End If
'                Else
'                 .SetText Col, Row, ""
'                End If
'               .Col = Col
'               .Row = Row
'                UpdateCell fps, .Col, .Row, .Text
'            End If
'
'            'chuyen cot ky hieu ve chu cai in hoa
'            If Col = .ColLetterToNumber("C") Then
'                 .GetText Col, Row, varTemp
'                 .SetText Col, Row, UCase(varTemp)
'                 .Col = Col
'                 .Row = Row
'                 UpdateCell fps, .Col, .Row, .Text
'            End If
            
'        Case 3
'            'tinh so dong
'            .Col = .ColLetterToNumber("B")
'            .Row = 10
'            count = 0
'            Do While .value <> "aa"
'                count = count + 1
'                .Row = .Row + 1
'            Loop
'            If Col = .ColLetterToNumber("E") And Row < 10 + count Then
'                .GetText Col, Row, varTemp
'                If varTemp <> "" And varTemp <> "../../...." Then
'                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    Else
'                        .SetFocus
'                        .SetActiveCell Col, Row
'                    End If
'                Else
'                 .SetText Col, Row, ""
'                End If
'                .Col = Col
'                .Row = Row
'                UpdateCell fps, .Col, .Row, .Text
'            End If
'
'            'chuyen cot ky hieu ve chu cai in hoa
'            If Col = .ColLetterToNumber("C") And Row < 10 + count Then
'                 .GetText Col, Row, varTemp
'                 .SetText Col, Row, UCase(varTemp)
'                 .Col = Col
'                 .Row = Row
'                 UpdateCell fps, .Col, .Row, .Text
'            End If
            
        End Select
    .EventEnabled(EventAllEvents) = True
End With
End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        'update sheet KHBS
        UpdateSheets
        ' end
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
'Chuc nang: Kiem tra loi truoc khi in
'Doi voi to khai
'- Khong nhap ten hang ma nhap chi tiet thi bao loi
'- Bat loi voi nhom 1, 2
    '- Nhap Chi tiet ma khong nhap Doanh so ban thi bao loi khong chon Checkbox
    '- Checkbox duoc chon, ma khong nhap Doanh so ban thi bao loi
    '- Nhap Chi tiet ma khong nhap Gia tinh thue thi bao loi khong chon Checkbox
    '- Checkbox duoc chon, ma khong nhap Gia tinh thue thi bao loi
    '- Neu nhap Doanh so ban va Gia tinh thue ma Doanh so ban nho hon thi bao loi
'Doi voi ban dieu chinh
'- Ky dieu chinh lon hon hoac bang ky tren to khai thi bao loi
'- Neu nhap ky dieu chinh ma khong chon ten hang thi bao loi
'- Chon ten hang ma khong nhap ky dieu chinh thi bao loi
'- Khong ke khai ten hang va ky dieu chinh ma co chi tiet thi bao loi
'- Tao mang chua tong Chenh lech dieu chinh theo tung mat hang
'- Neu nhap Gia trinh thue ma khong nhap Thue suat hoac nguoc lai thi bao loi
'- Dieu chinh hang tren ky khai ma khac tong dieu chinh mat hang tren ban giai trinh thi bao loi
Private Sub CheckDynamicError()
    ' Sheet 1
    
    Dim LoiTonTai As Boolean
    
    Dim chkTTDB As Variant
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim iCol As Long, iRow As Long, iSection As Integer
       
    Dim vErrMsgV13 As Variant, vErrMsgAA13 As Variant, vErrMsgAA13_2 As Variant, vErrMsgG13 As Variant, vErrMsgAJ13 As Variant
    Dim vErrMsgE13 As Variant
     
    Dim ColV13 As Long, ColAA13 As Long, ColAA13_2 As Long, ColG13 As Long
    Dim ColS13 As Long, ColAG13 As Long, ColAJ13 As Long, ColP13 As Long
    Dim ColE13 As Long
    Dim ColLoiV13 As Long, ColLoiAA13 As Long, ColLoiAA13_2 As Long, ColLoiG13 As Long, ColLoiAJ13 As Long
    
    
    Dim RowV13 As Long, RowAA13 As Long, RowAA13_2 As Long, RowG13 As Long, RowAJ13 As Long, RowP13 As Long
    Dim RowE13 As Long
    Dim RowLoiV13 As Long, RowLoiAA13 As Long, RowLoiAA13_2 As Long, RowLoiG13 As Long, RowLoiAJ13 As Long
        
    Dim GT5 As Variant, GT6 As Variant, GT2 As Variant, GT3 As Variant
    Dim GT4 As Variant, GT8 As Variant, GT9 As Variant, GTMa As Variant
        
    Dim XacDinhLoiV13 As Boolean, XacDinhLoiAA13 As Boolean, XacDinhLoiAA13_2 As Boolean, XacDinhLoiG13 As Boolean, XacDinhLoiAJ13 As Boolean
    Dim XacDinhLoiE13 As Boolean
    
    Dim i As Integer, iCurrentSheet As Integer, totalRow As Integer
    
    Dim iTongChenhLech As Double, strIdChenhLech As String
     
    Dim arrIdHangHoaDieuChinh As New Scripting.Dictionary
    
    Dim varAA13, varAD13, varAG13, varAJ13 As Variant
    Dim errorAG13 As Variant
    Dim counterr, countS2, countS3, dateErrS2, dateErrS3, nameErrS2, nameErrS3, iFlagvNgayHD As Integer
    ' Khai bao cac bien cho du lieu cua Sheet 2
    Dim s2Nullerr, s2Dateerr, s2Datevl, s2Namevl, s2kyhieu, s2so, s2tenkh, s2tenhang, s2soluong, s2dongia As Variant
    ' Khai bao cac bien cho du lieu cua Sheet 3
    Dim s3Nullerr, s3Dateerr, s3Datevl, s3Namevl, s3kyhieu, s3so, s3nguyenlieu, s3soluong, s3thuedanop As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer, iFlagNgayTT As Integer
    Dim iflagTongGT As Variant, errTongGT As Variant, errNgayTinhThue As Variant
    'sheet 4
    Dim errLoiGT As Variant
    
    With fps

    iCurrentSheet = .Sheet
    .Sheet = .SheetCount

    CheckErrorMST
    .Sheet = iCurrentSheet

    .EventEnabled(EventAllEvents) = False
    XacDinhLoiV13 = False
    XacDinhLoiAA13 = False
    XacDinhLoiAA13_2 = False
    XacDinhLoiG13 = False
    XacDinhLoiE13 = False
     i = 1
     
     ColE13 = .ColLetterToNumber("E")
     ColV13 = .ColLetterToNumber("V")
     ColP13 = .ColLetterToNumber("P")
     ColAA13 = .ColLetterToNumber("AA")
     ColAA13_2 = .ColLetterToNumber("AA")
     ColG13 = .ColLetterToNumber("G")
     ColS13 = .ColLetterToNumber("S")
     ColAG13 = .ColLetterToNumber("AG")
     ColAJ13 = .ColLetterToNumber("AJ")
     
     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("E"), 16, vErrMsgV13
     .GetText .ColLetterToNumber("E"), 17, vErrMsgAA13
     .GetText .ColLetterToNumber("E"), 18, vErrMsgAA13_2
     .GetText .ColLetterToNumber("E"), 19, vErrMsgG13
     '.GetText .ColLetterToNumber("E"), 30, vErrMsgE13
     .GetText .ColLetterToNumber("E"), 20, errorAG13
     .GetText .ColLetterToNumber("E"), 21, s2Nullerr
     .GetText .ColLetterToNumber("E"), 21, s3Nullerr
     .GetText .ColLetterToNumber("E"), 24, errLoiGT
     .GetText .ColLetterToNumber("E"), 25, errTongGT
     .GetText .ColLetterToNumber("B"), 25, iflagTongGT
    .GetText .ColLetterToNumber("E"), 26, errNgayTinhThue
    
    ' Doi voi to khai Sheet 1
    LoiTonTai = False
    
    mCurrentSheet = 1
    .Sheet = 1
    iSection = 1
    
    ' Lay gia tri Checkbox
    .GetText .ColLetterToNumber("N"), 34, chkTTDB
     
    Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'bien luu tong so dong
    totalRow = 0
    
    Do
        If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
            i = 1
            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
            .Row = iRow
            arrIdHangHoa.RemoveAll
                   
            Do
                .GetText ColE13, .Row, GTMa
                .GetText ColG13, .Row, GT2
                .GetText ColP13, .Row, GT3
                .GetText ColS13, .Row, GT4
                .GetText ColV13, .Row, GT5
                .GetText ColAA13, .Row, GT6
                .GetText ColAG13, .Row, GT8
                .GetText ColAJ13, .Row, GT9
                .GetText .ColLetterToNumber("AA"), .Row, varAA13
                .GetText .ColLetterToNumber("AD"), .Row, varAD13
                .GetText .ColLetterToNumber("AJ"), .Row, varAJ13
                .GetText .ColLetterToNumber("AG"), .Row, varAG13
                
                
                ' Trung ten hang hoa, dich vu
                If Trim(GTMa) <> "" Then
                    If arrIdHangHoa.Exists(GTMa) Then
                        .Col = ColG13
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgE13
                        .BackColor = mErrorColor
                        If XacDinhLoiE13 = False And LoiTonTai = False Then
                            ColLoiG13 = ColG13
                            RowLoiG13 = .Row
                            LoiTonTai = True
                            XacDinhLoiE13 = True
                        End If
                    Else
                        arrIdHangHoa.Add GTMa, .value
                        .Col = ColG13
                        If .CellNote = vErrMsgE13 Then
                            .Col = ColG13
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    End If
                End If
                 ' Khong nhap ten hang ma nhap chi tiet thi bao loi
                 If GT2 = "" And (GT3 <> "" Or GT4 > 0 Or GT5 > 0 Or GT6 > 0 Or GT8 > 0 Or GT9 <> 0) Then
                        .Col = ColG13
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgG13
                        .BackColor = mErrorColor
                        If XacDinhLoiG13 = False And LoiTonTai = False Then
                            ColLoiG13 = ColG13
                            RowLoiG13 = .Row
                            LoiTonTai = True
                            XacDinhLoiG13 = True
                        End If
                 Else
                        .Col = ColG13
                        If .CellNote = vErrMsgG13 Then
                            .Col = ColG13
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                 End If
                
                ' Bat loi voi nhom 1, 2
                If iSection < 3 Then
                
                    ' Nhap Chi tiet ma khong nhap Doanh so ban thi bao loi khong chon Checkbox
                    ' Checkbox duoc chon, ma khong nhap Doanh so ban thi bao loi
                    If ((GT3 <> "" Or GT4 > 0 Or GT8 > 0 Or GT9 <> 0 Or GT6 > 0) And GT5 = 0 And (chkTTDB = "0" Or chkTTDB = "")) Or ((chkTTDB = "0" Or chkTTDB = "") And GT5 = 0 And GT2 <> "") Then
                        .Col = ColV13
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgV13
                        .BackColor = mErrorColor
                        If XacDinhLoiV13 = False And LoiTonTai = False Then
                            ColLoiV13 = ColV13
                            RowLoiV13 = .Row
                            LoiTonTai = True
                            XacDinhLoiV13 = True
                        End If
                    Else
                        .Col = ColV13
                        If .CellNote = vErrMsgV13 Then
                         .Col = ColV13
                         .CellNote = ""
                         .BackColor = mNonErrorColor
                        End If
                    End If
                    'Nhap Chi tiet ma khong nhap Gia tinh thue thi bao loi khong chon Checkbox
                    'Checkbox duoc chon, ma khong nhap Gia tinh thue thi bao loi
                    If ((GT3 <> "" Or GT4 > 0 Or GT8 > 0 Or GT9 <> 0 Or GT5 > 0) And GT6 = 0 And (chkTTDB = "0" Or chkTTDB = "")) Or ((chkTTDB = "0" Or chkTTDB = "") And GT6 = 0 And GT2 <> "") Then
                        .Col = ColAA13
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgAA13
                        .BackColor = mErrorColor
                        If XacDinhLoiAA13 = False And LoiTonTai = False Then
                            ColLoiAA13 = ColAA13
                            RowLoiAA13 = .Row
                            LoiTonTai = True
                            XacDinhLoiAA13 = True
                        End If
                    Else
                        .Col = ColAA13
                        If .CellNote = vErrMsgAA13 Then
                         .Col = ColAA13
                         .CellNote = ""
                         .BackColor = mNonErrorColor
                        End If
                    End If
                    
                    'Neu nhap Doanh so ban va Gia tinh thue ma Doanh so ban nho hon thi bao loi
                    If GT6 > 0 And GT5 > 0 And GT6 >= GT5 Then
                        .Col = ColAA13_2
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgAA13_2
                        .BackColor = mErrorColor
                        If XacDinhLoiAA13_2 = False And LoiTonTai = False Then
                            ColLoiAA13_2 = ColAA13_2
                            RowLoiAA13_2 = .Row
                            LoiTonTai = True
                            XacDinhLoiAA13_2 = True
                        End If
                    Else
                        .Col = ColAA13_2
                        If .CellNote <> vErrMsgAA13 And .CellNote = vErrMsgAA13_2 Then
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                         End If
                    End If
                        'Khau tru` khong the > phan phai nop
                        If varAA13 * varAD13 / 100 + varAJ13 < varAG13 Then
                            .Col = .ColLetterToNumber("AG")
                            .CellNote = errorAG13
                            .BackColor = mErrorColor
                            counterr = counterr + 1
                        Else
                            .Col = .ColLetterToNumber("AG")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            LoiTonTai = False
                        End If
                 End If
                 
                 .Row = i + iRow
                 .Col = .ColLetterToNumber("C")
                 i = i + 1
            Loop Until .Text = "aa"
            
            totalRow = .Row
            
            iSection = iSection + 1
            End If
        If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
        Else
            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling
        End If
    Loop Until xmlCellsNode Is Nothing
    
'    If XacDinhLoiE13 Then
'        .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), 30, IIf(XacDinhLoiE13, "0", "1")
'        Dim strFormula As String
'        .Sheet = 1
'        strFormula = "'" & .SheetName & "'!" & Trim(.ColNumberToLetter(CLng(ColLoiG13))) & Trim(str(RowLoiG13))
'        .Sheet = .SheetCount
'        .Col = .ColLetterToNumber("C")
'        .Row = 17
'        .Formula = strFormula
'    Else
'        .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), 30, "1"
'    End If
    Dim strFormula As String
    If XacDinhLoiAA13 Then
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 17, IIf(XacDinhLoiAA13, "0", "1")
        .Sheet = 1
        strFormula = "'" & .SheetName & "'!" & Trim(.ColNumberToLetter(CLng(ColLoiAA13))) & Trim(str(RowLoiAA13))
        .Sheet = .SheetCount
        .Col = .ColLetterToNumber("C")
        .Row = 17
        .Formula = strFormula
    Else
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 17, "1"
    End If
    
    If XacDinhLoiV13 Then
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 16, IIf(XacDinhLoiV13, "0", "1")
        .Sheet = 1
        strFormula = "'" & .SheetName & "'!" & Trim(.ColNumberToLetter(CLng(ColLoiV13))) & Trim(str(RowLoiV13))
        .Sheet = .SheetCount
        .Col = .ColLetterToNumber("C")
        .Row = 16
        .Formula = strFormula
    Else
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 16, "1"
    End If
    
    If XacDinhLoiAA13_2 Then
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 18, IIf(XacDinhLoiAA13_2, "0", "1")
    
        .Sheet = 1
        strFormula = "'" & .SheetName & "'!" & Trim(.ColNumberToLetter(CLng(ColLoiAA13_2))) & Trim(str(RowLoiAA13_2))
        .Sheet = .SheetCount
        .Col = .ColLetterToNumber("C")
        .Row = 18
        .Formula = strFormula
    Else
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 18, "1"
    End If
    
     If XacDinhLoiG13 Then
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 19, IIf(XacDinhLoiG13, "0", "1")
    
        .Sheet = 1
        strFormula = "'" & .SheetName & "'!" & Trim(.ColNumberToLetter(CLng(ColLoiG13))) & Trim(str(RowLoiG13))
        .Sheet = .SheetCount
        .Col = .ColLetterToNumber("C")
        .Row = 19
        .Formula = strFormula
    Else
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 19, "1"
    End If
    
    If counterr > 0 Then
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 20, "0"
        .Sheet = 1
    Else
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 20, "1"
        .Sheet = 1
    End If
    
    ' Kiem tra cac rang buoc du lieu cua Sheet 2 (Bang ke HDDV hang hoa Ban ra)
'    .Sheet = 2
'
'    If .SheetVisible = True Then
'         iFlagNgayTT = 0
'        .Row = 37
'        .Col = .ColLetterToNumber("B")
'        Do While .value <> "aa"
'            countS2 = countS2 + 1
'            .Row = .Row + 1
'        Loop
'
'        For i = 37 To 36 + countS2
'        .Row = i
'            .GetText .ColLetterToNumber("C"), i, s2kyhieu
'            .GetText .ColLetterToNumber("D"), i, s2so
'            .GetText .ColLetterToNumber("E"), i, s2Datevl
'            .GetText .ColLetterToNumber("F"), i, s2tenkh
'            .GetText .ColLetterToNumber("J"), i, s2tenhang
'            .GetText .ColLetterToNumber("L"), i, s2Namevl
'            .GetText .ColLetterToNumber("N"), i, s2soluong
'            .GetText .ColLetterToNumber("O"), i, s2dongia
'            If s2kyhieu = "" And s2so = "" And s2Datevl = "" And s2tenkh = "" And s2Namevl = "" And s2soluong = 0 And s2dongia = 0 Then
'                .Col = .ColLetterToNumber("C")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("D")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("E")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("F")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("J")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("L")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("N")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("O")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            Else
'                If s2Datevl = "" Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = s2Nullerr
'                    .BackColor = mErrorColor
'                    dateErrS2 = dateErrS2 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'
'                    If CInt(Right(s2Datevl, 4)) > CInt(TAX_Utilities_v1.Year) Then
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                        .BackColor = mErrorColor
'                        dateErrS2 = dateErrS2 + 1
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                     ElseIf CInt(Right(s2Datevl, 4)) = CInt(TAX_Utilities_v1.Year) And CInt(Right("00" & Mid(s2Datevl, 4, 2), 2)) > CInt(TAX_Utilities_v1.Month) Then
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = errNgayTinhThue
'                        .BackColor = mErrorColor
'                        dateErrS2 = dateErrS2 + 1
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        iFlagNgayTT = 1
'                    End If
'                End If
'                If s2Namevl = "" Then
'                    .Col = .ColLetterToNumber("L")
'                    .CellNote = s2Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS2 = nameErrS2 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("L")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'
'                If s2kyhieu = "" Then
'                    .Col = .ColLetterToNumber("C")
'                    .CellNote = s2Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS2 = nameErrS2 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("C")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s2so = "" Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNote = s2Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS2 = nameErrS2 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("D")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s2tenkh = "" Then
'                    .Col = .ColLetterToNumber("F")
'                    .CellNote = s2Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS2 = nameErrS2 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("F")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s2tenhang = "" Then
'                    .Col = .ColLetterToNumber("J")
'                    .CellNote = s2Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS2 = nameErrS2 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("J")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
''                If s2soluong = "0" Then
''                    .Col = .ColLetterToNumber("N")
''                    .CellNote = s2Nullerr
''                    .BackColor = mErrorColor
''                    nameErrS2 = nameErrS2 + 1
''                    strFocusSheetName = "'" & .SheetName & "'"
''                    strFocusCol = str(.Col)
''                    strFocusRow = str(.Row)
''                Else
''                    .Col = .ColLetterToNumber("N")
''                    .CellNote = ""
''                    .BackColor = mNonErrorColor
''                    LoiTonTai = False
''                End If
''                If s2dongia = "0" Then
''                    .Col = .ColLetterToNumber("O")
''                    .CellNote = s2Nullerr
''                    .BackColor = mErrorColor
''                    nameErrS2 = nameErrS2 + 1
''                    strFocusSheetName = "'" & .SheetName & "'"
''                    strFocusCol = str(.Col)
''                    strFocusRow = str(.Row)
''                Else
''                    .Col = .ColLetterToNumber("O")
''                    .CellNote = ""
''                    .BackColor = mNonErrorColor
''                    LoiTonTai = False
''                End If
'            End If
'        Next
'    End If
    
    
'    ' Kiem tra cac rang buoc du lieu cua Sheet 3 (Bang ke HDDV hang hoa Mua vao)
'    .Sheet = 3
'    If .SheetVisible = True Then
'        Dim giaTriCT10 As Double
'        Dim giaTriCT7 As Double
'        Dim iFlagChiTieu As Integer
'        Dim valueCT7 As Variant
'        Dim count4 As Integer
'        Dim valueCT8S1 As Double
'        Dim strTemp As Variant
'
'        iFlagChiTieu = 0
'        giaTriCT10 = 0
'        giaTriCT7 = 0
'        .Row = 10
'        .Col = .ColLetterToNumber("B")
'        Do While .value <> "aa"
'            .Col = .ColLetterToNumber("K")
'            .GetText .ColLetterToNumber("K"), .Row, strTemp
'            giaTriCT10 = giaTriCT10 + Val(strTemp)
'            .Col = .ColLetterToNumber("B")
'            countS3 = countS3 + 1
'            .Row = .Row + 1
'        Loop
'        .Col = .ColLetterToNumber("B")
'        .Row = 19 + countS3
'        Do
'            .GetText .ColLetterToNumber("K"), .Row, valueCT7
'            giaTriCT7 = giaTriCT7 + Val(valueCT7)
'            .Col = .ColLetterToNumber("B")
'            .Row = .Row + 1
'        Loop Until .Text = "aa"
'            count4 = .Row + 1
'        For i = 10 To 9 + countS3
'        .Row = i
'            .GetText .ColLetterToNumber("C"), i, s3kyhieu
'            .GetText .ColLetterToNumber("D"), i, s3so
'            .GetText .ColLetterToNumber("E"), i, s3Datevl
'            .GetText .ColLetterToNumber("F"), i, s3nguyenlieu
'            .GetText .ColLetterToNumber("G"), i, s3soluong
'            .GetText .ColLetterToNumber("H"), i, s3thuedanop
'
'
'
'            If s3kyhieu = "" And s3so = "" And s3Datevl = "" And s3nguyenlieu = "" And s3soluong = 0 And s3thuedanop = 0 Then
'                .Col = .ColLetterToNumber("C")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("D")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("E")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("F")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("G")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                .Col = .ColLetterToNumber("H")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            Else
'                If s3Datevl = "" Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = s3Nullerr
'                    .BackColor = mErrorColor
'                    dateErrS3 = dateErrS3 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                    If CInt(Right(s3Datevl, 4)) > CInt(TAX_Utilities_v1.Year) Then
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = GetAttribute(GetMessageCellById("0112"), "Msg") 'static
'                        .BackColor = mErrorColor
'                        dateErrS3 = dateErrS3 + 1
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                     ElseIf CInt(Right(s3Datevl, 4)) = CInt(TAX_Utilities_v1.Year) And CInt(Right("00" & Mid(s3Datevl, 4, 2), 2)) > CInt(TAX_Utilities_v1.Month) Then
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = GetAttribute(GetMessageCellById("0112"), "Msg") 'static
'                        .BackColor = mErrorColor
'                        dateErrS3 = dateErrS3 + 1
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                     ElseIf CInt(Right(s3Datevl, 4)) = CInt(TAX_Utilities_v1.Year) And CInt(TAX_Utilities_v1.Month) - CInt(Right("00" & Mid(s3Datevl, 4, 2), 2)) > 2 Then
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = errNgayTinhThue  'static
'                        .BackColor = mErrorColor
'                        dateErrS3 = dateErrS3 + 1
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        iFlagNgayTT = 1
'                    End If
'                End If
'                If s3kyhieu = "" Then
'                    .Col = .ColLetterToNumber("C")
'                    .CellNote = s3Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS3 = nameErrS3 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("C")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s3so = "" Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNote = s3Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS3 = nameErrS3 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("D")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s3nguyenlieu = "" Then
'                    .Col = .ColLetterToNumber("F")
'                    .CellNote = s3Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS3 = nameErrS3 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                Else
'                    .Col = .ColLetterToNumber("F")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s3soluong = "0" Then
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = s3Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS3 = nameErrS3 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'
'                Else
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'                If s3thuedanop = "0" Then
'                    .Col = .ColLetterToNumber("H")
'                    .CellNote = s3Nullerr
'                    .BackColor = mErrorColor
'                    nameErrS3 = nameErrS3 + 1
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'
'                Else
'                    .Col = .ColLetterToNumber("H")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                    LoiTonTai = False
'                End If
'            End If
'        Next
'        'check chi tieu
'        If giaTriCT7 > giaTriCT10 Then
'            .Col = .ColLetterToNumber("K")
'            .Row = count4
'            .BackColor = mErrorColor
'            .CellNote = errLoiGT
'            iFlagChiTieu = 1
'        Else
'            .Col = .ColLetterToNumber("K")
'            .Row = count4
'            .BackColor = mNonErrorColor
'            .CellNote = ""
'        End If
'        .Sheet = 1
'
''        valueCT8S1 = Val(.Text)
'        If iflagTongGT = "0" Then
'            .Col = .ColLetterToNumber("AG")
'            .Row = totalRow + 1
'            .BackColor = mErrorColor
'            .CellNote = errTongGT
'            'iflagTongGT = 2 khi su dung PL01_2 va` co loi
'            iflagTongGT = "2"
'        Else
'            .Col = .ColLetterToNumber("AG")
'            .Row = totalRow + 1
'            .BackColor = mNonErrorColor
'            .CellNote = ""
'        End If
'
'
'    End If
    ' Sau khi valid du lieu Sheet 2 (Bang ke ban ra) va Sheet 3 (Bang ke mua vao) thi moi set lai gia tri cho Header
         .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), 24, IIf(iFlagChiTieu = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 26, IIf(iFlagNgayTT = 1, "0", "1")
        If (iflagTongGT <> "2") Then
            .SetText .ColLetterToNumber("B"), 25, "1"
        End If
        If (dateErrS2 <> 0 Or nameErrS2 <> 0 Or dateErrS3 <> 0 Or nameErrS3 <> 0) Then
             .Sheet = .SheetCount
             .Row = 21
             .Col = .ColLetterToNumber("B")
             .value = "0"
             ' setcell for focus
              strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
             .Col = .ColLetterToNumber("C")
             .Formula = strFormula
    
        Else
             .Sheet = .SheetCount
             .Row = 21
             .Col = .ColLetterToNumber("B")
             .value = "1"
        End If
    
    .Sheet = iCurrentSheet
    End With
End Sub

Sub CheckErrorMST()
    Dim vError1 As Variant, vError2 As Variant, vError3 As Variant
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    With fps
        
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 12, vError1
        .GetText .ColLetterToNumber("E"), 13, vError2
        .GetText .ColLetterToNumber("E"), 14, vError3
        
        .GetText .ColLetterToNumber(SxMST1Col), SxMST1Row, MST1
        .GetText .ColLetterToNumber(SxMST2Col), SxMST2Row, MST2
        .GetText .ColLetterToNumber(SxMST3Col), SxMST3Row, MST3
        .GetText .ColLetterToNumber(SxMST4Col), SxMST4Row, MST4
        .GetText .ColLetterToNumber(SxMST5Col), SxMST5Row, MST5
        .GetText .ColLetterToNumber(SxMST6Col), SxMST6Row, MST6
        .GetText .ColLetterToNumber(SxMST7Col), SxMST7Row, MST7
        .GetText .ColLetterToNumber(SxMST8Col), SxMST8Row, MST8
        .GetText .ColLetterToNumber(SxMST9Col), SxMST9Row, MST9
        .GetText .ColLetterToNumber(SxMST10Col), SxMST10Row, MST10
        .GetText .ColLetterToNumber(SxMST11Col), SxMST11Row, MST11
        .GetText .ColLetterToNumber(SxMST12Col), SxMST12Row, MST12
        .GetText .ColLetterToNumber(SxMST13Col), SxMST13Row, MST13
        
        strCheck = CheckTaxCode(MST1, MST2, MST3, MST4, MST5, MST6, MST7, MST8, MST9, MST10, MST11, MST12, MST13)
        iFlagTaxCode1 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)
        iFlagTaxCode2 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)
        iFlagTaxCode3 = CInt(strCheck)
        
        If iFlagTaxCode1 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 12
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 12
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 13
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 13
            .Text = "1"
        End If
        
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 14
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 14
            .Text = "1"
        End If
        
        .Sheet = 1
        .Col = .ColLetterToNumber("F")
        .Row = 5
        .CellNote = ""
        .BackColor = mFormColor
        If iFlagTaxCode1 = 1 Then
            .CellNote = .CellNote & "> " & vError1
            .BackColor = mErrorColor
        End If
        If iFlagTaxCode2 = 1 Then
            .CellNote = .CellNote & IIf(Trim(.CellNote) = "", "", vbCrLf) & "> " & vError2
            .BackColor = mErrorColor
        End If
        If iFlagTaxCode3 = 1 Then
            .CellNote = .CellNote & IIf(Trim(.CellNote) = "", "", vbCrLf) & "> " & vError3
            .BackColor = mErrorColor
        End If
        .Sheet = iCurrentSheet
    End With
End Sub

Public Sub SetActiveSheet()
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
'             Case "02"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
            Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetKHBS - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetKHBS - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        'SetAttribute xmlNodeValid, "Active", "0"
                        'FlagThaydoi = "1"
                    End If
        End Select
    Next
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean

End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Chuc nang: Huy thong bao loi
'Doi voi to khai
'- Bat loi voi nhom 1, 2
    '- Nhap Doanh so ban ma khong nhap Gia tinh thue thi bao loi khong chon Checkbox
    '- Checkbox duoc chon, neu chon ten hang ma khong nhap Gia tinh thue thi bao loi
    '- Nhap Gia tinh thue ma khong nhap Doanh so ban thi bao loi khong chon Checkbox
    '- Checkbox duoc chon, neu chon ten hang ma khong nhap Doanh so ban thi bao loi
    '- Neu nhap Doanh so ban va Gia tinh thue ma Doanh so ban nho hon thi bao loi
'- Khong nhap ten hang ma nhap chi tiet thi bao loi
Public Sub SetData()
'sheet 1
    Dim chkTTDB As Variant
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim iCol As Long, iRow As Long, iSection As Integer
       
    Dim vErrMsgV13 As Variant, vErrMsgAA13 As Variant, vErrMsgAA13_2 As Variant, vErrMsgG13 As Variant, vErrMsgAJ13 As Variant
    
    Dim ColV13 As Long, ColAA13 As Long, ColAA13_2 As Long, ColG13 As Long
    Dim ColS13 As Long, ColAG13 As Long, ColAJ13 As Long, ColP13 As Long
    Dim ColLoiV13 As Long, ColLoiAA13 As Long, ColLoiAA13_2 As Long, ColLoiG13 As Long, ColLoiAJ13 As Long
    
    Dim RowV13 As Long, RowAA13 As Long, RowAA13_2 As Long, RowG13 As Long, RowAJ13 As Long, RowP13 As Long
    Dim RowLoiV13 As Long, RowLoiAA13 As Long, RowLoiAA13_2 As Long, RowLoiG13 As Long, RowLoiAJ13 As Long
    
    Dim GT5 As Variant, GT6 As Variant, GT2 As Variant, GT3 As Variant
    Dim GT4 As Variant, GT8 As Variant, GT9 As Variant
        
    Dim XacDinhLoiV13 As Boolean, XacDinhLoiAA13 As Boolean, XacDinhLoiAA13_2 As Boolean, XacDinhLoiG13 As Boolean, XacDinhLoiAJ13 As Boolean
    
    Dim i As Integer, iCurrentSheet As Integer
    
    'Sheet 2
    Dim strtempKy As String, inttempThang As Integer, inttempNam As Integer
    
    Dim vErrMsgS2P8 As Variant, vErrMsgS2S12 As Variant, vErrMsgS2S13 As Variant
    Dim vErrMsgS2G8 As Variant, vErrMsgS2P8_2 As Variant, vErrMsgS2G8_3 As Variant
    Dim vErrMsgS2E8 As Variant, vErrMsgS2P8_3 As Variant, vErrMsgS2AA15 As Variant
    
    Dim ColS2P8 As Long, ColS2S12 As Long, ColS2S13 As Long, ColS2V12 As Long, ColS2V13 As Long
    Dim ColS2G8 As Long, ColS2E8 As Long, ColS2AA15 As Long
    
    Dim ColLoiS2P8 As Long, ColLoiS2S12 As Long, ColLoiS2S13 As Long, ColLoiS2V12 As Long, ColLoiS2V13 As Long
    Dim ColLoiS2G8 As Long, ColLoiS2P8_2 As Long, ColLoiS2G8_3 As Long, ColLoiS2AA15 As Long
    
    Dim RowS2P8 As Long, RowS2S12 As Long, RowS2S13 As Long, RowS2V12 As Long, RowS2V13 As Long
    Dim RowS2G8 As Long, RowS2E8 As Long, RowS2AA15 As Long
    
    Dim RowLoiS2P8 As Long, RowLoiS2S12 As Long, RowLoiS2S13 As Long, RowLoiS2V12 As Long, RowLoiS2V13 As Long
    Dim RowLoiS2G8 As Long, RowLoiS2P8_2 As Long, RowLoiS2G8_3 As Long, RowLoiS2AA15 As Long
    
    Dim GTS23 As Variant, GTS2Ma As Variant
    Dim GTS2G8 As Variant, GTS2P8 As Variant
    Dim GTS2S10 As Variant, GTS2V10 As Variant, GTS2AD10 As Variant
    Dim GTS2S11 As Variant, GTS2V11 As Variant, GTS2AD11 As Variant
    Dim GTS2S12 As Variant, GTS2V12 As Variant, GTS2AD12 As Variant
    Dim GTS2S13 As Variant, GTS2V13 As Variant, GTS2AD13 As Variant
    Dim GTS2S14 As Variant, GTS2V14 As Variant, GTS2AD14 As Variant
    Dim GTS2S15 As Variant, GTS2V15 As Variant, GTS2AD15 As Variant
    Dim GTS2AA15 As Variant
        
    Dim XacDinhLoiS2P8 As Boolean, XacDinhLoiS2S12 As Boolean, XacDinhLoiS2S13 As Boolean, XacDinhLoiS2V12 As Boolean, XacDinhLoiS2V13 As Boolean
    Dim XacDinhLoiS2G8 As Boolean, XacDinhLoiS2P8_2 As Boolean, XacDinhLoiS2G8_3 As Boolean
    Dim XacDinhLoiS2E8 As Boolean, XacDinhLoiS2AA15 As Boolean
    
    Dim iTongChenhLech As Double, strIdChenhLech As String
     
    Dim arrIdHangHoaDieuChinh As New Scripting.Dictionary
        
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    
    Dim lCol As Long, lRow As Long, IsUpdate As Boolean
    
    With fps

    iCurrentSheet = .Sheet
    .Sheet = .SheetCount

    CheckErrorMST
    .Sheet = iCurrentSheet

    .EventEnabled(EventAllEvents) = False
    XacDinhLoiV13 = False
    XacDinhLoiAA13 = False
    XacDinhLoiAA13_2 = False
    XacDinhLoiG13 = False
     i = 1
     
     ColV13 = .ColLetterToNumber("V")
     ColP13 = .ColLetterToNumber("P")
     ColAA13 = .ColLetterToNumber("AA")
     ColAA13_2 = .ColLetterToNumber("AA")
     ColG13 = .ColLetterToNumber("G")
     ColS13 = .ColLetterToNumber("S")
     ColAG13 = .ColLetterToNumber("AG")
     ColAJ13 = .ColLetterToNumber("AJ")
     
     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("E"), 16, vErrMsgV13
     .GetText .ColLetterToNumber("E"), 17, vErrMsgAA13
     .GetText .ColLetterToNumber("E"), 18, vErrMsgAA13_2
     .GetText .ColLetterToNumber("E"), 19, vErrMsgG13
    
    'Doi voi to khai
    'Sheet 1
    
    mCurrentSheet = 1
    .Sheet = 1
    iSection = 1
    
    'Lay gia tri Checkbox
    .GetText .ColLetterToNumber("N"), 34, chkTTDB
     
    Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 38))
    Set xmlCellsNode = xmlCellNode.parentNode
    Do
    If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
        i = 1
        ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
        .Row = iRow
               
        Do
            .GetText ColG13, .Row, GT2
            .GetText ColP13, .Row, GT3
            .GetText ColS13, .Row, GT4
            .GetText ColV13, .Row, GT5
            .GetText ColAA13, .Row, GT6
            .GetText ColAG13, .Row, GT8
            .GetText ColAJ13, .Row, GT9
            
            'Bat loi voi nhom 1, 2
            If iSection < 3 Then
                'Nhap Doanh so ban ma khong nhap Gia tinh thue thi bao loi
                'Checkbox duoc chon, neu chon ten hang ma khong nhap Gia tinh thue thi bao loi
                
                 .Col = ColAA13
                 .CellNote = ""
                 .BackColor = mNonErrorColor
                
                'Nhap Gia tinh thue ma khong nhap Doanh so ban thi bao loi
                'Checkbox duoc chon, neu chon ten hang ma khong nhap Doanh so ban thi bao loi
                
                .Col = ColV13
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                'Neu nhap Doanh so ban va Gia tinh thue ma Doanh so ban nho hon thi bao loi
                
                .Col = ColAA13_2
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                .Col = ColAJ13
                .CellNote = ""
                .BackColor = mNonErrorColor
                
             End If
             
             'Khong nhap ten hang ma nhap chi tiet thi bao loi
            
            .Col = ColG13
            .CellNote = ""
            .BackColor = mNonErrorColor
             
             .Row = i + iRow
             .Col = .ColLetterToNumber("C")
             i = i + 1
        Loop Until .Text = "aa"
        iSection = iSection + 1
        End If
        If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
        Else
            Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling
        End If
    Loop Until xmlCellsNode Is Nothing
    
    .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 17, "1"
    .SetText .ColLetterToNumber("B"), 16, "1"
    .SetText .ColLetterToNumber("B"), 18, "1"
    .SetText .ColLetterToNumber("B"), 19, "1"
    
    arrIdHangHoa.RemoveAll
    .Sheet = 1
    mCurrentSheet = 1

    XacDinhLoiAJ13 = False
    ColAJ13 = .ColLetterToNumber("AJ")
    
    .Sheet = .SheetCount
    .GetText .ColLetterToNumber("E"), 47, vErrMsgAJ13
    
    .Sheet = 1
    mCurrentSheet = 1
    iSection = 1
    XacDinhLoiAJ13 = False
    
    .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 47, "1"
    
    .Sheet = 1
    mCurrentSheet = 1
    fps_ButtonClicked .ColLetterToNumber("N"), 34, 1
    
    .Refresh
    
    'Dung cho xoa Ban giai trinh
    .Sheet = iCurrentSheet
    .ActiveSheet = iCurrentSheet
    .EventEnabled(EventAllEvents) = True
    
    End With
End Sub

'Description: IsValidatedCell function check whether a Cell with lCol, lRow
'             is validated cell which the validation of this cell and the Cell
'             with lAnchorRow, lAnchorCol are the same.
'Author:ThanhDX
'Date:09/02/2006
'Input:
'Return:
Private Function IsValidatedCell(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long, ByVal lRow As Long, ByVal lCol As Long) As Boolean
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    If lAnchorCol = lCol And lAnchorRow = lRow Then
        IsValidatedCell = True
        Exit Function
    End If
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    
    Do
        Set xmlCellsNode = xmlCellsNode.nextSibling
        If Not xmlCellsNode Is Nothing Then
            If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = fps.ColNumberToLetter(lCol) & "_" & lRow Then
                IsValidatedCell = True
                Exit Function
            End If
        End If
    Loop Until xmlCellsNode Is Nothing
    
End Function


Private Function CheckValidNgayHD(vdate As String) As Boolean
  
   Dim dd As String, mm As String, yyyy As String, dDate As Date, dDateNow As Date
   Dim sDateNow As String
   On Error GoTo e
    dd = Right("00" & Left(vdate, InStr(vdate, "/") - 1), 2)
    mm = Right("00" & Mid(vdate, 4, 2), 2)
    yyyy = Right("0000" & vdate, 4)
    sDateNow = "01/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
    dDateNow = Format(CDate(sDateNow), "dd/mm/yyyy")
    dDate = Format(CDate(vdate), "dd/mm/yyyy")
        If CInt(yyyy) > CInt(TAX_Utilities_v1.Year) Then
            CheckValidNgayHD = False
            Exit Function
        End If
        If DateAdd("m", 3, dDate) < dDateNow Then
        CheckValidNgayHD = False
            Exit Function
        End If
   CheckValidNgayHD = True
   Exit Function
   
e:
    CheckValidNgayHD = False
 End Function

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
    Dim blChon As Boolean
    Dim arrDanhsach() As String
    
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    
    blChon = False
    strPath = GetCatalogueFileName(blChon)
    
    'combo box khbs
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(6) = "1") Or blChon = False Then
'                    strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
'                  '  strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + Chr$(9)
'                    strComboName = strComboName + arrDanhsach(2) + Chr$(9)
                    strComboMCT = strComboMCT + "10" + Chr$(9)
                End If
            End If
        Next
    End If
    
    With fps
         .Sheet = .ActiveSheet
         mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("BC")
        .Row = 9
        .TypeComboBoxList = strComboHanghoa + strComboDichvu
        .Row = 13
        .TypeComboBoxList = strComboHanghoa + strComboDichvu
        .Col = .ColLetterToNumber("BD")
        .Row = 9
        .TypeComboBoxList = strComboHienHanghoa + strComboHienDichvu
        .Row = 13
        .TypeComboBoxList = strComboHienHanghoa + strComboHienDichvu
        .Col = .ColLetterToNumber("BE")
        .Row = 9
        .TypeComboBoxList = strComboMCT
        .Text = "10"
        .Lock = True
        .Row = 13
        .TypeComboBoxList = strComboMCT
        .Text = "10"
        .Lock = True
        
'Cho phep nhap am cot 4, 5, 6, 8 to khai bo sung
        .Sheet = 1
        .Row = 38
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AA")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AG")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        
        .Row = 42
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AA")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AG")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        
        .Row = 47
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AA")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AG")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        
        .Row = 51
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AA")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AG")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        
        .Row = 55
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AA")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AG")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        
     End With
    
End Sub


Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim xmlCell As MSXML.IXMLDOMNode
    Dim xmlCellUpdate As MSXML.IXMLDOMNode
    Dim xmlCellGoc As MSXML.IXMLDOMNode
    Dim strId As String
    With fps
       If loaiKHBS = "frmKHBS_TT" Then
            .Sheet = 1
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateLastKHBSCell fps, lCol, lRow, varTemp
            Next
       Else
            For Each xmlCell In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               If Left(GetAttribute(xmlCell, "CellID"), 2) = "BC" Then
                    strId = Trim(Mid(GetAttribute(xmlCell, "Value"), 200, 10))
                        
                    For Each xmlCellGoc In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
                        If Left(GetAttribute(xmlCellGoc, "CellID"), 1) = "E" Then
                            If GetAttribute(xmlCellGoc, "Value") = strId Then
                            'update secsion
                               Set xmlCellUpdate = seachCellNode(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell"), GetAttribute(xmlCellGoc, "CellID"))
                               updateSecsion xmlCellUpdate.parentNode
                                'updateSecsion (xmlCellGoc.parentNode)
                            Else
                            'insert secsion
                            
                            End If
                            
                        End If
                    Next
               
               End If
            Next
       End If
        
        
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Private Sub updateSecsion(xmlCellNode As MSXML.IXMLDOMNode)
Dim xmlCellNodeData As MSXML.IXMLDOMNode
Dim varTemp As Variant
Dim lCol As Long, lRow As Long
With fps
    For Each xmlCellNodeData In xmlCellNode.childNodes
     If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "S" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "V" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AA" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "AG" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AJ" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AN" Then
        ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
        .Sheet = 1
        .GetText lCol, lRow, varTemp
        UpdateLastKHBSCell fps, lCol, lRow, varTemp
     End If
    Next
End With
End Sub



Private Sub fillData()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim arrDanhsach() As String
    Dim xmlCell As MSXML.IXMLDOMNode
    With fps
        Dim i As Integer
        Dim nodecount As Integer
        Dim strFormular As String
        
        nodecount = TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(2).childNodes.length
        nodecount = nodecount + TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(4).childNodes.length
        nodecount = nodecount + TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(6).childNodes.length
        nodecount = nodecount + TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(7).childNodes.length
        nodecount = nodecount + TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(8).childNodes.length
        
        
        ReDim arrDanhSachKHBS(nodecount - 1)
        i = 0
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
            If GetAttribute(xmlCellNodeData.parentNode.parentNode, "Dynamic") = "1" Then
                If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "F" Then
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                     arrDanhSachKHBS(i) = varTemp & "~" & GetAttribute(xmlCellNodeData, "CellID")
                     i = i + 1
                End If
            End If
        Next
        
      .Sheet = .SheetCount - 1
    End With
End Sub


Private Function seachCellNode(xmlCellNodelist As MSXML.IXMLDOMNodeList, cellid As String) As MSXML.IXMLDOMNode
Dim nodeReturn As MSXML.IXMLDOMNode
    For Each nodeReturn In xmlCellNodelist
        If GetAttribute(nodeReturn, "CellID") = cellid Then
            Set seachCellNode = nodeReturn
            Exit For
            Exit Function
        End If
    Next
End Function


Sub fps_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Long, TipWidth As Long, TipText As String, ShowTip As Boolean)
    ' Set tip to display and set tip's content
      If Col = fps.ColLetterToNumber("G") Then
        ShowTip = True
        TipText = TipText
      End If
End Sub
Public Sub FinishImport()
'Dim i As Integer, iCol As Long, iRow As Long
'Dim maTM As Variant, strId As Variant
'Dim idx As Integer
'Dim varTemp As Variant
'With fps
'fps.EventEnabled(EventAllEvents) = False
''           If .ActiveSheet = 2 Then
''                    i = 1
''                    .Col = .ColLetterToNumber("B")
''                    .Row = 37
''                    Do
''                         .Text = str(i)
'''                         .Col = .ColLetterToNumber("C")
'''                         .Text = "0"
'''                         UpdateCell fps, .Col, .Row, .Text
''                         ' Lay ve ma tieu muc
''                         .GetText .ColLetterToNumber("I"), .Row, maTM
''                         idx = 0
''                         .Col = .ColLetterToNumber("K")
''                         For idx = 0 To .TypeComboBoxCount
''                            .typeComboboxCurSel = idx
''                            .GetText .ColLetterToNumber("K"), .Row, strId
''                            If Trim$(maTM) = Trim(Mid(strId, 200, 10)) Then
''                                .Col = .ColLetterToNumber("L")
''                                .typeComboboxCurSel = idx
''                                Exit For
''                            End If
''                         Next idx
''
''
''                         'chuyen ky hieu thanh chu hoa
''                         .Col = .ColLetterToNumber("C")
''                         .GetText .Col, .Row, varTemp
''                         .SetText .Col, .Row, UCase(varTemp)
''                         UpdateCell fps, .Col, .Row, .Text
''                         ' end
''
''                         .Col = .ColLetterToNumber("B")
''                         .Row = i + 37
''                         i = i + 1
''                    Loop Until .Text = "aa"
''                    .Row = iRow
''                    .Col = iCol
''                    .SetActiveCell iCol, iRow
''                    coutDyn = i - 1
''
''            End If
''            If .ActiveSheet = 3 Then
''                    i = 1
''                    .Col = .ColLetterToNumber("B")
''                    .Row = 10
''                    Do
''                         .Text = str(i)
''
''                         'chuyen ky hieu thanh chu hoa
''                         .Col = .ColLetterToNumber("C")
''                         .GetText .Col, .Row, varTemp
''                         .SetText .Col, .Row, UCase(varTemp)
''                         UpdateCell fps, .Col, .Row, .Text
''                         ' end
''
''                         .Col = .ColLetterToNumber("B")
''                         .Row = i + 10
''                         i = i + 1
''                    Loop Until .Text = "aa"
''
''
''                    .Row = iRow
''                    .Col = iCol
''                    .SetActiveCell iCol, iRow
''            End If
'        fps.EventEnabled(EventAllEvents) = True
'        End With
    'CellChange iCol, iRow
End Sub

Public Function getValueTK(strFileLoadData As String) As String

    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim result As String
    Dim iDom As Integer
    Dim MY As Date

    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("AI_15")
            result = "CT8_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01TTDB", "getValueTK", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
    Dim tempArrOld() As String
    Dim tempArrCurr() As String
    Dim tempArrOldValue() As String
    Dim tempArrCurrValue() As String
    Dim i, j As Integer
    Dim strMaCT As String
    Dim arrMapCT() As String ' Luu gia tri cua dong chi tieu
    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
    strMaCT = getMapCT
    With fps
                .Sheet = 1
                .GetText .ColLetterToNumber("AI"), 15, varTemp
                strCurrentValue = "CT8_" & Trim(varTemp)
                ' tinh chenh lech giua 2 lan dieu chinh
                tempArrOld = Split(strOldValue, "~")
                tempArrCurr = Split(strCurrentValue, "~")
                arrMapCT = Split(strMaCT, "~")
                strDieuChinhGiam = ""
                strDieuChinhTang = ""
                For i = 0 To UBound(tempArrOld)
                    tempArrOldValue = Split(tempArrOld(i), "_")
                    tempArrCurrValue = Split(tempArrCurr(i), "_")
                    ' lay cac gia tri cua tung cot trong sheet KHBS
                    For j = 0 To UBound(arrMapCT)
                        arrMapCTValue = Split(arrMapCT(j), "_")
                        ' phan tu so 0 la ma chi tieu
                        If arrMapCTValue(0) = Mid(tempArrOldValue(0), 3) Then
                            strMaCT = arrMapCT(j)
                            Exit For
                        End If
                    Next j

                If Left(strMaCT, 2) = "8_" Then
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                Else
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                End If
                Next i
                If strDieuChinhGiam <> "" Then
                    strDieuChinhGiam = Mid(strDieuChinhGiam, 2, Len(strDieuChinhGiam) - 1)
                End If
                If strDieuChinhTang <> "" Then
                    strDieuChinhTang = Mid(strDieuChinhTang, 2, Len(strDieuChinhTang) - 1)
                End If
                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang
      End With
End Function
Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim i As Integer
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") Then
                Exit For
            End If
        Next
    End If
    
    Set xmlNodeListItem = xmlNode.childNodes
    
    For Each xmlNodeItem In xmlNodeListItem
         tempValue = tempValue & "~" & GetAttribute(xmlNodeItem, "MCT") & "_" & GetAttribute(xmlNodeItem, "Caption")
    Next
    tempValue = Mid(tempValue, 2)
    getMapCT = tempValue
    Exit Function
ErrHandle:
    SaveErrorLog "cls_01TTDB", "getMapCT", Err.number, Err.Description
End Function



' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AN"), 15, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 32 o to khai lan truoc
Public Function getValueCTDC(strFileLoadData As String) As String

    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim result As String
    Dim iDom As Integer
    Dim MY As Date

    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("AN_15")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01TTDB", "getValueCTDC", Err.number, Err.Description
End Function
'******************************
'Set trang thai Flag KHBS active
Public Sub SetActiveKHBS() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetKHBS - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetKHBS - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(sheetKHBS - 1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(sheetKHBS - 1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(3).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                        FlagThaydoi = "1"
                    End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub


'update sheet KHBS
' 27-01-2011
Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim blCheckS2 As Boolean
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        
        
'        .Sheet = 2
'        mCurrentSheet = 2
'        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheckS2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'         End If
'        .Sheet = 3
'        mCurrentSheet = 3
'        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheckS2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'        End If
        
        
       ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 12
    denRowDL = 17
    'TuRow_DenRow_sheet
    'strTuRowDenRowPL = "8_10_2~10_12_3~20_30_4"
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

'
Public Sub getNhomHH(ByVal maNhom As String, ByRef tenNhom As String)
    Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetCatalogueFileName(False)
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maNhom = arrDanhsach(1) Then
                            tenNhom = arrDanhsach(2) + " (" + arrDanhsach(4) + " %)"
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


Public Sub getRowDanhMuc(ByVal maNhom As String, ByRef tenNhom As String)
    Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetCatalogueFileName(False)
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maNhom = arrDanhsach(1) Then
                            tenNhom = CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 15) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 5) + CPab(arrDanhsach(5), 14) + CPab(arrDanhsach(6), 1)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub

' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim i As Integer
    Dim j As Integer
    Dim maNhom    As Variant
    Dim tenNhom  As String
    
    Dim tong1, tong2 As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        ' group 1
        i = 1
        .Row = 38
        .Col = .ColLetterToNumber("C")
        Do
            .GetText .ColLetterToNumber("E"), .Row, maNhom
            getRowDanhMuc maNhom, tenNhom
            
            .Col = .ColLetterToNumber("F")
            .Text = tenNhom
            UpdateCell fps, .Col, .Row, .Text
            
            If tenNhom <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = tenNhom
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            .Col = .ColLetterToNumber("D")
            .Text = Trim(Left(tenNhom, 200))
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("R")
            .value = Trim(Mid(tenNhom, 245, 1))
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("C")
             .Text = str(i)
             UpdateCell fps, .Col, .Row, .value
             .Row = i + 38
             i = i + 1
        Loop Until .Text = "aa"
        ' group 2
        .Row = .Row + 3
        i = 1
         Do
         
            .GetText .ColLetterToNumber("E"), .Row, maNhom
            getRowDanhMuc maNhom, tenNhom
            
            .Col = .ColLetterToNumber("F")
            .Text = tenNhom
            UpdateCell fps, .Col, .Row, .Text
            
            If tenNhom <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = tenNhom
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            .Col = .ColLetterToNumber("D")
            .Text = Trim(Left(tenNhom, 200))
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("R")
            .value = Trim(Mid(tenNhom, 245, 1))
            UpdateCell fps, .Col, .Row, .Text

            
            .Col = .ColLetterToNumber("C")
             .Text = str(i)
             UpdateCell fps, .Col, .Row, .value
             .Row = .Row + 1
             i = i + 1
        Loop Until .Text = "aa"
        ' group 3
        .Row = .Row + 4
        i = 1
         Do
         
            .GetText .ColLetterToNumber("E"), .Row, maNhom
            getRowDanhMuc maNhom, tenNhom
            
            .Col = .ColLetterToNumber("F")
            .Text = tenNhom
            UpdateCell fps, .Col, .Row, .Text
            
            If tenNhom <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = tenNhom
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            .Col = .ColLetterToNumber("D")
            .Text = Trim(Left(tenNhom, 200))
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("R")
            .value = Trim(Mid(tenNhom, 245, 1))
            UpdateCell fps, .Col, .Row, .Text

            
            .Col = .ColLetterToNumber("C")
             .Text = str(i)
             UpdateCell fps, .Col, .Row, .value
             .Row = .Row + 1
             i = i + 1
        Loop Until .Text = "aa"
        ' group 4
        .Row = .Row + 3
        i = 1
         Do
         
            .GetText .ColLetterToNumber("E"), .Row, maNhom
            getRowDanhMuc maNhom, tenNhom
            
            .Col = .ColLetterToNumber("F")
            .Text = tenNhom
            UpdateCell fps, .Col, .Row, .Text
            
            If tenNhom <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = tenNhom
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            .Col = .ColLetterToNumber("D")
            .Text = Trim(Left(tenNhom, 200))
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("R")
            .value = Trim(Mid(tenNhom, 245, 1))
            UpdateCell fps, .Col, .Row, .Text

            
            .Col = .ColLetterToNumber("C")
             .Text = str(i)
             UpdateCell fps, .Col, .Row, .value
             .Row = .Row + 1
             i = i + 1
        Loop Until .Text = "aa"
        
        ' group 5
        .Row = .Row + 3
        i = 1
         Do
         
            .GetText .ColLetterToNumber("E"), .Row, maNhom
            getRowDanhMuc maNhom, tenNhom
            
            .Col = .ColLetterToNumber("F")
            .Text = tenNhom
            UpdateCell fps, .Col, .Row, .Text
            
            If tenNhom <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = tenNhom
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            .Col = .ColLetterToNumber("D")
            .Text = Trim(Left(tenNhom, 200))
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("R")
            .value = Trim(Mid(tenNhom, 245, 1))
            UpdateCell fps, .Col, .Row, .Text

            
            .Col = .ColLetterToNumber("C")
             .Text = str(i)
             UpdateCell fps, .Col, .Row, .value
             .Row = .Row + 1
             i = i + 1
        Loop Until .Text = "aa"
        
        .Row = .Row + 1
        .GetText .ColLetterToNumber("AG"), .Row, tong1
        .GetText .ColLetterToNumber("AN"), .Row, tong2
        
        
        .Col = .ColLetterToNumber("AI")
        .Row = 15
        .Text = tong1
        UpdateCell fps, .Col, .Row, .value
        
        .Col = .ColLetterToNumber("AN")
        .Row = 15
        .Text = tong2
        UpdateCell fps, .Col, .Row, .value
        
        .Col = .ColLetterToNumber("AI")
        .Row = 6
        .Text = tong1
        UpdateCell fps, .Col, .Row, .value
        
        .Col = .ColLetterToNumber("AN")
        .Row = 6
        .Text = tong2
        UpdateCell fps, .Col, .Row, .value
        
        
        .Sheet = 2
        If .SheetVisible = True Then
                tenNhom = ""
                mCurrentSheet = .Sheet
                i = 1
                .Row = 37
                Do
                    tenNhom = ""
                    .GetText .ColLetterToNumber("I"), .Row, maNhom
                     getNhomHH maNhom, tenNhom
                    .Col = .ColLetterToNumber("L")
                    .Text = tenNhom
                    UpdateCell fps, .Col, .Row, .Text
                    
                    .GetText .ColLetterToNumber("I"), .Row, maNhom
                    getRowDanhMuc maNhom, tenNhom
                    
                    .Col = .ColLetterToNumber("K")
                    .Text = tenNhom
                    UpdateCell fps, .Col, .Row, .Text
                    
                     .Col = .ColLetterToNumber("B")
                     .Text = str(i)
                     .Col = .ColLetterToNumber("B")
                     UpdateCell fps, .Col, .Row, .value
                     .Row = i + 37
                     i = i + 1
                Loop Until .Text = "aa"
            
        End If
        
        .Sheet = 3
        If .SheetVisible = True Then
                mCurrentSheet = .Sheet
                i = 1
                j = 1
                .Col = .ColLetterToNumber("B")
                .Row = 10
                Do
                     .Text = str(i)
                     .Col = .ColLetterToNumber("B")
                     UpdateCell fps, .Col, .Row, .value
                     .Row = i + 10
                     i = i + 1
                Loop Until .Text = "aa"
                
                .Row = 17 + i + j
                Do
                     .Text = str(j - 1)
                     .Col = .ColLetterToNumber("B")
                     UpdateCell fps, .Col, .Row, .value
                     .Row = j + i + 17
                     j = j + 1
                Loop Until .Text = "aa"
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
