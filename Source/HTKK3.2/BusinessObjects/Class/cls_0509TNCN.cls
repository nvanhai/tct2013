VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_0509TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "To khai quyet toan thu nhap ca nhan mau 04-TNCN" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String

Private arrData5A() As String
Private arrData5B() As String
Dim arrData0509() As String
Private totalRow5A_09 As Long
Private rowData As Long

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag ...
'No parameter
Public Sub Prepare1()
    With fps
        
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim i, j As Long
    Dim totalRow As Long
    ' Lay du lieu cua bang ke 5A
    prepareData5A
    ' Lay du lieu cua bang ke 5B
    prepareData5B
    ' Tong hop du lieu tu cac bang ke va phu luc len mang arrData0509
    prepareData0509
     
    With fps
        .Refresh
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .Row = 22
        
        
        totalRow = UBound(arrData0509) - 1
        For i = 0 To totalRow
            .Row = i + 22
            ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
            .Col = .ColLetterToNumber("C")
            .Text = "0"
            For j = 0 To 20
                ' Tu cot Ho va ten cho den cot thue da khau tru (D -> L). Dung cac thong tin theo bang ke 05
                .Col = 4 + j
                'dhdang sua loi tong hop len cell bi sai do dau "."
                'ngày 03-07-10
                If .Col = .ColLetterToNumber("R") Then
                    .Col = .ColLetterToNumber("AG")
                End If
                If .CellType = CellTypeNumber Then
                    .Text = Round(Val(arrData0509(i, j)), 0)
                Else
                    .Text = arrData0509(i, j)
                End If
                
'                If Trim(arrData0509(i, 1)) = "" Or Trim(arrData0509(i, 1)) = vbNullString Then
'                    .Col = .ColLetterToNumber("C")
'                    .Lock = True
'                Else
'                    .Col = .ColLetterToNumber("C")
'                    .Lock = False
'                End If
            Next
            ' Trong truong hop la dong cuoi cung trong mang thi ko Insert them row nao.
            If i <> totalRow Then
                .ReDraw = False
                .MaxRows = .MaxRows + 1
                .InsertRows .Row + 1, 1
                .CopyRowRange 22, 22, .Row + 1  ' Copy nguyen ban cac format cua dong bat dau la dong 22
            End If
            ' Set cong thuc cho cac cot sau
            ' Thue phai nop
            .Col = .ColLetterToNumber("V")
            '.Formula = "AC" & .Row
            .Formula = "AH" & .Row
            ' Thue mien giam
            .Col = .ColLetterToNumber("W")
            .Formula = "AD" & .Row
            ' Thue con phai nop
            .Col = .ColLetterToNumber("Y")
            .Formula = "AE" & .Row
            ' Con phai quyet toan thue
            .Col = .ColLetterToNumber("Z")
            .Formula = "AF" & .Row
        Next
        'dhdang
        'button Lua Chon In
        
        .ActiveSheet = .Sheet
        .Row = 18
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\Check.bmp"))
        '.TypeButtonText = "Chän"
        
        'dhdang
        'sap xep
        
        
        .ActiveSheet = .Sheet
        ' Sap xep theo cot Ho va ten
        .Row = 18
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0154"), "Msg")
        
        ' Sap xep theo cot quyet toan thue
        .Row = 20
        .Col = .ColLetterToNumber("AC")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = ""

'        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub prepareData5A()
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCells As MSXML.IXMLDOMNodeList
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim i, j As Long
    
    Dim fso As New FileSystemObject
    
    
    strDataFileName = GetAbsolutePath(TAX_Utilities_v1.DataFolder & "05_1_TNCN_" & TAX_Utilities_v1.Year & "_" & Replace(TAX_Utilities_v1.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v1.LastDay, "/", "") & ".xml")
    i = 0
    If fso.FileExists(strDataFileName) = False Then
        ReDim arrData5A(1, 18) As String
        arrData5A(0, 0) = ""
        arrData5A(0, 1) = ""
        arrData5A(0, 2) = ""
        arrData5A(0, 3) = "0"
        arrData5A(0, 4) = "0"
        arrData5A(0, 5) = "0"
        arrData5A(0, 6) = "0"
        arrData5A(0, 7) = "0"
        arrData5A(0, 8) = "0"
        arrData5A(0, 9) = "0"
        arrData5A(0, 10) = ""
        arrData5A(0, 11) = "0"
        arrData5A(0, 12) = "0"
        arrData5A(0, 13) = "0"
        arrData5A(0, 14) = "0"
        arrData5A(0, 15) = "0"
        arrData5A(0, 16) = "0"
        arrData5A(0, 17) = ""
        Exit Sub
    End If
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCells = xmlDomData.getElementsByTagName("Cells")
        
        ReDim arrData5A(xmlNodeListCells.length, 18) As String
        
        For Each xmlNodeCells In xmlNodeListCells
            Set xmlNodeListCell = xmlNodeListCells.Item(i).childNodes
            j = 0
            For Each xmlNodeCell In xmlNodeListCell
'                If j = 6 Then
'                    arrData5A(i, j) = IIf(Val(GetAttribute(xmlNodeCell, "Value")) - 48000000 < 0, 0, Val(GetAttribute(xmlNodeCell, "Value")) - 48000000)
'                Else
                    arrData5A(i, j) = GetAttribute(xmlNodeCell, "Value")
'                End If
                j = j + 1
            Next
            i = i + 1
        Next
        Set xmlDomData = Nothing
        Set xmlNodeListCells = Nothing
        Set xmlNodeListCell = Nothing
        Set xmlNodeCells = Nothing
        Set xmlNodeCell = Nothing
        
    End If
End Sub

Private Sub prepareData5B()
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCells As MSXML.IXMLDOMNodeList
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim i, j As Long
    
    Dim fso As New FileSystemObject
    
    
    strDataFileName = GetAbsolutePath(TAX_Utilities_v1.DataFolder & "05_2_TNCN_" & TAX_Utilities_v1.Year & ".xml")
    i = 0
    If fso.FileExists(strDataFileName) = False Then
        ReDim arrData5B(1, 10) As String
        arrData5B(0, 0) = ""
        arrData5B(0, 1) = ""
        arrData5B(0, 2) = ""
        arrData5B(0, 3) = "0"
        arrData5B(0, 4) = "0"
        arrData5B(0, 5) = "0"
        arrData5B(0, 6) = "0"
        arrData5B(0, 7) = "0"
        arrData5B(0, 8) = "0"
        arrData5B(0, 9) = ""
        Exit Sub
    End If
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCells = xmlDomData.getElementsByTagName("Cells")
        
        ReDim arrData5B(xmlNodeListCells.length, 8) As String
        
        For Each xmlNodeCells In xmlNodeListCells
            Set xmlNodeListCell = xmlNodeListCells.Item(i).childNodes
            j = 0
            For Each xmlNodeCell In xmlNodeListCell
                arrData5B(i, j) = GetAttribute(xmlNodeCell, "Value")
                If j = 9 Then Exit For
                j = j + 1
            Next
            i = i + 1
        Next
        Set xmlDomData = Nothing
        Set xmlNodeListCells = Nothing
        Set xmlNodeListCell = Nothing
        Set xmlNodeCells = Nothing
        Set xmlNodeCell = Nothing
        
    End If
End Sub

Private Sub prepareData0509()
    Dim totalRow5A, totalRow5B As Long
    Dim row5A, row5B As Long
        
    ' arrData5A Bang ke 5A co 10 phan tu
    totalRow5A = UBound(arrData5A)
    ' arrData5B Bang ke 5B co 6 phan tu
    totalRow5B = UBound(arrData5B)
    For row5A = 0 To totalRow5A - 1
        For row5B = 0 To totalRow5B - 1
            ' Neu ton tai ca MST tren 5A va 5B
            If (Trim(arrData5A(row5A, 1)) = Trim(arrData5B(row5B, 1))) Then
                If Trim(arrData5A(row5A, 1)) <> "" Then
                    ' Thu nhap chiu thue = Thu nhap chiu thue 5A + Thu nhap chiu thue 5B
                    arrData5A(row5A, 4) = Val(arrData5A(row5A, 4)) + Val(arrData5B(row5B, 4))
                    ' Thu nhap lam can cu tinh giam thue = Thu nhap lam can cu tinh giam thue 5A + Thu nhap lam can cu tinh giam thue 5B
                    ' lam viec tai KKT
                    arrData5A(row5A, 5) = Val(arrData5A(row5A, 5)) + Val(arrData5B(row5B, 5))
                    ' lam viec theo HD
                    arrData5A(row5A, 6) = Val(arrData5A(row5A, 6)) + Val(arrData5B(row5B, 6))
                    ' Thue TNCN da khau tru = Thue TNCN da khau tru 5A + Thue TNCN da khau tru 5B
                    arrData5A(row5A, 12) = Val(arrData5A(row5A, 12)) + Val(arrData5B(row5B, 7))
                    ' giam theo KKT
                    arrData5A(row5A, 13) = Val(arrData5A(row5A, 13)) + Val(arrData5B(row5B, 8))
                    Exit For
                End If
            End If
        Next
    Next
    ' arrData0509 la tong hop cac bang ke va phu luc tren co 19 phan tu
    totalRow5A_09 = totalRow5A
    ' Dinh nghia lai mang 0509 lay theo bang ke 5A
    ReDim arrData0509(totalRow5A, 29) As String
    For row5A = 0 To totalRow5A - 1
        ' Cot 0 la cot Ho va ten
        arrData0509(row5A, 0) = arrData5A(row5A, 0)
        ' Cot 1 la cot Ma so thue
        arrData0509(row5A, 1) = arrData5A(row5A, 1)
        ' Cot 2 la cot So CMT/Ho chieu
        arrData0509(row5A, 2) = arrData5A(row5A, 2)
        ' Cot 3 uy quyen quyet toan thay
        arrData0509(row5A, 3) = arrData5A(row5A, 3)
        ' Cot 4 la Thu nhap chiu thue
        arrData0509(row5A, 4) = arrData5A(row5A, 4)
        ' Cot 5 la TNCN lam can cu giam tru theo KKT
        arrData0509(row5A, 5) = arrData5A(row5A, 5)
        ' Cot 6 la cot tong so tien giam tru theo HD
        arrData0509(row5A, 6) = arrData5A(row5A, 6)
        ' Cot 7 la cot tong giam tru gia canh
        arrData0509(row5A, 7) = arrData5A(row5A, 7)
        ' Cot 8 la cot tu thien , khuyen hoc
        arrData0509(row5A, 8) = arrData5A(row5A, 8)
        ' Cot 9 la bao hiem duoc tru
        arrData0509(row5A, 9) = arrData5A(row5A, 9)
        ' Cot 10 quy huu tri
        arrData0509(row5A, 10) = arrData5A(row5A, 10)
        ' cot 11 thu nhap tinh thue
        arrData0509(row5A, 11) = arrData5A(row5A, 11)
        ' cot 12 so thue duoc khau tru
        arrData0509(row5A, 12) = arrData5A(row5A, 12)
        ' cot 13 so thue duoc giam theo KKT
        arrData0509(row5A, 13) = arrData5A(row5A, 13)
        
        
        ' Cot 14 la cot Thu nhap chiu thue
        arrData0509(row5A, 14) = arrData5A(row5A, 4)
        
        
        
        ' Cot 21 la cot check co quyet toan tai CQCT ko?
        arrData0509(row5A, 22) = arrData5A(row5A, 3)
        ' Cot 22 la cot tong so thue phai nop tren bang ke 5A
        arrData0509(row5A, 23) = arrData5A(row5A, 15)
        ' Cot 23 la cot so thue da nop thua
        arrData0509(row5A, 24) = arrData5A(row5A, 16)
        ' Cot 24 la cot so thue con phai khau tru them
        arrData0509(row5A, 25) = arrData5A(row5A, 17)
                
        ' Cot 11 la Giam tru cho ban than
        ' Neu nam 2009 thi giam tru 6 thang la 24.000.000 (Tuong duong voi 4.000.000/Thang) va tu nam 2010 la 12 thang tuong duong voi 48.000.000
        If TAX_Utilities_v1.Year = 2009 Then
            arrData0509(row5A, 17) = "24000000"
        ElseIf TAX_Utilities_v1.Year > 2009 Then
            arrData0509(row5A, 17) = "48000000"
        Else
            arrData0509(row5A, 17) = "0"
        End If
        
        ' Cot 18 la giam tru gia canh
        arrData0509(row5A, 18) = Val(arrData5A(row5A, 7))
        ' Cot 19 la tu thien
        arrData0509(row5A, 19) = arrData5A(row5A, 8)
        ' Cot 20 la bao hiem bat buoc
        arrData0509(row5A, 20) = arrData5A(row5A, 9)
        
        ' Cot 21 la huu tri
        arrData0509(row5A, 21) = arrData5A(row5A, 10)
        
        ' Cot 16 la Thu nhap tinh thue la tong thu nhap chiu thue - giam tru ban than - giam tru nguoi phu thuoc - tu thien - bao hiem bat buoc
        
        arrData0509(row5A, 22) = arrData5A(row5A, 11)
'        If Val(arrData0509(row5A, 15)) < 0 Then arrData0509(row5A, 15) = "0"
        
        ' Cot 17 la Thue phai nop la thu nhap tinh thue tinh theo bieu thue luy tien (Doi voi nam 2009 thi tinh 6 thang tu nam 2010 thi tinh du 12 thang)
        ' arrData0509(row5A, 16) = tinh theo bieu thue luy tien cua arrData0509(row5A, 14)
        arrData0509(row5A, 23) = "0"
'        Dim thuNhapTinhThueThang As Double, thuePhaiNopThang As Double
'        Dim soThangTinhThue As Integer
'        ' Neu la nam 2009 thi thu nhap tinh thue tren 6 thang
'        If TAX_Utilities_v1.Year = 2009 Then
'            soThangTinhThue = 6
'            thuNhapTinhThueThang = Val(arrData0509(row5A, 15)) / soThangTinhThue
'        ElseIf TAX_Utilities_v1.Year > 2009 Then ' thu nhap tinh thue tinh tren 12 thang
'            soThangTinhThue = 12
'            thuNhapTinhThueThang = Val(arrData0509(row5A, 15)) / soThangTinhThue
'        Else ' Khong tinh
'            soThangTinhThue = 0
'            thuNhapTinhThueThang = 0
'        End If
'
'        If thuNhapTinhThueThang < 5000000 Then
'            thuePhaiNopThang = thuNhapTinhThueThang * 0.05
'        ElseIf thuNhapTinhThueThang >= 5000000 And thuNhapTinhThueThang < 10000000 Then
'            thuePhaiNopThang = (thuNhapTinhThueThang * 0.1) - 250000
'        ElseIf thuNhapTinhThueThang >= 10000000 And thuNhapTinhThueThang < 18000000 Then
'            thuePhaiNopThang = (thuNhapTinhThueThang * 0.15) - 750000
'        ElseIf thuNhapTinhThueThang >= 18000000 And thuNhapTinhThueThang < 32000000 Then
'            thuePhaiNopThang = (thuNhapTinhThueThang * 0.2) - 1650000
'        ElseIf thuNhapTinhThueThang >= 32000000 And thuNhapTinhThueThang < 52000000 Then
'            thuePhaiNopThang = (thuNhapTinhThueThang * 0.25) - 3250000
'        ElseIf thuNhapTinhThueThang >= 52000000 And thuNhapTinhThueThang < 80000000 Then
'            thuePhaiNopThang = (thuNhapTinhThueThang * 0.3) - 5850000
'        Else
'            thuePhaiNopThang = (thuNhapTinhThueThang * 0.35) - 9850000
'        End If
'        arrData0509(row5A, 16) = Round(thuePhaiNopThang * soThangTinhThue, 0)
        
        ' Cot 17 la Thue duoc mien giam
        ' Thue duoc mien giam = ThuePhaiNop * (ThuNhapGiamThue / ThuNhapChiuThue) * 0.5, 0)
'        If Val(arrData5A(row5A, 3)) = 0 Then
'            arrData0509(row5A, 17) = "0"
'        Else
'            arrData0509(row5A, 17) = Round(Val(arrData0509(row5A, 16)) * (Val(arrData5A(row5A, 8)) / Val(arrData5A(row5A, 3))) * 0.5, 0)
'        End If
'
        arrData0509(row5A, 24) = "0"
        
        ' Cot 18 la so thue da khau tru la so thue da khau tru tren bang ke 5A
        '= So thue TNCN da khau tru[Tren 5A+ Tren 5B] - So thue nop thua + So thue con khau tru them
        arrData0509(row5A, 25) = Val(arrData5A(row5A, 10))
'        arrData0509(row5A, 18) = Val(arrData5A(row5A, 9)) - Val(arrData0509(row5A, 23)) + Val(arrData0509(row5A, 24))
        
        ' Cot 19 la so thue con phai nop
        ' arrData0509(row5A, 19) = Thue phai nop - Thue mien giam - Thue da khau tru
        arrData0509(row5A, 26) = "0"
'        arrData0509(row5A, 19) = Round(Val(arrData0509(row5A, 16)) - Val(arrData0509(row5A, 17)) - Val(arrData0509(row5A, 18)), 0)
        
'        If (Val(arrData0509(row5A, 19)) >= -1 And Val(arrData0509(row5A, 19)) <= 1) Then arrData0509(row5A, 19) = 0
        
        ' Cot 20 la co phai quyet toan thue hay khong?
        arrData0509(row5A, 27) = ""
'        If Val(arrData0509(row5A, 19)) = 0 Then
'            arrData0509(row5A, 20) = TAX_Utilities_v1.Convert("Kh«ng", TCVN, UNICODE)
'        Else
'            arrData0509(row5A, 20) = TAX_Utilities_v1.Convert("Cã", TCVN, UNICODE)
'        End If
                
    Next
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
End Function

'dhdang
'26/01/10
Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim Y As Long
    Static so As Long
    If so = 0 Then
        so = 1
    End If
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 1 Then
               .Sheet = 1
                
                If Col = .ColLetterToNumber("D") And Row = 20 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("X"), 21 + totalRow5A_09, SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("AC") And Row = 20 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("AC"), 22, .ColLetterToNumber("AC"), (21 + totalRow5A_09), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
            End If
            .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    'UpdateSheets
   
End Sub


'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        SetData
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
        
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    With fps
        
        
    End With
End Sub

Public Sub SetActiveSheet()

End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean

End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()


End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Private Sub UpdateSheets()
    
End Sub

Public Sub FinishImport()

End Sub

Private Sub InsertNode(ByVal pCol As Long, ByVal pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    mCurrentSheet = 1
    
    ' Get cellspan for merge cell on interface templates
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fps, xmlNodeCells, lRow2s, lLRowBound, lURowBound)
    
    'If Not xmlNodeCells.nextSibling Is Nothing Then GoTo EXIT_SUB
    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    If Val(GetAttribute(xmlNodeCells.parentNode, "MaxRows")) = xmlNodeCells.parentNode.childNodes.length Then GoTo EXIT_SUB
  
    
    ' increase value of row in xmlDocument
    IncreaseRowInDOM fps, TAX_Utilities_v1.Data(0), lURowBound + 1, lRows, lRow2s
    'IncreaseRowInDOM lURowBound + 1, lRows, lRow2s

    Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fps, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fps, lCol, lRow + lRows)
                
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
        
        ' Reset empty value for new node
        fps.Col = lCol
        fps.Row = lRow
        Select Case fps.CellType
            Case CellTypeNumber
                SetAttribute xmlNodeNewCell, "Value", "0"
            Case Else
                SetAttribute xmlNodeNewCell, "Value", vbNullString
        End Select
        
        ' Set new ID for node (CellID2)
        ParserCellID fps, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fps, lCol, lRow + lRow2s)
    Next
    rowData = lURowBound + 22 ' Cac dong tiep theo phai bat dau tu Cell dong cuoi cung + 22 (Cell dong dau tien)
    ' Insert new node to DOM object
    If Not xmlNodeCells.nextSibling Is Nothing Then
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, xmlNodeCells.nextSibling
    Else
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
    End If
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_0509TNCN", "InsertNode", Err.number, Err.Description
End Sub

Private Sub IncreaseRowInDOM(fps As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To 0 Step -1
        ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fps, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fps, lCol, lRow + lRow2s)
        End If
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_0509TNCN", "IncreaseRowInDOM", Err.number, Err.Description
End Sub

Private Sub DeleteNode(ByVal intSheet As Integer, ByVal pCol As Long, ByVal pRow As Long, Optional ByVal blnForce As Boolean = True)
    On Error GoTo ErrorHandle
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    Dim xmlNodeCells As MSXML.IXMLDOMNode, xmlNodeTemp As MSXML.IXMLDOMNode
    Dim xmlNodeTemp2 As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    
    mCurrentSheet = 1
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fps, xmlNodeCells, lRow2s, lLRowBound, lURowBound)

    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    
'*********************************************************
' added
'Date: 01/03/06
    'Check whether user want to delete
'    If lRows > 1 And blnForce And Not IsEmptyValue(xmlNodeCells) Then
'        If DisplayMessage("0075", msYesNo, miQuestion, , mrYes) = mrNo Then
'            Exit Sub
'        End If
'    End If
'*********************************************************
        
    If xmlNodeCells.parentNode.childNodes.length <= 1 Then
        ClearRows xmlNodeCells
        TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
        GoTo EXIT_SUB
    Else
        If GetAttribute(xmlNodeCells.firstChild, "FirstCell") = "0" Then
            'Set FirstCell attr to "0" for next Cells node
            Set xmlNodeCells = xmlNodeCells.nextSibling
            SetAttribute xmlNodeCells.firstChild, "FirstCell", "0"
            Set xmlNodeCells = xmlNodeCells.previousSibling
        End If
    End If
    
    'If xmlNodeCells.xml = xmlNodeCells.parentNode.childNodes(0).xml Then GoTo EXIT_SUB
    
    'Jump active cell to prevous section
    'fps.SetActiveCell fps.ActiveCol, fps.ActiveRow - lRows
    
'    ' Delete curent row on Form
'    DeleteRow intSheet, lLRowBound, lRows
'    'fps.SetFocus
    
    xmlNodeCells.parentNode.removeChild xmlNodeCells
    
    ' Decrease value of row in xmlDocument
    DecreaseRowInDOM intSheet, lLRowBound + 1, lRows, lRow2s
    
    fps.Col = fps.ActiveCol
    fps.Row = fps.ActiveRow
    If fps.Lock = True Then
        Do
            fps.Row = fps.Row - 1
        Loop Until (Not fps.Lock) Or (fps.Row = 1)
        fps.SetActiveCell fps.Col, fps.Row
    End If
'   ************************************
'    added
'   Date: 12/04/06
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
'   ************************************
EXIT_SUB:
    Set xmlNodeCells = Nothing
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_0509TNCN", "DeleteNode", Err.number, Err.Description
End Sub

Private Sub DecreaseRowInDOM(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
    
    If TAX_Utilities_v1.Data(0) Is Nothing Then Exit Sub
    Set xmlNodeListCell = TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
    
    For i = 0 To xmlNodeListCell.length - 1
        ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Decrease value of row attribute - 1 "CellID"
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fps, lCol, lRow - lRows)
            
            ' Decrease value of row attribute - 1 "CellID2"
            ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fps, lCol, lRow - lRow2s)
        End If
    Next
    
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_0509TNCN", "DecreaseRowInDOM", Err.number, Err.Description
End Sub

Private Function GetDynRowCount(pGrid As fpSpread, xmlNodeCells As MSXML.IXMLDOMNode, Optional ByRef lReportRows As Long, Optional ByRef lMinRow As Long, Optional lMaxRow As Long)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lRow As Long, lCol As Long
    Dim lRow2 As Long, lCol2 As Long
    Dim lMaxRow2 As Long, lMinRow2 As Long
    
    lMinRow = 100000
    lMaxRow = 0
    lMinRow2 = 100000
    lMaxRow2 = 0
    
    If Not xmlNodeCells Is Nothing Then
        For Each xmlNodeCell In xmlNodeCells.childNodes
            'Get CellID
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
            
            'Get CellID2
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol2, lRow2
            
            'Get max row
            If lRow > lMaxRow Then
                lMaxRow = lRow
            End If
            
            'Get min row
            If lRow < lMinRow Then
                lMinRow = lRow
            End If
            
            'Get max row
            If lRow2 > lMaxRow2 Then
                lMaxRow2 = lRow2
            End If
            
            'Get min row
            If lRow2 < lMinRow2 Then
                lMinRow2 = lRow2
            End If
        Next
        
        GetDynRowCount = lMaxRow - lMinRow + 1
        lReportRows = lMaxRow2 - lMinRow2 + 1
    End If
End Function

Private Function IsEmptyValue(xmlCellsNode As MSXML.IXMLDOMNode) As Boolean
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim blnIsEmptyValue As Boolean
    
    blnIsEmptyValue = True
    
    For Each xmlCellNode In xmlCellsNode.childNodes
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
        fps.Col = lCol
        fps.Row = lRow
        Select Case fps.CellType
            Case CellTypeNumber, CellTypePercent
                If Not IsNullNumber(GetAttribute(xmlCellNode, "Value")) Then
                    blnIsEmptyValue = False
                    Exit For
                End If
            Case CellTypePic
                If Not IsNullPic(GetAttribute(xmlCellNode, "Value")) Then
                    blnIsEmptyValue = False
                    Exit For
                End If
'            Case CellTypeDate
'                If GetAttribute(xmlCellNode, "Value") <> "" Then
'                    blnIsEmptyValue = False
'                    Exit For
'                End If
            Case Else
                If GetAttribute(xmlCellNode, "Value") <> "" Then
                    blnIsEmptyValue = False
                    Exit For
                End If
        End Select
    Next
    
    IsEmptyValue = blnIsEmptyValue
End Function

Private Sub ClearRows(xmlCellsNode As MSXML.IXMLDOMNode) '(ByVal lRow As Long, ByVal lRows As Long)
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode

    With fps
        For Each xmlCellNode In xmlCellsNode.childNodes
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
            .Sheet = mCurrentSheet
            .Col = lCol
            .Row = lRow
            If .Lock = False Or (.Lock = True And .Formula = vbNullString) Then
                Select Case .CellType
                    Case CellTypeNumber
                        '.Text = "0"
                        SetAttribute xmlCellNode, "Value", "0"
                    Case Else
                        '.Text = ""
                        SetAttribute xmlCellNode, "Value", ""
                End Select
           End If
                
            
            If .CellNote <> vbNullString Then
                .CellNote = vbNullString
                .BackColor = vbWhite
            End If
            '*******************************
        Next
    End With
End Sub

Private Function IsNullNumber(ByVal strValue As String) As Boolean
    strValue = Replace$(strValue, "0", "")
    strValue = Replace$(strValue, ".", "")
    If Trim(strValue) = "" Then IsNullNumber = True
End Function

Private Function IsNullPic(ByVal strValue As String) As Boolean
    strValue = Replace$(strValue, "/", "")
    strValue = Replace$(strValue, "\", "")
    strValue = Replace$(strValue, ".", "")
    If Trim(strValue) = "" Then IsNullPic = True
End Function

Public Sub prepareDataPrinter()
    Dim i As Integer
    Dim iFirst As Byte
    Dim varCheck As Variant
    Dim rowGrid As Long
    
    With fps
        fps.EventEnabled(EventAllEvents) = False
        .Sheet = 1
        ' Xoa du lieu trong object 05_TNCN_09.xml, xoa tat va chi giu lai mot cells duy nhat
        ' Sau do clear va set lai cac gia tri theo cell type
        Do
            DeleteNode .Sheet, .ColLetterToNumber("D"), 22
            .Col = .ColLetterToNumber("B")
            .Row = i + 22
            i = i + 1
        Loop Until .Text = "aa"
        ' Dat iFirst = 1, danh dau gia tri duoc check dau tien se chi can update du lieu la xong
        ' Ko sinh ra mot note trong object 05_TNCN_09
        iFirst = 1
        i = 1
        .Row = 22
        ' rowData = 22 la dong du lieu dau tien duoc cap nhat vao trong object 05_TNCN_09
        ' Va du lieu bat dau tu dong 22 (Tham chieu den thiet ke ..\InterfaceTemplate\xml\05_TNCN_09.xml)
        rowData = 22
        rowGrid = 22
        Do
             ' Lay xem co nhung ca nhan nao duoc chon de in => varCheck ="1"
             .GetText .ColLetterToNumber("C"), .Row, varCheck
             If varCheck = "1" Then
                ' Voi nguoi dau tien duoc check thi chi can update du lieu do trong Object 05_TNCN_09
                If iFirst = 1 Then
                    ' Insert them mot dong du lieu vao Object 05_TNCN_09, thiet lap cac gia tri mac dinh theo cellType
                    updateData rowGrid, rowData
                    ' Dat lai iFirst = 2 de danh dau cho cac truong hop duoc check lan thu 2 tro di
                    iFirst = 2
                Else
                    ' Insert them mot dong du lieu vao Object 05_TNCN_09, thiet lap cac gia tri mac dinh theo cellType
                    InsertNode .ColLetterToNumber("D"), rowData
                    ' Cap nhat gia tri tren fps theo .Row vao dong du lieu thuc su trong Object theo rowData
                    updateData rowGrid, rowData
                End If
                
             End If
             .Col = .ColLetterToNumber("B")
             rowGrid = i + 22
             .Row = rowGrid
             i = i + 1
        Loop Until .Text = "aa"
    
        fps.EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub updateData(ByVal Row As Long, ByVal rowData As Long)
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        mCurrentSheet = .Sheet
        .Row = Row
        ' Cot D la cot chua Ho ten nguoi quyet toan
        .Col = .ColLetterToNumber("D")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot E la cot MST cua nguoi quyet toan
        .Col = .ColLetterToNumber("E")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot F la tong thu nhap chiu thue cua nguoi quyet toan
        .Col = .ColLetterToNumber("F")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot G la thu nhap luong cua nguoi quyet toan
        .Col = .ColLetterToNumber("G")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot H la cot thu nhap thuong cua nguoi quyet toan
        .Col = .ColLetterToNumber("H")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot I la cot thu nhap khac cua nguoi quyet toan
        .Col = .ColLetterToNumber("I")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot J la tong so nguoi phu thuoc cua nguoi quyet toan
        .Col = .ColLetterToNumber("J")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot K la cot tong so thang giam tru cua nguoi quyet toan
        .Col = .ColLetterToNumber("K")
        UpdateCell fps, .Col, rowData, .value
        
        'Cot L Thue da khau tru cua cua nguoi quyet toan
        .Col = .ColLetterToNumber("L")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot M la cot thu nhap chiu thue cua nguoi quyet toan
        .Col = .ColLetterToNumber("M")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot N la cot giam tru cho ban than cua nguoi quyet toan
        .Col = .ColLetterToNumber("N")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot O la cot giam tru cho nguoi phu thuoc cua nguoi quyet toan
        .Col = .ColLetterToNumber("O")
        UpdateCell fps, .Col, rowData, .value
                
        ' Cot P la cot thu nhap tinh thue cua nguoi quyet toan
        .Col = .ColLetterToNumber("P")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot Q la cot thue phai nop cua nguoi quyet toan
        .Col = .ColLetterToNumber("Q")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot R la cot thue mien giam cua nguoi quyet toan
        .Col = .ColLetterToNumber("R")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot S la cot thu nhap khac cua nguoi quyet toan
        .Col = .ColLetterToNumber("S")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
        .Col = .ColLetterToNumber("T")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
        .Col = .ColLetterToNumber("U")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot V la cot thu nhap do lam viec trong khu kinh te cua nguoi quyet toan
        .Col = .ColLetterToNumber("V")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
        .Col = .ColLetterToNumber("W")
        UpdateCell fps, .Col, rowData, .value
        
        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
        .Col = .ColLetterToNumber("X")
        UpdateCell fps, .Col, rowData, .value
    End With
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub
