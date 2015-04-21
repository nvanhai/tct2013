VERSION 5.00
Object = "{CF75B519-FBCE-4FF9-A3A9-1CA0AAC58B2C}#1.0#0"; "TBarCode5.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmReportData 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin FPUSpreadADO.fpSpread fpsReport 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _Version        =   458752
      _ExtentX        =   20558
      _ExtentY        =   12303
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmReportData.frx":0000
   End
   Begin TBARCODE5LibCtl.TBarCode5 TBarCode 
      Height          =   1095
      Left            =   3750
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      _cx             =   4683
      _cy             =   1931
      BackColor       =   16777215
      BackStyle       =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   "Adjust Properties"
      BarCode         =   55
      CDMethod        =   1
      CountCheckDigits=   0
      EscapeSequences =   0   'False
      Format          =   ""
      GuardWidth      =   0
      ModulWidth      =   ""
      Orientation     =   0
      PrintDataText   =   -1  'True
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:5B:6B:7B:8B:1S:2S:3S:4S:5S:6S"
      RatioDefault    =   "1:2:3:4:5:6:7:8:1:2:3:4:5:6"
      TextColor       =   0
      LastError       =   "The operation completed successfully. "
      LastErrorNo     =   0
      MustFit         =   0   'False
      TextDistance    =   -1
      NotchHeight     =   -1
      PDF417_Rows     =   -1
      PDF417_Columns  =   -1
      PDF417_ECLevel  =   -1
      PDF417_RowHeight=   -1
      MAXI_Mode       =   4
      MAXI_AppendIndex=   -1
      MAXI_AppendCount=   -1
      MAXI_Undercut   =   -1
      MAXI_Preamble   =   0   'False
      MAXI_PostalCode =   ""
      MAXI_CountryCode=   ""
      MAXI_ServiceClass=   ""
      MAXI_Date       =   "96"
      CountModules    =   840
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   7
      DM_Size         =   0
      DM_Rectangular  =   0   'False
      DM_Format       =   0
      DM_AppendIndex  =   -1
      DM_AppendCount  =   -1
      DM_AppendFileID =   -1
      PDF417_SegmentIndex=   -1
      PDF417_FileID   =   ""
      PDF417_LastSegment=   0   'False
      PDF417_FileName =   ""
      PDF417_SegmentCount=   -1
      PDF417_TimeStamp=   -1
      PDF417_Sender   =   ""
      PDF417_Addressee=   ""
      PDF417_FileSize =   -1
      PDF417_CheckSum =   -1
      QR_Version      =   0
      QR_Format       =   0
      QR_FmtAppIndicator=   ""
      QR_ECLevel      =   1
      QR_Mask         =   -1
      QR_AppendIndex  =   -1
      QR_AppendCount  =   -1
      QR_AppendParity =   -1
      PDF417_RatioRowCol=   ""
      InterpretInputAs=   0
      OptResolution   =   0   'False
      CBF_Rows        =   -1
      CBF_Columns     =   -1
      CBF_RowHeight   =   -1
      CBF_RowSeparatorHeight=   -1
      CBF_Format      =   0
      DisplayText     =   ""
      BarWidthReduction=   -1
      TextAlignment   =   0
      Quality         =   1702057308
   End
End
Attribute VB_Name = "frmReportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' Company           : CMC Soft
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Interface
' Form, Module
'   or Class name   : frmReportData
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

Private xmlDocumentInit()   As MSXML.DOMDocument    '
Private intCurrentBarcode As Integer                'Current barcode image
Private lBarcodeNumber As Integer                   'Number of barcode images
Private lPageNumber As Integer                      'Number of pages
Private intCurrPage As Integer                      '
Private blnHasPagePrinted As Boolean
Private arrStrPrintedPages() As String              'Array of printed pages
Private objTaxBusiness      As Object               'Private business object (clsReport001, clsReport002, clsReport003, ...)

Private isDLT As Boolean


'****************************************************
'Description:Form_Load procedure initialize the values of controls
'   Step 1: Load excel template to fpsReport grid.
'   Step 2: Setup fpsReport grid.
'   Step 3: Fill fpsReport grid with data.
'   Step 4: Create an object and assign fpsReport grid to it.
'****************************************************
Private Sub Form_Load()

On Error GoTo ErrHandle
    Dim fso As New FileSystemObject
    ' Phuc vu in thong tin DLT
    Dim strTemp As String
    Dim vTemp As Variant
    Dim arrSheet() As String
    Dim arrRow() As String
    ' End
    
    fpsReport.hDCPrinter = Printer.hDC
    
    If fso.FileExists("..\InterfaceTemplates\Template.xls") Then
'        If fpsReport.IsExcelFile("..\InterfaceTemplates\Template.xls") Then
'            fpsReport.ImportExcelBook GetAbsolutePath("..\InterfaceTemplates\Template.xls"), vbNullString
'        End If
        fpsReport.LoadFromFile "..\InterfaceTemplates\Template.xls"
    End If
    
    Set fso = Nothing
    
    LoadTemplate fpsReport, False 'Load Report Template
    SetupSpread ' Initialize Spread grid
    LoadInitFiles
    SetupReportData fpsReport, False 'Load data to grid
    
    ' Setup Spread trong truong hop in bo sung cac to khai quyet toan
    If flgPrintBoSung = False Then GoTo nextPrinter
    Dim i As Integer
    Dim test As Boolean
    Dim varTemp As String
    Dim listInBoSung5A_Copy() As String
    Dim listInBoSung5B_Copy() As String
    Dim j    As Integer, j2 As Integer
    
    With fpsReport
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "17" Then
            .sheet = 2
            .Row = 15
            ReDim Preserve listInBoSung5A_Copy(UBound(listInBoSung5A))
            ReDim Preserve listInBoSung5B_Copy(UBound(listInBoSung5B))
            'get so phan tu mang tren 5A va 5B
            j = UBound(listInBoSung5A)
            j = j - 1
            j2 = UBound(listInBoSung5B)
            j2 = j2 - 1
            
            'copy mang va dao chieu
            'copy 5A
            For i = 0 To UBound(listInBoSung5A) - 1
               If InStr(1, listInBoSung5A(j), ".") > 1 Then
                    listInBoSung5A_Copy(i) = Replace(listInBoSung5A(j), ".", "")
               Else
                    listInBoSung5A_Copy(i) = listInBoSung5A(j)
               End If
               j = j - 1
            Next
            'copy 5B
            For i = 0 To UBound(listInBoSung5B) - 1
               If InStr(1, listInBoSung5B(j2), ".") > 1 Then
                    listInBoSung5B_Copy(i) = Replace(listInBoSung5B(j2), ".", "")
               Else
                    listInBoSung5B_Copy(i) = listInBoSung5B(j2)
               End If
               j2 = j2 - 1
            Next
            Do
                .Col = .ColLetterToNumber("B")

                For i = UBound(listInBoSung5A_Copy) - 1 To 0 Step -1

                    If Trim(.Text) = Trim(listInBoSung5A_Copy(i)) Then
                        .RowHidden = False
                        test = True
                        If i = UBound(listInBoSung5A_Copy) - 1 And UBound(listInBoSung5A_Copy) <> 1 Then
                            listInBoSung5A_Copy(i) = Empty
                            ReDim Preserve listInBoSung5A_Copy(UBound(listInBoSung5A_Copy) - 1)
                        End If
                        Exit For
                    Else
                       ' .DeleteRows .Row, 0
                       '.Height = 0
                        .RowHidden = True
                    End If
                Next
                
                                
                .Row = .Row + 1
                
            Loop Until .Text = "aa"
            
            ' Visible dong cam ket la du lieu to khai dung
            .Row = .Row + 2
            .RowHidden = True
            
            ' Enable dong cam ket la du lieu chi cua to khai bo sung
            .Row = .Row + 1
            .RowHidden = False
            
            .sheet = 3
            .Row = 15
            Do
                .Col = .ColLetterToNumber("B")

                For i = UBound(listInBoSung5B_Copy) - 1 To 0 Step -1

                    If Trim(.Text) = Trim(listInBoSung5B_Copy(i)) Then
                        .RowHidden = False
                        test = True
                        If i = UBound(listInBoSung5B_Copy) - 1 And UBound(listInBoSung5B_Copy) <> 1 Then
                            listInBoSung5B_Copy(i) = Empty
                            ReDim Preserve listInBoSung5B_Copy(UBound(listInBoSung5B_Copy) - 1)
                        End If
                        Exit For
                    Else
                        .RowHidden = True
                    End If
                Next
                .Row = .Row + 1
                
            Loop Until .Text = "aa"
            ' Visible dong cam ket la du lieu to khai dung
            .Row = .Row + 2
            .RowHidden = True
            
            ' Enable dong cam ket la du lieu chi cua to khai bo sung
            .Row = .Row + 1
            .RowHidden = False
        End If
        
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "42" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "43" Then
            .sheet = 2
            .Row = 39

            Do
                .Col = .ColLetterToNumber("B")
                For i = 0 To UBound(listInBoSung2A)
                    If Trim(.Text) = Trim(listInBoSung2A(i)) Then
                        .RowHidden = False
                        Exit For
                    Else
                        .RowHidden = True
                    End If
                Next
                .Row = .Row + 1
                
            Loop Until .Text = "aa"
            ' Visible dong cam ket la du lieu to khai dung
            .Row = .Row + 2
            .RowHidden = True
            
            ' Enable dong cam ket la du lieu chi cua to khai bo sung
            .Row = .Row + 1
            .RowHidden = False
        End If
        
        ' to khai 06 KK TNCN
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "59" Then
            .sheet = 2
            .Row = 15
            Do
                .Col = .ColLetterToNumber("B")
                
                For i = 0 To UBound(listInBoSung6B)
                    If Trim(.Text) = Trim(listInBoSung6B(i)) Then
                       .RowHidden = False
                       test = True
    
                        Exit For
                    Else
                       ' .DeleteRows .Row, 0
                       '.Height = 0
                        .RowHidden = True
                    End If
                Next
                .Row = .Row + 1
                
            Loop Until .Text = "aa"
            
            ' Visible dong cam ket la du lieu to khai dung
            .Row = .Row + 2
            .RowHidden = True
            
            ' Enable dong cam ket la du lieu chi cua to khai bo sung
            .Row = .Row + 1
            .RowHidden = False
        End If

        
    End With
    
    ' End Setup Spread trong truong hop in bo sung cac to khai quyet toan


'    If GetAttribute(TAX_Utilities_v2.NodeValidity, "Class2") <> vbNullString Then
'        Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_v2.NodeValidity, "Class2"))
'        Set objTaxBusiness.fpsGrid = fpsReport
'        objTaxBusiness.Process
'    End If
    

' Print report footer
    
nextPrinter:
    Dim font1 As String
    'font1 = "/fn""Arial""/fz""9""/fb1/fi1/fu1"
    'dhdang sua in check gia han
    'ngay 19/04/2011
    fpsReport.sheet = 1
    If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "11") Or (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "12") Then
     'dhdang sua check gia han nop theo thong tu
         ' q Quy ke khai
        ' y nam ke khai
        ' dNgayTaiChinh ngay tai chinh lay tren man hinh HTKK
        ' dThangTaiChinh thang tai chinh tren phan thong tin chung HTKK
        ' dType: 0 tra ve quy, 1 tra ve nam
        Dim Quynamtc As Integer
        Dim Namtc As Integer
        Dim QuyKK As Integer
        Dim NamKK As Integer
        Dim iNgayTaiChinh As Integer
        Dim iThangTaiChinh As Integer
        Dim strNgayTaiChinh As String
        Dim Thongtu         As String
        Dim ThangKK As Integer
        
        Dim rowCheckGH As Long
        Dim colCheckGH As String
'
'        If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "11") Then
'            rowCheckGH = 38
'            colCheckGH = "H"
'        Else
'            rowCheckGH = 46
'            colCheckGH = "G"
'        End If
'
'        QuyKK = TAX_Utilities_v2.ThreeMonths
'        NamKK = TAX_Utilities_v2.Year
'
        ' ID to khai
        Dim idToKhai As String
        Dim dsTK_checkDLT As String
'
'        'Lay ngay bat dau nam tai chinh
'        If GetAttribute(TAX_Utilities_v2.NodeMenu, "FinanceYear") = "1" Then
'            strNgayTaiChinh = GetNgayBatDauNamTaiChinh
'            iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
'            iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
'        Else
'            strNgayTaiChinh = "01/01"
'            iNgayTaiChinh = 1
'            iThangTaiChinh = 1
'        End If
'
'        Quynamtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 0)
'        Namtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 1)
'
'        If (Namtc = 2010 And Quynamtc = 4) Then
'            Thongtu = GetAttribute(GetMessageCellById("0187"), "Msg")
'            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'        ElseIf Namtc = 2011 And (Quynamtc = 1 Or Quynamtc = 2) Then
'            Thongtu = GetAttribute(GetMessageCellById("0260"), "Msg")
'            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'        ElseIf Namtc = 2011 And (Quynamtc = 3 Or Quynamtc = 4) Then
'            Thongtu = GetAttribute(GetMessageCellById("0243"), "Msg")
'            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'        ElseIf Namtc = 2010 And Quynamtc <> 4 Then
'            Thongtu = GetAttribute(GetMessageCellById("0186"), "Msg")
'            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'        ElseIf Namtc = 2009 Then
'            Thongtu = GetAttribute(GetMessageCellById("0185"), "Msg")
'            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
''        ElseIf Namtc = 2013 And (Quynamtc = 1 Or Quynamtc = 2 Or Quynamtc = 3) Then
''            Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
''            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'        Else
''            fpsReport.SetText fpsReport.ColLetterToNumber("B"), rowCheckGH, ""
''            fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'             Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
'             fpsReport.SetText fpsReport.ColLetterToNumber(colCheckGH), rowCheckGH, Thongtu
'        End If

    ElseIf (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "03") Then
'        If TAX_Utilities_v2.Year = 2009 Then
'        fpsReport.SetText fpsReport.ColLetterToNumber("E"), 25, GetAttribute(GetMessageCellById("0185"), "Msg")
'        ElseIf TAX_Utilities_v2.Year = 2010 Then
'            fpsReport.SetText fpsReport.ColLetterToNumber("E"), 25, GetAttribute(GetMessageCellById("0187"), "Msg")
'        ElseIf TAX_Utilities_v2.Year = 2011 Then
'            fpsReport.SetText fpsReport.ColLetterToNumber("E"), 25, GetAttribute(GetMessageCellById("0243"), "Msg")
''        ElseIf TAX_Utilities_v2.Year = 2013 Then
''            fpsReport.SetText fpsReport.ColLetterToNumber("E"), 25, GetAttribute(GetMessageCellById("0263"), "Msg")
''        ElseIf Val(TAX_Utilities_v2.Year) = 2011 Or Val(TAX_Utilities_v2.Year) > 2013 Then
''            fpsReport.SetText fpsReport.ColLetterToNumber("E"), 25, ""
''            fpsReport.Row = 25
''            fpsReport.Col = fpsReport.ColLetterToNumber("C")
''            fpsReport.Text = ""
''            fpsReport.Lock = True
'        Else
'            fpsReport.SetText fpsReport.ColLetterToNumber("E"), 25, GetAttribute(GetMessageCellById("0263"), "Msg")
'        End If
    ElseIf (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "01") Then
        If TAX_Utilities_v2.Year = 2012 And (TAX_Utilities_v2.month = 4 Or TAX_Utilities_v2.month = 5 Or TAX_Utilities_v2.month = 6) Then
            fpsReport.SetText fpsReport.ColLetterToNumber("G"), 33, GetAttribute(GetMessageCellById("0262"), "Msg")
            fpsReport.Row = 34
            fpsReport.Col = fpsReport.ColLetterToNumber("O")
            fpsReport.FontBold = False
'        ElseIf TAX_Utilities_v2.Year = 2013 And (TAX_Utilities_v2.month = 1 Or TAX_Utilities_v2.month = 2 Or TAX_Utilities_v2.month = 3) Then
'            fpsReport.SetText fpsReport.ColLetterToNumber("G"), 33, GetAttribute(GetMessageCellById("0263"), "Msg")
'            fpsReport.Row = 34
'            fpsReport.Col = fpsReport.ColLetterToNumber("C")
'            fpsReport.Text = ""
'            fpsReport.Lock = True
'            fpsReport.Col = fpsReport.ColLetterToNumber("O")
'            fpsReport.Text = ""
'            fpsReport.Lock = True
        Else
'            fpsReport.SetText fpsReport.ColLetterToNumber("G"), 33, ""
'            fpsReport.Row = 34
'            fpsReport.Col = fpsReport.ColLetterToNumber("C")
'            fpsReport.Text = ""
'            fpsReport.Lock = True
'            fpsReport.Col = fpsReport.ColLetterToNumber("O")
'            fpsReport.Text = ""
'            fpsReport.Lock = True
'            fpsReport.Row = 33
'            fpsReport.RowHidden = True
            fpsReport.SetText fpsReport.ColLetterToNumber("G"), 33, GetAttribute(GetMessageCellById("0262"), "Msg")
            fpsReport.Row = 34
            fpsReport.Col = fpsReport.ColLetterToNumber("O")
            fpsReport.FontBold = False
'            fpsReport.SetText fpsReport.ColLetterToNumber("G"), 33, GetAttribute(GetMessageCellById("0263"), "Msg")
            fpsReport.Row = 34
            fpsReport.Col = fpsReport.ColLetterToNumber("C")
            fpsReport.SetText fpsReport.ColLetterToNumber("C"), 34, GetAttribute(GetMessageCellById("0264"), "Msg")
            fpsReport.FontSize = 6
            fpsReport.Col = fpsReport.ColLetterToNumber("O")
            fpsReport.FontSize = 7
'            fpsReport.RowHeight(34) = 20
'            fpsReport.Row = 67
'            fpsReport.RowHeight(67) = 45
            fpsReport.Lock = True
'            fpsReport.Col = fpsReport.ColLetterToNumber("O")
'            fpsReport.Text = ""
'            fpsReport.Lock = True
        End If
    ElseIf (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73") Then
        fpsReport.SetText fpsReport.ColLetterToNumber("E"), 83, GetAttribute(GetMessageCellById("0265"), "Msg")
        fpsReport.Row = 83
        fpsReport.Col = fpsReport.ColLetterToNumber("E")
        fpsReport.FontBold = False
        ' kiem tra chi tieu 36,42,43
        Dim ct36 As Variant
        fpsReport.GetText fpsReport.ColLetterToNumber("AV"), 72, ct36
        If Val(ct36) > 0 Then
            fpsReport.SetText fpsReport.ColLetterToNumber("BD"), 72, "%"
        Else
            fpsReport.SetText fpsReport.ColLetterToNumber("BD"), 72, ""
        End If
        
        
    End If
    'end
    '------------------
    font1 = "/fn""Arial""/fz""8""/fb1/fi1/fu1"
    
       
    
    If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 15 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 16 Then
      fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0178"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 53 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 37 Then
      fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0128"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 46 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 47 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 48 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 49 Then
      fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0190"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 50 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 51 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0179"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0180"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0181"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0182"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0287"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 54 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 38 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0129"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0130"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0131"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0132"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 39 Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 40 Then
      fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0133"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0136"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 36 Then
        'fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0134"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0189"), "Msg")
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0134"), "Msg")
        ' vttoan: them ma muc to 01_NTNN (thue nha thau)
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 70 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0191"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0192"), "Msg")
        ' nkhoan
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 81 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0258"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 82 Then
        'fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0259"), "Msg")
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0289"), "Msg")
        ' vttoan: them ma tieu muc 03_GTGT
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0193"), "Msg")
        ' to khai 03B_GTGT
'    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "95" Then
'        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0193"), "Msg")
'        'dntai: them tieu muc to 01_TTDB , 01A_TNDN, 01B_TNDN , 01_GTGT
        ' to khai 03A_GTGT
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "88" Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0288"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" And strLoaiNNKD <> "0" Then
      fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0194"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "11" Then
        If strLoaiNNKD = "1052" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0195"), "Msg")
        ElseIf strLoaiNNKD = "1056" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0196"), "Msg")
        ElseIf strLoaiNNKD = "1057" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0197"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "12" Then
        If strLoaiNNKD = "1052" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0198"), "Msg")
        ElseIf strLoaiNNKD = "1056" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0199"), "Msg")
        ElseIf strLoaiNNKD = "1057" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0200"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "03" Then
        If strLoaiNNKD = "1052" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0249"), "Msg")
        ElseIf strLoaiNNKD = "1056" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0250"), "Msg")
        ElseIf strLoaiNNKD = "1057" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0251"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Then
        If strLoaiNNKD = "1701" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0201"), "Msg")
        ElseIf strLoaiNNKD = "1704" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0202"), "Msg")
        ElseIf strLoaiNNKD = "1705" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0203"), "Msg")
        End If
        
'        ' PL 01-3
'        ' format lai cac cot so tien ve text
'        Dim idx As Long
'        fpsReport.sheet = 4
'        idx = 20
'        Do
'            fpsReport.Row = idx
'            fpsReport.Col = fpsReport.ColLetterToNumber("M")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("AJ")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("AY")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("BK")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("BU")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("CC")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("CK")
'            fpsReport.CellType = 1
'            fpsReport.TypeVAlign = TypeVAlignCenter
'            'fpsReport.TypeTextWordWrap = True
'
'            fpsReport.Col = fpsReport.ColLetterToNumber("B")
'            idx = idx + 1
'        Loop Until fpsReport.Text = "aa"
        
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "73" Then
'        If strLoaiNNKD = "1054" Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0348"), "Msg")
'        ElseIf strLoaiNNKD = "1053" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0348"), "Msg")
        'End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "56" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0355"), "Msg")
    'dntai:them tieu muc to 06_TNCN
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 59 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0231"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0232"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0233"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0234"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0332"), "Msg")
    '01B/TNDN-DK
'    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 99 Then
'        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0277"), "Msg")
    'set footer for 01/TAIN-DK,02/TAIN-DK, 02/TNDN-DK
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 98 Then
'        If (strDauTho = "1") Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0275"), "Msg")
'        ElseIf (strCondensate = "1") Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0276"), "Msg")
'        ElseIf (strKhiThienNhien = "1") Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0277"), "Msg")
'        End If
        'fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0300"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0301"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0302"), "Msg")
        If strLoaiTkDk = "DT" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0300"), "Msg")
        ElseIf strLoaiTkDk = "KTN" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0302"), "Msg")
        ElseIf strLoaiTkDk = "CD" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0301"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 92 Then
'        If (strDauTho = "1") Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0275"), "Msg")
'        ElseIf (strCondensate = "1") Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0276"), "Msg")
'        ElseIf (strKhiThienNhien = "1") Then
'            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0277"), "Msg")
'        End If
        'fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0300"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0301"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0302"), "Msg")
        If strLoaiTkDk = "DT" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0303"), "Msg")
        ElseIf strLoaiTkDk = "KTN" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0305"), "Msg")
        ElseIf strLoaiTkDk = "CD" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0304"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 93 Then
        If strLoaiTkDk = "DT" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0326"), "Msg")
        ElseIf strLoaiTkDk = "KTN" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0328"), "Msg")
        ElseIf strLoaiTkDk = "CD" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0327"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 89 Then
        If strLoaiTkDk = "DT" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0329"), "Msg")
        ElseIf strLoaiTkDk = "KTN" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0331"), "Msg")
        ElseIf strLoaiTkDk = "CD" Then
            fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0330"), "Msg")
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 99 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0306"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 25 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0284"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0285"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 94 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0307"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 96 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0308"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 25 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0322"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0323"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 26 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0320"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0321"), "Msg")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = 80 Then
        fpsReport.PrintFooter = font1 & GetAttribute(GetMessageCellById("0127"), "Msg") & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0333"), "Msg") & "/n" & GetAttribute(GetMessageCellById("0334"), "Msg")
    End If
    
    'KHBS khong thay doi so thue se ko in phu luc va co cau canh bao
    idToKhai = GetAttribute(TAX_Utilities_v2.NodeMenu, "ID")

'    If strKHBS = "TKBS" And (idToKhai = "01" Or idToKhai = "02" Or idToKhai = "04" Or idToKhai = "11" Or idToKhai = "12" Or idToKhai = "06" Or idToKhai = "05" Or idToKhai = "03" Or idToKhai = "71" Or idToKhai = "72" Or idToKhai = "73" Or idToKhai = "77" Or idToKhai = "80" Or idToKhai = "81" Or idToKhai = "82" Or idToKhai = "83" Or idToKhai = "85" Or idToKhai = "86" Or idToKhai = "87" Or idToKhai = "89" _
'    Or idToKhai = "90" Or idToKhai = "95" Or idToKhai = "70") Then
    If strKHBS = "TKBS" And (InStr(1, strIdKHBS_TT156, "~" & Trim$(idToKhai) & "~", vbTextCompare) > 0) Or idToKhai = "01" Then
        If objTaxBusiness Is Nothing Then
            Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_v2.NodeValidity, "Class"))
            objTaxBusiness.SetActiveKHBS
            If objTaxBusiness.FlagThaydoi = "1" Then
                fpsReport.PrintFooter = fpsReport.PrintFooter & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0226"), "Msg")
            End If
        Else
            objTaxBusiness.SetActiveKHBS
            If objTaxBusiness.FlagThaydoi = "1" Then
                fpsReport.PrintFooter = fpsReport.PrintFooter & "/n/fb0/fi0/fu0" & GetAttribute(GetMessageCellById("0226"), "Msg")
            End If
        End If
    End If
        
    ' Kiem tra neu khong co DLT se an cac row DLT di
    dsTK_checkDLT = "~01~02~04~71~72~11~12~73~03~46~47~48~49~25~15~16~50~51~36~74~75~70~80~81~82~17~42~43~26~59~76~41~06~77~05~86~87~85~88~90~23~94~96~97~98~99~93~92~89~24~55~56~84~"

    If InStr(1, dsTK_checkDLT, "~" & Trim$(idToKhai) & "~", vbTextCompare) > 0 Then
        If objTaxBusiness Is Nothing Then
            Set objTaxBusiness = CreateObject(GetAttribute(TAX_Utilities_v2.NodeValidity, "Class"))
        End If
        objTaxBusiness.inThongTinDL
        isDLT = objTaxBusiness.FlagDLThue
        strTemp = objTaxBusiness.strTuRowDenRowPL
        If strTemp <> "" Then
            arrSheet = Split(strTemp, "~")
        End If
        ' To khai
        If isDLT = False Then
            For i = objTaxBusiness.tuRowDL To objTaxBusiness.denRowDL
                fpsReport.sheet = 1
                fpsReport.Row = i
                fpsReport.RowHidden = True
            Next
        End If
        ' PL
        If isDLT = False Then
            For Each vTemp In arrSheet
                arrRow = Split(CStr(vTemp), "_")
                For i = CLng(arrRow(0)) To CLng(arrRow(1))
                    fpsReport.sheet = CLng(arrRow(2))
                    fpsReport.Row = i
                    fpsReport.RowHidden = True
                Next
            Next
        End If
        
        ' kiem tra to khai 02/TNDN
        If idToKhai = "73" Then
            If objTaxBusiness.isCheckGH = True Then
                fpsReport.sheet = 1
                fpsReport.Row = 83
                fpsReport.RowHidden = True
                fpsReport.Col = fpsReport.ColLetterToNumber("C")
                fpsReport.Text = ""
            End If
        End If
    End If
    
    
    SetSheetVisible fpsReport
    
    SetupPrinter
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
'    SetFormCaption Me, imgCaption, lblCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intCurrentBarcode = 0
    lBarcodeNumber = 0
    lPageNumber = 0
    intCurrPage = 0
    ReDim arrStrPrintedPages(0)
    Set objTaxBusiness = Nothing
End Sub

Private Sub fpsReport_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    fpsReport.sheet = NewSheet
End Sub

'**********************************************
'Description: Format fpsGrid
'   Print align: Left=0.25 inch, Top=0.25 inch, Bottom=0.25 inch
'   Number type: Value from -99999999999999 to 9999999999999
'**********************************************
Private Sub SetupSpread()
On Error GoTo ErrHandle
    Dim intCtrl As Integer
    Dim intCol As Integer, intRow As Integer
    Dim vConfirm As Boolean
    With fpsReport
    vConfirm = False
        For intCtrl = 1 To fpsReport.SheetCount
            .sheet = intCtrl ' Set Sheet index
            .ColHeadersShow = False ' invisible ColHeader
            .RowHeadersShow = False 'Invisible RowHeader
            .MaxCols = .DataColCnt - 1 ' Number of col contain data
            .MaxRows = .DataRowCnt - 1 ' Number of row contain data
            .PrintCenterOnPageH = True
            'Set margin left and margin top
            
            .PrintMarginLeft = 0.125 * 1440
            
            .PrintMarginTop = 0.75 * 1440
            .PrintMarginBottom = 0.5 * 1440
            .PrintUseDataMax = False
         
'            If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = 17 Then
'                If Len(GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("P_15"), "Value")) > 10 Or _
'                   Len(GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("P_16"), "Value")) > 10 Then
'                        If vConfirm = False Then
'                            DisplayMessage "0121", msOKOnly, miWarning
'                            vConfirm = True
'                        End If
'                        Printer.PaperSize = vbPRPSA3
'                End If
'            End If
            ' BC26
'            If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Caption") = "BC26-AC" Then
'                Printer.PaperSize = vbPRPSA3
'                Printer.Orientation = vbPRORLandscape
'            End If
            ' end
            ' to khai thue nha thau in theo kho giay nam ngang
            If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "70" Then
                Printer.PaperSize = vbPRPSA4
                Printer.Orientation = vbPRORLandscape
            End If
                            
            
            For intCol = 1 To .MaxCols
                For intRow = 1 To .MaxRows
                    .Col = intCol
                    .Row = intRow
                    If .CellType = CellTypeNumber Then
                        .TypeNumberSeparator = "."
                        .TypeNumberDecimal = ","
                        .TypeNumberMax = 99999999999999#
                        .TypeNumberMin = -99999999999999#
                        .TypeNumberNegStyle = TypeNumberNegStyle1
                    ElseIf .CellType = CellTypeEdit Then
                        .TypeEditMultiLine = True
                    End If
                    
                Next intRow
                
'                 If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = 17 And intCtrl = 1 And .Row > 12 _
'                 And (Len(GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("P_15"), "Value")) > 10 Or Len(GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("P_16"), "Value")) > 10) Then
'                    If intCol = .ColLetterToNumber("C") Or intCol = .ColLetterToNumber("G") Or intCol = .ColLetterToNumber("CP") Then
'                        .ColWidth(intCol) = .ColWidth(intCol) + 8
'                        .PrintCenterOnPageH = True
'                    End If
'                 End If
'                 If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = 17 And intCtrl = 2 _
'                 And (Len(GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("P_15"), "Value")) > 10 Or Len(GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("P_16"), "Value")) > 10) Then
'                       .ColWidth(intCol) = .ColWidth(intCol) + 1.2
'                 End If
                 
            Next intCol
            
        Next intCtrl
    End With
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "SetupSpread", Err.Number, Err.Description
End Sub

'****************************************************
'Description:GetAbsolutePossitionInTwips procedure convert the position
'   of mark cell to the absolute position on the paper.
'   Step 1: Find the mark cell by "BarcodeImage" string
'   Step 2: Evaluate the position of mark cell in twips
'Output:
'   lXOffset: the X cordinate of barcode position
'   lYOffset: the Y cordinate of barcode position
'****************************************************
Private Sub GetAbsolutePossitionInTwips(ByRef lXOffset As Long, ByRef lYOffset As Long) 'ByRef intPage As Integer, ByRef lXOffset As Long, ByRef lYOffset As Long)
Dim lWidth As Long, lHeight As Long, lCol As Long, lRow As Long
Dim lCtrl As Long, lLeft As Long, lTop As Long
Dim lTemp As Long, lPageRowBreak As Long

With fpsReport
    'Get Col, Row position of cell containt barcode image
    For lRow = 1 To .MaxRows
        .Row = lRow
        For lCol = 1 To .MaxCols
            .Col = lCol
            If UCase(.Text) = UCase("BarcodeImage") Then
                .Text = ""
                GoTo exitFor
            End If
        Next lCol
    Next lRow
exitFor:

    'Not found barcode position
    If lCol = .MaxCols + 1 And lRow = .MaxRows + 1 Then
        ' If the mark did not find, resize margin top of grid
        ' and print barcode on top of paper
        'fpsReport.PrintMarginTop = 0.75 * 1440
        'DisplayMessage "0020", msOKOnly, miCriticalError
        Exit Sub
    End If
    'Get row that the page is broken and number of page
    Do
        lTemp = fpsReport.PrintNextPageBreakRow
        If lTemp <> -1 And lTemp <= lRow Then
            lPageRowBreak = lTemp
        End If
    Loop Until lTemp = -1 Or lTemp > lRow
    
    If lPageRowBreak <> 0 Then
        'Get top position in twips
        For lCtrl = lPageRowBreak To lRow - 1
            'Convert row height to twips
            .RowHeightToTwips lRow, .RowHeight(lCtrl), lHeight
            lTop = lTop + lHeight '+ 5 'Sai so cho chieu cao va do rong cua line
        Next lCtrl
    Else
        'Get top position in twips
        For lCtrl = 1 To lRow - 1
            'Convert row height to twips
            .RowHeightToTwips lRow, .RowHeight(lCtrl), lHeight
            lTop = lTop + lHeight '+ 5 'Sai so cho chieu cao va do rong cua line
        Next lCtrl
    End If
    
    'Get left position in twips
    For lCtrl = 1 To lCol - 1
        'Convert column width to twips
        .ColWidthToTwips .ColWidth(lCtrl), lWidth
        lLeft = lLeft + lWidth '+ 5 'Sai so cho chieu cao va do rong cua line
    Next lCtrl
        
    lXOffset = .PrintMarginLeft + lLeft
    lYOffset = .PrintMarginTop + lTop
    
End With
End Sub

'*****************************************************
'Description: Create and print the barcodes directly to printer
'   Step 1: Format size of barcode image:
'           width: 40 millimeters
'           height: 18 millimeters
'Input: strValue string (equal to or less the limited string
'       value of barcode encode Ocx component),
'       Column of the specified location,
'       Row of the specified location.
'*****************************************************
Private Sub PrintBarcodes(ByVal arrStrValue As Variant, ByVal intNumberOfBarcode As Integer, ByVal intStart As Integer) ', Optional blnEndOfSheet As Boolean = False)
On Error GoTo ErrHandle
    Dim intCtrl As Integer, intTemp As Integer
    Dim lXOffset As Long, lYOffset As Long
    Dim lXSize As Long, lYSize As Long
    Dim lXRange As Long, lYRange As Long
    
    Dim i As Integer    ' Variable for PDF Barcode of iHTKK
    Dim strPrefix As String     ' Variable for PDF Barcode of iHTKK
    Dim arrStrValueBarCode As Variant  ' Variable for PDF Barcode of iHTKK
    
    Printer.ScaleMode = vbPixels
'    Initialize params for printer
    
    lXSize = Printer.ScaleX(40, vbMillimeters)
    lYSize = Printer.ScaleY(18, vbMillimeters)
    ' doi voi TH co 4 MAVACH/1 trang
    If intNumberOfBarcode = 4 Then
        lXRange = Printer.ScaleX(6, vbMillimeters)
    Else
        lXRange = Printer.ScaleX(28, vbMillimeters)
    End If
    'lXRange = Printer.ScaleX(28, vbMillimeters)
    lYRange = Printer.ScaleY(5, vbMillimeters)
    If Printer.Orientation = 1 Then
        lXOffset = Printer.ScaleX(6.1 * 1440, vbTwips)
    Else
        lXOffset = Printer.ScaleX(9.5 * 1440, vbTwips)
    End If
    'lXOffset = Printer.ScaleWidth - Printer.ScaleX(0.6 * 1440, vbTwips) - lXSize
    lYOffset = Printer.ScaleY(0.025 * 1440, vbTwips)
    
    'Begin iHTKK
    ' Set gia tri default strBarcodeInPDF = vbnullString phuc vu cho iHTKK
    strBarcodeInPDF = vbNullString
    ' Chuan bi chuoi ma vach de in len tung trang dung cho iHTKK
    For intCtrl = intNumberOfBarcode - 1 To 0 Step -1
        If intStart + intCtrl <= UBound(arrStrValue) And IsPrintedPage(intCurrPage) Then
            ' Xu ly chuoi dau trong trang dau tien phai set lai So luong ma mach chua du lieu luon luon la 1
            ' Phuc vu cho iHTKK
            If intCurrPage = 1 And intStart = 1 Then
                strPrefix = GetPrefix(1) ' Lay thong tin cua header
                strPrefix = strPrefix & "001001" ' Luon luon la 1 ma vach chua tat ca cac du lieu
                If intCtrl = 0 Then
                    strBarcodeInPDF = strPrefix & Mid$(arrStrValue(intStart + intCtrl), 39, Len(arrStrValue(intStart + intCtrl)) - 39) & strBarcodeInPDF
                Else
                    strBarcodeInPDF = Mid$(arrStrValue(intStart + intCtrl), 39, Len(arrStrValue(intStart + intCtrl)) - 39) & strBarcodeInPDF
                End If
            Else
                strBarcodeInPDF = Mid$(arrStrValue(intStart + intCtrl), 39, Len(arrStrValue(intStart + intCtrl)) - 39) & strBarcodeInPDF
            End If
            ' Ket thuc lay chuoi ma vach phuc vu cho iHTKK
        End If
    Next intCtrl
    ' Ket thuc chuan bi chuoi ma vach de in len tung trang dung cho iHTKK
    
    ' Debug.Print strBarcodeInPDF phuc vu cho iHTKK
    ' Thiet lap toa do (0, 0) in Barcode len Header cua tung trang phuc vu cho iHTKK
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 1
    
    ' Begin print Barcode into page 1 or pages in PDF Barcode of iHTKK
    ' Ghep them the nhan dang ma vach <TCT-BARCODE> chuoi ma vach </TCT-BARCODE>
    strBarcodeInPDF = "<TCT-BARCODE>" & strBarcodeInPDF & "</TCT-BARCODE>"
    arrStrValueBarCode = CutStringByNumChar(strBarcodeInPDF, 124)
    For i = 1 To UBound(arrStrValueBarCode)
        Printer.FontSize = 1
        Printer.ForeColor = vbWhite
        Printer.Print arrStrValueBarCode(i)
        Printer.ForeColor = vbBlack
    Next
    ' End print Barcode into page 1 or pages in PDF Barcode of iHTKK
    ' End iHTKK


    ' Neu la to quyet toan TNCN thi ko in ma vach, Cac to khai khac van in ma vach binh thuong
    'If (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "17" Or (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "41" And isDLT = True)) And (intCurrPage = 1 Or intCurrPage = 2) Then
    If (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "17") And (intCurrPage = 1 Or intCurrPage = 2) Then
'   Print right align
        For intCtrl = intNumberOfBarcode - 1 To 0 Step -1
            If intStart + intCtrl <= UBound(arrStrValue) And IsPrintedPage(intCurrPage) Then
                PrintNormalBarcode arrStrValue(intStart + intCtrl), lXOffset, lYOffset, lXSize, lYSize
                lXOffset = lXOffset - lXSize - lXRange
            End If
        Next intCtrl
    'ElseIf (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "42" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "43" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "59" Or (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "41" And isDLT = False)) And intCurrPage = 1 Then
    ElseIf (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "42" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "43") And intCurrPage = 1 Then
'   Print right align
        For intCtrl = intNumberOfBarcode - 1 To 0 Step -1
            If intStart + intCtrl <= UBound(arrStrValue) And IsPrintedPage(intCurrPage) Then
                PrintNormalBarcode arrStrValue(intStart + intCtrl), lXOffset, lYOffset, lXSize, lYSize
                lXOffset = lXOffset - lXSize - lXRange
            End If
        Next intCtrl
    ElseIf (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "59" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "26") And ((isDLT = False And intCurrPage = 1) Or (isDLT = True And (intCurrPage = 1 Or intCurrPage = 2))) Then
'   Print right align
        For intCtrl = intNumberOfBarcode - 1 To 0 Step -1
            If intStart + intCtrl <= UBound(arrStrValue) And IsPrintedPage(intCurrPage) Then
                PrintNormalBarcode arrStrValue(intStart + intCtrl), lXOffset, lYOffset, lXSize, lYSize
                lXOffset = lXOffset - lXSize - lXRange
            End If
        Next intCtrl
    'ElseIf (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "17" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "41" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "42" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "43" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "59") Then
    ElseIf (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "17" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "42" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "43" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "59" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "26") Then
'   Print right align
        For intCtrl = intNumberOfBarcode - 1 To 0 Step -1
            If intStart + intCtrl <= UBound(arrStrValue) And IsPrintedPage(intCurrPage) Then
                PrintNormalBarcode arrStrValue(intStart + intCtrl), lXOffset, lYOffset, lXSize, lYSize
                lXOffset = lXOffset - lXSize - lXRange
            End If
        Next intCtrl
    End If
''   Print last barcodes on the new page
'    If blnEndOfSheet And intStart + intNumberOfBarcode <= UBound(arrStrValue) Then
'        'print new page
'        PrintNewPage
'
'        intTemp = 0 'Count amount of barcode images in one row
'        'lXOffset = Printer.ScaleX(6 * 1440, vbTwips)
'        If Printer.Orientation = 1 Then
'            lXOffset = Printer.ScaleX(6.1 * 1440, vbTwips)
'        Else
'            lXOffset = Printer.ScaleX(9.5 * 1440, vbTwips)
'        End If
'        For intCtrl = UBound(arrStrValue) To intStart + intNumberOfBarcode Step -1
'            intTemp = intTemp + 1
'
'            If IsPrintedPage(intCurrPage) Then _
'                PrintNormalBarcode arrStrValue(intCtrl), lXOffset, lYOffset, lXSize, lYSize
'
'            If intTemp > intNumberOfBarcode Then
'                lYOffset = lYOffset + lYSize + lYRange
'                'lXOffset = Printer.ScaleX(6 * 1440, vbTwips)
'                If Printer.Orientation = 1 Then
'                    lXOffset = Printer.ScaleX(6.1 * 1440, vbTwips)
'                Else
'                    lXOffset = Printer.ScaleX(9.5 * 1440, vbTwips)
'                End If
'            Else
'                lXOffset = lXOffset - lXSize - lXRange
'            End If
'        Next intCtrl
'    End If
'
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "PrintBarcodes", Err.Number, Err.Description

End Sub

'*****************************************************
'Description: Create and print a barcode directly to printer
'*****************************************************
Private Sub PrintNormalBarcode(ByVal strValue As String, ByVal lXPos As Long, ByVal lYPos As Long, lBarcodeWidth As Long, lBarcodeHeight As Long)
On Error GoTo ErrHandle
    
    TBarCode.Text = strValue
    Debug.Print strValue
    'dhdang sua day chuoi ma vach mau in sang phai
    'ngay 16-05-2011
'    If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Caption") = "BC26-AC" Then
'           ' TBarCode.BCDraw Printer.hDC, 7000, lYPos, lBarcodeWidth, lBarcodeHeight
'            TBarCode.BCDraw Printer.hDC, lXPos, lYPos, lBarcodeWidth, lBarcodeHeight
'    Else
    TBarCode.BCDraw Printer.hDC, lXPos, lYPos, lBarcodeWidth, lBarcodeHeight
'    End If
    
Exit Sub
ErrHandle:
    SaveErrorLog Me.Name, "PrintBarcode", Err.Number, Err.Description
End Sub

'*****************************************************
'Description: Print data and barcode image by sheet parameter
'           Step 1: Set default printer (local)
'           Step 2: Setup printer
'           Step 3: Print grid by sheet parameter
'           Step 4: Print barcode if it's first sheet
'*****************************************************
Public Sub PrintTax()
Dim strMsg As String
Dim intSheet As Integer, intIndex As Integer
Dim xmlNodeSheet As MSXML.IXMLDOMNode

On Error GoTo ErrHandle

    intCurrentBarcode = 0
    intCurrPage = 0
    blnHasPagePrinted = False
    
    'Init barcode license
    TBarCode.BarCode = eBC_PDF417
    TBarCode.Orientation = deg0
    Call TBarCode.LicenseMe("HCMTAX", eLicKindDeveloper, 10000, "B48994B2", eLicProd2D)
    
    'Print grid
    intSheet = 1
    fpsReport.sheet = intSheet 'Call Form_Load
    
    '*********************************
    '  added
    ' Date: 06/04/06
    ' Process printing session
    If TAX_Utilities_v2.DataChanged Then
        If intDataSession >= 999 Then
            intDataSession = 0
        Else
            intDataSession = intDataSession + 1
        End If
        If intPrintingSession >= 999 Then
            intPrintingSession = 0
        Else
            intPrintingSession = intPrintingSession + 1
        End If
    End If
    '*********************************
    'Setup printer and get number of barcode
    'SetupPrinter
    
    For intSheet = 1 To fpsReport.SheetCount - 1
        fpsReport.sheet = intSheet
        'fpsReport.PrintUseDataMax = False
        If intCurrPage <= lPageNumber Then
            Printer.Orientation = fpsReport.PrintOrientation
            
            intIndex = 0
            For Each xmlNodeSheet In TAX_Utilities_v2.NodeValidity.childNodes
                If UCase(GetAttribute(xmlNodeSheet, "ID")) = UCase(fpsReport.SheetName) Then
                    If GetAttribute(xmlNodeSheet, "Active") <> "0" Then
                        If Val(GetAttribute(TAX_Utilities_v2.NodeMenu, "ID")) = 68 Then
                            PrintSheetAC intSheet, intIndex
                        Else
                            PrintSheet intSheet, intIndex
                        End If
                    End If
                    Exit For
                End If
                intIndex = intIndex + 1
            Next
        End If
    Next intSheet
    
    If Not blnHasPagePrinted Then
        PrinterKillDoc
        Exit Sub
    End If
    PrinterEndDoc
    
    '*********************************
    '  added
    ' Date: 06/04/06
    If TAX_Utilities_v2.DataChanged Then
        If SaveSessionValueToFile(TAX_Utilities_v2.DataFolder & "Session.dat", , False) Then
            TAX_Utilities_v2.DataChanged = False
        End If
    End If
    '*********************************
Exit Sub
ErrHandle:
    SaveErrorLog "frmReportData", "PrintTax", Err.Number, Err.Description
'    If Err.Number = 396 Then
'        PrinterEndDoc
'    End If
    
End Sub

'*****************************************************
'Description: Print data and barcode image by sheet parameter
'           Step 1: Set default printer (local)
'           Step 2: Setup printer
'           Step 3: Print grid by sheet parameter
'           Step 4: Print barcode if it's first sheet
'*****************************************************
Private Sub PrintSheet(ByVal intSheet As Integer, ByVal intBarcodeIndex As Integer)
Dim intPageCount As Integer, intPage As Integer ', intPageNoByBarcode As Integer
Dim intSizeOfBarcode  As Integer, intCtrl As Integer
Dim lXOffset As Long, lYOffset As Long
Dim lPageBeginOfSheet As Long, lPageEndOfSheet As Long
Dim strValue As String, strPrefix As String 'mark string
Dim strTemp As String, intBarcodesOnPage As Integer 'Count of bacodes in one page
Dim arrStrValue As Variant
Dim clsConverter As New clsUnicodeTCVNConverter

On Error GoTo ErrHandle

    'Get string value which parser to barcode
    strValue = strDataBarcode(intBarcodeIndex)
    
    'Get number of page
    fpsReport.OwnerPrintPageCount Printer.hDC, 0, 0, Printer.Width, Printer.Height, intPageCount
    
    If LenB(strValue) / intPageCount <= 1500 Then
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount))
        intBarcodesOnPage = 1
    ElseIf LenB(strValue) / (intPageCount * 2) <= 1500 Then
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount * 2))
        intBarcodesOnPage = 2
    ElseIf LenB(strValue) / (intPageCount * 3) <= 1500 Then
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount * 3))
        intBarcodesOnPage = 3
    Else
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount * 4))
        intBarcodesOnPage = 4
    End If
    
    arrStrValue = CutStringByNumByte(strValue, IIf(intSizeOfBarcode Mod 2 = 0, intSizeOfBarcode + 2, intSizeOfBarcode + 1))
    
    'Get mark string
    strPrefix = GetPrefix(intSheet)
    
    'BeginPage  of sheet
    lPageBeginOfSheet = intCurrPage
    
    'Add mark
    For intCtrl = 1 To UBound(arrStrValue)
        intCurrentBarcode = intCurrentBarcode + 1
        'Add current barcode
    '****************************************
    '  added
    ' Date: 07/04/2006
        strTemp = strPrefix & format(intCurrentBarcode, "000")
        strTemp = strTemp & format(lBarcodeNumber, "000")
    '****************************************
        'Add prefix string to barcode string
        arrStrValue(intCtrl) = strTemp & CStr(arrStrValue(intCtrl)) & "#"
        arrStrValue(intCtrl) = clsConverter.Convert(CStr(arrStrValue(intCtrl)), UNICODE, TCVN)   'TAX_Utilities_v2.Compress(TAX_Utilities_v2.Convert(CStr(arrStrValue(intCtrl)), UNICODE, TCVN))
    Next intCtrl

    'Print first page
    If intCurrPage = 0 Then
        intCurrPage = intCurrPage + 1
    End If
    
    With fpsReport
        For intPage = 1 To intPageCount
            'Print page number
            If IsPrintedPage(intCurrPage) Then
                'Printer.CurrentX = Printer.ScaleWidth - Printer.ScaleX(20, vbMillimeters)   'Printer.ScaleX(180, vbMillimeters)
                If Printer.Orientation = 1 Then
                    Printer.CurrentX = Printer.ScaleX(7 * 1400, vbTwips)
                    Printer.CurrentY = Printer.ScaleY(Printer.Height - 0.55 * 1440, vbTwips) '11.15 * 1440, vbTwips)
                Else
                    Printer.CurrentX = Printer.ScaleX(10.6 * 1400, vbTwips)
                    Printer.CurrentY = Printer.ScaleY(Printer.Height - 0.55 * 1440, vbTwips)  '7.7 * 1440, vbTwips)
                End If

                'Printer.CurrentY = Printer.ScaleHeight - Printer.ScaleY(0.13, vbInches)     'Printer.ScaleY(4, vbMillimeters) ' Printer.ScaleY(280, vbMillimeters)
                Printer.ForeColor = vbBlack
                Printer.FontSize = 8
                Printer.Print "Trang " & intCurrPage & "/" & lPageNumber
            End If
            
            If intCurrPage < lPageNumber Then
                PrintPage intPage, arrStrValue, intBarcodesOnPage
                PrintNewPage
            Else
                PrintPage intPage, arrStrValue, intBarcodesOnPage ', True
                intCurrPage = intCurrPage + 1
            End If
        Next
    End With
    
'    lPageEndOfSheet = intCurrPage
'    If Not IsValidPrintedPage(lPageBeginOfSheet, lPageEndOfSheet) Then
'        PrinterKillDoc
'        Exit Sub
'    End If
'    PrinterEndDoc
    
    Exit Sub
ErrHandle:
    
    If Err.Number = 482 Then 'Printer error
        DisplayMessage "0057", msOKOnly, miCriticalError
        PrinterEndDoc
    End If
    
    SaveErrorLog "frmReportData", "PrintSheet", Err.Number, Err.Description
    
End Sub

Private Function GetPrefix(intSheet As Integer) As String
Dim strReturn As String, strTaxID As String, bFound As Boolean
Dim xmlNodeSheet As MSXML.IXMLDOMNode

' htphuong edit for KHBS  19/05/08
' Add app version to prefix

If strKHBS = "frmKHBS_BS" Or strKHBS = "TKBS" Then
    strReturn = "bs" & Replace$(APP_VERSION, ".", "") & GetAttribute(TAX_Utilities_v2.NodeMenu, "ID")
Else
    strReturn = "aa" & Replace$(APP_VERSION, ".", "") & GetAttribute(TAX_Utilities_v2.NodeMenu, "ID")
End If


'Get Tax ID
strTaxID = GetTaxIDString()

strReturn = strReturn & strTaxID

'Add period
If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
    ' to khai thang/quy GTGT, TNCN
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "04" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "94" Then
        If strQuy = "TK_THANG" Then
            strReturn = strReturn & TAX_Utilities_v2.month & TAX_Utilities_v2.Year
        Else
            strReturn = strReturn & "0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year
        End If
    Else
        strReturn = strReturn & TAX_Utilities_v2.month & TAX_Utilities_v2.Year
    End If
ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "13" Then
        If strQuy = "TK_THANG" Then
            strReturn = strReturn & TAX_Utilities_v2.month & TAX_Utilities_v2.Year
        Else
            strReturn = strReturn & "0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year
        End If
    Else
        strReturn = strReturn & "0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year
    End If
ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "1" Then
    strReturn = strReturn & "00" & TAX_Utilities_v2.Year
End If

'******************************
' added
'Date: 06/04/06
'Add printed session.
strReturn = strReturn & format(intDataSession, "000") & format(intPrintingSession, "000")
'******************************
'Restore sheet parameter
fpsReport.sheet = intSheet
GetPrefix = strReturn
End Function

Private Function GetTaxIDString() As String
Dim strReturn As String, strTemp As Variant
Dim intCtrl As Integer

'Move to last sheet
fpsReport.sheet = fpsReport.SheetCount
For intCtrl = 3 To 16
    If intCtrl <> 13 Then
        fpsReport.GetText intCtrl, 2, strTemp 'Tax_id_numbers
        If strTemp = "" Then strTemp = " "
        strReturn = strReturn & CStr(strTemp)
    End If
Next intCtrl

GetTaxIDString = strReturn
End Function

Private Function GetNumberOfBarcode() As Integer
Dim arrStrValue As Variant
Dim intSheet As Integer, intNumberOfBarcode As Integer, intIndex As Integer
Dim xmlNodeSheet As MSXML.IXMLDOMNode

For intSheet = 1 To fpsReport.SheetCount - 1
    fpsReport.sheet = intSheet
    intIndex = 0
    For Each xmlNodeSheet In TAX_Utilities_v2.NodeValidity.childNodes
        If UCase(GetAttribute(xmlNodeSheet, "ID")) = UCase(fpsReport.SheetName) Then
            If GetAttribute(xmlNodeSheet, "Active") <> "0" Then
                If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = 68 Then
                    arrStrValue = CutStringByNumByte(strDataBarcode(intIndex), 1000)
                    intNumberOfBarcode = intNumberOfBarcode + UBound(arrStrValue)
                Else
                    arrStrValue = CutStringByNumByte(strDataBarcode(intIndex), 1500)
                    intNumberOfBarcode = intNumberOfBarcode + UBound(arrStrValue)
                End If
            End If
            Exit For
        End If
        intIndex = intIndex + 1
    Next
Next intSheet

GetNumberOfBarcode = intNumberOfBarcode

End Function

Private Sub PrintNewPage()
    Dim lXOffset As Long, lYOffset As Long
    
    intCurrPage = intCurrPage + 1
    If Not IsPrintedPage(intCurrPage - 1) Then Exit Sub
    Printer.NewPage
End Sub

Private Sub PrintPage(intPage As Integer, arrStrValue As Variant, ByVal intBarcodesOnPage As Integer) ', Optional blnEndOfSheet As Boolean = False)
    '.OwnerPrintDraw Printer.hDC, -0.15 * 1440, 0.25 * 1440, Printer.Width - 0.15 * 1440, Printer.Height - 0.25 * 1440, intPage
    If IsPrintedPage(intCurrPage) Then
        fpsReport.OwnerPrintDraw Printer.hDC, 0, 0, Printer.Width, Printer.Height, intPage
        'Print barcode
        
        PrintBarcodes arrStrValue, intBarcodesOnPage, intBarcodesOnPage * (intPage - 1) + 1 ', blnEndOfSheet
    End If
    
End Sub

Public Sub SetPrintedPages(arrStrPages As Variant)
    arrStrPrintedPages = arrStrPages
End Sub

Private Function IsPrintedPage(intPage As Integer) As Boolean
Dim intCtrl As Integer, blnReturn As Boolean
Dim arrStrTemp() As String

blnReturn = False

If arrStrPrintedPages(0) = "All" Then
    IsPrintedPage = True
    blnHasPagePrinted = True
    Exit Function
End If

For intCtrl = 0 To UBound(arrStrPrintedPages())
    If InStr(1, arrStrPrintedPages(intCtrl), "-") = 0 Then
        If intPage = CInt(arrStrPrintedPages(intCtrl)) Then
            blnReturn = True
            blnHasPagePrinted = True
            Exit For
        End If
    ElseIf InStr(InStr(1, arrStrPrintedPages(intCtrl), "-") + 1, arrStrPrintedPages(intCtrl), "-") = 0 Then
        arrStrTemp = Split(arrStrPrintedPages(intCtrl), "-")
        If CInt(arrStrTemp(0)) <= intPage And CInt(arrStrTemp(1)) >= intPage Then
            blnReturn = True
            blnHasPagePrinted = True
            Exit For
        End If
    End If
Next intCtrl

IsPrintedPage = blnReturn
End Function

Private Function IsValidPrintedPage(lBeginPage As Long, lEndPage As Long) As Boolean
Dim lCtrl As Long
Dim arrStrTemp() As String

If arrStrPrintedPages(0) = "All" Then
    IsValidPrintedPage = True
    Exit Function
End If
For lCtrl = 0 To UBound(arrStrPrintedPages())
    If InStr(1, arrStrPrintedPages(lCtrl), "-") = 0 Then
        If CInt(arrStrPrintedPages(lCtrl)) <= lEndPage And CInt(arrStrPrintedPages(lCtrl)) > lBeginPage Then
            IsValidPrintedPage = True
            Exit Function
        End If
    Else
        arrStrTemp = Split(arrStrPrintedPages(lCtrl), "-")
        If Not (CInt(arrStrTemp(0)) > lEndPage Or CInt(arrStrTemp(1)) <= lBeginPage) Then
            IsValidPrintedPage = True
            Exit Function
        End If
    End If
Next lCtrl
IsValidPrintedPage = False
End Function

'Description: SetupPrinter procedure setup Bottom parameter to
'             the specified sheet.
'Output:
'    lNumberOfBarcode: Number of barcode images will be printed
Private Sub SetupPrinter()
    Dim lCtrl As Integer, intIndex As Integer, intIndex2 As Integer
    Dim intPageCount As Integer, lPageNoByBarcode As Long
    Dim lNumberOfBarcode As Long
    Dim arrStrValue As Variant
    Dim xmlNodeSheet As MSXML.IXMLDOMNode
    Dim arrLngRowPageBreak() As Long, lngTemp As Long
    Dim blnActiveSheet As Boolean
    
    Dim strIdKHBS As String
    
    lBarcodeNumber = 0
    lPageNumber = 0
    With fpsReport
        For lCtrl = 1 To .SheetCount - 1
            .sheet = lCtrl ' go to sheet
            'fpsReport.PrintOrientation = PrintOrientationDefault
            Printer.Orientation = fpsReport.PrintOrientation
            intIndex = 0
            blnActiveSheet = False
            For Each xmlNodeSheet In TAX_Utilities_v2.NodeValidity.childNodes
                If UCase(GetAttribute(xmlNodeSheet, "ID")) = UCase(.SheetName) Then
                    If GetAttribute(xmlNodeSheet, "Active") <> "0" Then
                        blnActiveSheet = True
                        'Get array of barcode string
                        'Neu BC 26 se dat ma vach khoang 1000
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = 68 Then
                            arrStrValue = CutStringByNumByte(strDataBarcode(intIndex), 1000)
                        ' end if
                        Else
                            arrStrValue = CutStringByNumByte(strDataBarcode(intIndex), 1500)
                        End If
                        'dhdang sua ngay 17/06/2010
                        'sua loi in bag ke to khai 05KK-TNCN bang ke in bi loi tran trang
                        'Get number of page calculated by number of barcode
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "17" And flgPrintBoSung = True And .sheet = 2 Then
                                lPageNoByBarcode = IIf(UBound(listInBoSung5A) Mod 40 = 0, UBound(listInBoSung5A) \ 40, UBound(listInBoSung5A) \ 40 + 1)
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "17" And flgPrintBoSung = True And .sheet = 3 Then
                                lPageNoByBarcode = IIf(UBound(listInBoSung5B) Mod 40 = 0, UBound(listInBoSung5B) \ 40, UBound(listInBoSung5B) \ 40 + 1)
                        Else
                                lPageNoByBarcode = IIf(UBound(arrStrValue) Mod 4 = 0, UBound(arrStrValue) \ 4, UBound(arrStrValue) \ 4 + 1)
                        End If
                        
                        
                        'Recalculate page and number of pages
                        If .PrintPageCount < lPageNoByBarcode Then
                            Do
                                .PrintMarginBottom = .PrintMarginBottom + 0.25 * 1440
                            Loop Until .PrintPageCount >= lPageNoByBarcode
                        End If
                    End If
                    Exit For
                End If
                intIndex = intIndex + 1
            Next

            'Get breaked page rows
            intIndex2 = 0
            Do
                lngTemp = .PrintNextPageBreakRow
                If lngTemp <> -1 Then
                    ReDim Preserve arrLngRowPageBreak(intIndex2)
                    arrLngRowPageBreak(intIndex2) = lngTemp
                    intIndex2 = intIndex2 + 1
                Else
                    Exit Do
                End If
            Loop Until False
            
            'Set breaked page rows
            If intIndex2 > 0 Then ' Exist breaked page row
                ' xu ly cho to khai 16TH
                Dim rowGroupI As Double
                If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "95" Then
                    .sheet = 1
                    .Row = 38
                    .Col = .ColLetterToNumber("B")
                    Do
                       .Row = .Row + 1
                       rowGroupI = .Row
                    Loop Until .Text = "aa"
                End If
                ' end
                
                For intIndex2 = 0 To UBound(arrLngRowPageBreak) - 1
                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "95" Then
                        ' xu ly rieng cho to khai 16TH
                        If arrLngRowPageBreak(intIndex2) < rowGroupI + 5 And arrLngRowPageBreak(intIndex2) > rowGroupI Then
                            .Row = rowGroupI + 1
                            .RowPageBreak = True
                        Else
                            .Row = arrLngRowPageBreak(intIndex2)
                            .RowPageBreak = True
                        End If
                    Else
                        .Row = arrLngRowPageBreak(intIndex2)
                        .RowPageBreak = True
                    End If
                Next intIndex2
                'Resize Last page
                'htphuong edit breakpage
                
                Dim rowBreak As Double
                                
                If lCtrl = 1 And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "11" Then
                    .Row = 67
                    .RowPageBreak = True
                ElseIf lCtrl = 1 And isDLT = False And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "12" Then
                    .Row = 72
                    .RowPageBreak = True
                ElseIf lCtrl = 1 And isDLT = True And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" Then
                    .Row = 59
                    .RowPageBreak = True
                ElseIf lCtrl = 1 And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "41" Then
                    .Row = 63
                    .RowPageBreak = True
                ElseIf lCtrl = 1 And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "75" Then
                    .Row = 42
                    .RowPageBreak = True
                ElseIf lCtrl = 1 And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "76" Then
                    .Row = 50
                    .RowPageBreak = True
                ElseIf lCtrl = 4 And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "03" Then
                    .Row = 51
                    .RowPageBreak = True
                ElseIf lCtrl = 1 And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "70" Then
                    .sheet = 1
                    .Row = 53
                    .Col = .ColLetterToNumber("C")
                    Do
                       .Row = .Row + 1
                       rowBreak = .Row
                    Loop Until .Text = "aa"
                  If .MaxRows - arrLngRowPageBreak(intIndex2) <= 8 Then
                        .Row = rowBreak
                        If .Row < arrLngRowPageBreak(intIndex2) Then
                            .Row = rowBreak
                            .RowPageBreak = True
                        Else
                            .Row = arrLngRowPageBreak(intIndex2) - 8
                            .RowPageBreak = True
                        End If
                    Else
                        .Row = arrLngRowPageBreak(intIndex2)
                        .RowPageBreak = True
                    End If
                Else
                    If .MaxRows - arrLngRowPageBreak(intIndex2) <= 5 Then
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "81" Then
                            If lCtrl = .SheetCount - 1 Then
                                .Row = GetLastDataRowKHBS(lCtrl)
                            Else
                                ' To khai 03/NTNN
                                .Row = 10 + GetLastDataRow(lCtrl)
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "71" Then
                            ' To khai 04/GTGT
                            If lCtrl = .SheetCount - 1 Then
                                .Row = GetLastDataRowKHBS(lCtrl)
                            Else
                                .Row = GetLastDataRow1(lCtrl)
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "90" Then
                            ' To khai 01/TBVMT
                            If lCtrl = 1 Then
                                .Row = GetLastDataRow1(lCtrl)
                            ElseIf lCtrl = .SheetCount - 1 Then
                                .Row = GetLastDataRowKHBS(lCtrl)
                            Else
                                .Row = GetLastDataRow(lCtrl)
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "05" Then
                             If lCtrl = 1 Then
                                .Row = GetLastDataRow2(lCtrl)
                             ElseIf lCtrl = .SheetCount - 1 Then
                                .Row = GetLastDataRowKHBS(lCtrl)
                            Else
                                .Row = GetLastDataRow(lCtrl)
                            End If
                        Else
                            If lCtrl = .SheetCount - 1 Then
                                'strIdKHBS_TT156 = "~01~02~04~71~72~11~12~73~70~81~06~05~86~90~94~96~98~99~92~"
                                strIdKHBS = GetAttribute(TAX_Utilities_v2.NodeMenu, "ID")
                                If InStr(1, strIdKHBS_TT156, "~" & Trim$(strIdKHBS) & "~", vbTextCompare) > 0 Or strIdKHBS = "01" Then
                                    .Row = GetLastDataRowKHBS(lCtrl)
                                Else
                                    .Row = GetLastDataRow(lCtrl)
                                End If
                            Else
                                If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" And lCtrl = 4 Then
                                    .Row = .MaxRows - 10
                                Else
                                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "92" And isDLT = True And lCtrl = 1 Then
                                        .Row = 78
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "97" And isDLT = True And lCtrl = 1 Then
                                        .Row = GetLastDataRowTD(lCtrl)
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "93" And isDLT = True And lCtrl = 1 Then
                                        .Row = 85
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "73" And isDLT = False And lCtrl = 1 Then
                                        .Row = arrLngRowPageBreak(intIndex2) - 1
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "77" And lCtrl = 1 Then
                                        .Row = GetLastDataRowTD(lCtrl) + 2
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "95" Then
                                        ' xu ly rieng cho to khai 16TH
                                        If arrLngRowPageBreak(intIndex2) < rowGroupI + 5 And arrLngRowPageBreak(intIndex2) > rowGroupI Then
                                            .Row = rowGroupI + 1
                                        Else
                                            .Row = GetLastDataRow(lCtrl)
                                        End If
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "55" Then
                                        .Row = GetLastDataRowDyn(lCtrl, 14)
                                    Else
                                        .Row = GetLastDataRow(lCtrl)
                                    End If
                                End If
                            End If
                        End If
                                                    
                        If .Row < arrLngRowPageBreak(intIndex2) Then
                            .RowPageBreak = True
                        Else
                            If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "95" Then
                                .RowPageBreak = True
                            Else
                                .Row = arrLngRowPageBreak(intIndex2) - 5
                                .RowPageBreak = True
                            End If
                        End If
                    Else
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" And .MaxRows - arrLngRowPageBreak(intIndex2) <= 6 And lCtrl = 2 Then
                            .Row = GetLastDataRow(lCtrl)
                            .RowPageBreak = True
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" And .MaxRows - arrLngRowPageBreak(intIndex2) <= 8 And lCtrl = 3 Then
                            .Row = GetLastDataRow(lCtrl)
                            .RowPageBreak = True
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "95" Then
                            ' xu ly rieng cho to khai 16TH
                            If arrLngRowPageBreak(intIndex2) <= rowGroupI + 5 And arrLngRowPageBreak(intIndex2) > rowGroupI Then
                                .Row = rowGroupI + 1
                                .RowPageBreak = True
                            Else
                                .Row = arrLngRowPageBreak(intIndex2)
                                .RowPageBreak = True
                            End If
                        ElseIf lCtrl = .SheetCount - 1 Then
                            .Row = GetLastDataRowKHBS(lCtrl)
                            .Row = .Row - 3
                            .RowPageBreak = True
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "55" Then
                             .Row = GetLastDataRowDyn(lCtrl, 14)
                             .RowPageBreak = True
                        Else
                            .Row = arrLngRowPageBreak(intIndex2)
                            .RowPageBreak = True
                        End If
                    End If
                End If
            End If
            
            'end
                                    
            If blnActiveSheet Then
                ' Tinh lai so ma vach tren mau AC
                If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "13" Then
                    intPageCount = .PrintPageCount
                
                    If LenB(strDataBarcode(intIndex)) / intPageCount <= 1000 Then
                        lBarcodeNumber = lBarcodeNumber + intPageCount
                    ElseIf LenB(strDataBarcode(intIndex)) / (intPageCount * 2) <= 1000 Then
                        lBarcodeNumber = lBarcodeNumber + intPageCount * 2
                    ElseIf LenB(strDataBarcode(intIndex)) / (intPageCount * 3) <= 1000 Then
                        lBarcodeNumber = lBarcodeNumber + intPageCount * 3
                    Else
                        lBarcodeNumber = lBarcodeNumber + intPageCount * 4
                    End If
        
                    'Get number of pages in sheet
                    lPageNumber = lPageNumber + intPageCount
                Else
                    ' To khai khac van giu nguyen ma kich thuoc vach
                    intPageCount = .PrintPageCount
                
                    If LenB(strDataBarcode(intIndex)) / intPageCount <= 1500 Then
                        lBarcodeNumber = lBarcodeNumber + intPageCount
                    ElseIf LenB(strDataBarcode(intIndex)) / (intPageCount * 2) <= 1500 Then
                        lBarcodeNumber = lBarcodeNumber + intPageCount * 2
                    ElseIf LenB(strDataBarcode(intIndex)) / (intPageCount * 3) <= 1500 Then
                        lBarcodeNumber = lBarcodeNumber + intPageCount * 3
                    Else
                        lBarcodeNumber = lBarcodeNumber + intPageCount * 4
                    End If
        
                    'Get number of pages in sheet
                    lPageNumber = lPageNumber + intPageCount
                End If
            End If
        Next lCtrl
    End With
End Sub

''' LoadInitFiles description
''' Set max len for string type cell
''' Set min value for numeric type cell
''' Set max value for numeric type cell
''' Call after load template
''' No parameter
Private Sub LoadInitFiles()
    On Error GoTo ErrorHandle
    Dim i As Long, lCol As Long, lRow As Long
    Dim xmlNodeListIni As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim fso As New FileSystemObject
    
    For i = 0 To fpsReport.SheetCount - 2
        ReDim Preserve xmlDocumentInit(i)
        Set xmlDocumentInit(i) = New MSXML.DOMDocument
        If fso.FileExists(GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(i), "ReportIni"))) Then
            xmlDocumentInit(i).Load GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(i), "ReportIni"))
            
            Set xmlNode = xmlDocumentInit(i).getElementsByTagName("Sections")(0)
            If GetAttribute(xmlNode, "PaperOrientation") <> "" Then
                fpsReport.sheet = i + 1
                fpsReport.PrintOrientation = IIf(GetAttribute(xmlNode, "PaperOrientation") = "Landscape", PrintOrientationLandscape, PrintOrientationPortrait)
            End If
            Set xmlNodeListIni = xmlDocumentInit(i).getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListIni
                fpsReport.sheet = i + 1
                ParserCellID fpsReport, GetAttribute(xmlNode, "CellID"), lCol, lRow
                fpsReport.Col = lCol
                fpsReport.Row = lRow
                If Val(GetAttribute(xmlNode, "MaxLen")) <> 0 Then
                    fpsReport.TypeMaxEditLen = Val(GetAttribute(xmlNode, "MaxLen"))
                End If
                If fpsReport.CellType = CellTypeNumber Then
                    fpsReport.TypeNumberMin = Val(GetAttribute(xmlNode, "MinValue"))
                    fpsReport.TypeNumberMax = Val(GetAttribute(xmlNode, "MaxValue"))
                End If
                fpsReport.CellTag = GetAttribute(xmlNode, "HelpContextID") & fpsReport.CellTag
            Next
        End If
    Next
    
    Set fso = Nothing
    Set xmlNode = Nothing
    Set xmlNodeListIni = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog Me.Name, "LoadInitFiles", Err.Number, Err.Description
End Sub


Private Function GetLastDataRow(ByVal lngSheet As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - 2)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRow = lRow
End Function

Private Function GetLastDataRowDyn(ByVal lngSheet As Long, byvanumRumRow As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - byvanumRumRow)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRowDyn = lRow
End Function

Private Function GetLastDataRow3(ByVal lngSheet As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - 4)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRow3 = lRow
End Function

Private Function GetLastDataRowTD(ByVal lngSheet As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - 15)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRowTD = lRow
End Function

Private Function GetLastDataRowKHBS(ByVal lngSheet As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - 7)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRowKHBS = lRow
End Function

' Ham lay gia tri dong de ngat trang
Private Function GetLastDataRow1(ByVal lngSheet As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - 10)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRow1 = lRow
End Function

' Ham lay gia tri dong de ngat trang
Private Function GetLastDataRow2(ByVal lngSheet As Long) As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long
    
    Set xmlNodeList = TAX_Utilities_v2.Data(lngSheet - 1).getElementsByTagName("Cell")
    Set xmlNode = xmlNodeList(xmlNodeList.length - 14)
    
    ParserCellID fpsReport, GetAttribute(xmlNode, "CellID2"), lCol, lRow
    
    GetLastDataRow2 = lRow
End Function

' Tinh toan lai chuoi ma vach cho cac mau BC AC
' test
Private Sub PrintSheetAC(ByVal intSheet As Integer, ByVal intBarcodeIndex As Integer)
Dim intPageCount As Integer, intPage As Integer ', intPageNoByBarcode As Integer
Dim intSizeOfBarcode  As Integer, intCtrl As Integer
Dim lXOffset As Long, lYOffset As Long
Dim lPageBeginOfSheet As Long, lPageEndOfSheet As Long
Dim strValue As String, strPrefix As String 'mark string
Dim strTemp As String, intBarcodesOnPage As Integer 'Count of bacodes in one page
Dim arrStrValue As Variant
Dim clsConverter As New clsUnicodeTCVNConverter

On Error GoTo ErrHandle

    'Get string value which parser to barcode
    strValue = strDataBarcode(intBarcodeIndex)
    
    'Get number of page
    fpsReport.OwnerPrintPageCount Printer.hDC, 0, 0, Printer.Width, Printer.Height, intPageCount
    
    If LenB(strValue) / intPageCount <= 1000 Then
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount))
        intBarcodesOnPage = 1
    ElseIf LenB(strValue) / (intPageCount * 2) <= 1000 Then
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount * 2))
        intBarcodesOnPage = 2
    ElseIf LenB(strValue) / (intPageCount * 3) <= 1000 Then
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount * 3))
        intBarcodesOnPage = 3
    Else
        intSizeOfBarcode = Int(LenB(strValue) / (intPageCount * 4))
        intBarcodesOnPage = 4
    End If
    
    arrStrValue = CutStringByNumByte(strValue, IIf(intSizeOfBarcode Mod 2 = 0, intSizeOfBarcode + 2, intSizeOfBarcode + 1))
    
    'Get mark string
    strPrefix = GetPrefix(intSheet)
    
    'BeginPage  of sheet
    lPageBeginOfSheet = intCurrPage
    
    'Add mark
    For intCtrl = 1 To UBound(arrStrValue)
        intCurrentBarcode = intCurrentBarcode + 1
        'Add current barcode
    '****************************************
    '  added
    ' Date: 07/04/2006
        strTemp = strPrefix & format(intCurrentBarcode, "000")
        strTemp = strTemp & format(lBarcodeNumber, "000")
    '****************************************
        'Add prefix string to barcode string
        arrStrValue(intCtrl) = strTemp & CStr(arrStrValue(intCtrl)) & "#"
        arrStrValue(intCtrl) = clsConverter.Convert(CStr(arrStrValue(intCtrl)), UNICODE, TCVN)   'TAX_Utilities_v2.Compress(TAX_Utilities_v2.Convert(CStr(arrStrValue(intCtrl)), UNICODE, TCVN))
    Next intCtrl

    'Print first page
    If intCurrPage = 0 Then
        intCurrPage = intCurrPage + 1
    End If
    
    With fpsReport
        For intPage = 1 To intPageCount
            'Print page number
            If IsPrintedPage(intCurrPage) Then
                'Printer.CurrentX = Printer.ScaleWidth - Printer.ScaleX(20, vbMillimeters)   'Printer.ScaleX(180, vbMillimeters)
                If Printer.Orientation = 1 Then
                    Printer.CurrentX = Printer.ScaleX(7 * 1400, vbTwips)
                    Printer.CurrentY = Printer.ScaleY(Printer.Height - 0.55 * 1440, vbTwips) '11.15 * 1440, vbTwips)
                Else
                    Printer.CurrentX = Printer.ScaleX(10.6 * 1400, vbTwips)
                    Printer.CurrentY = Printer.ScaleY(Printer.Height - 0.55 * 1440, vbTwips)  '7.7 * 1440, vbTwips)
                End If

                'Printer.CurrentY = Printer.ScaleHeight - Printer.ScaleY(0.13, vbInches)     'Printer.ScaleY(4, vbMillimeters) ' Printer.ScaleY(280, vbMillimeters)
                Printer.ForeColor = vbBlack
                Printer.FontSize = 8
                Printer.Print "Trang " & intCurrPage & "/" & lPageNumber
            End If
            
            If intCurrPage < lPageNumber Then
                PrintPage intPage, arrStrValue, intBarcodesOnPage
                PrintNewPage
            Else
                PrintPage intPage, arrStrValue, intBarcodesOnPage ', True
                intCurrPage = intCurrPage + 1
            End If
        Next
    End With
    
'    lPageEndOfSheet = intCurrPage
'    If Not IsValidPrintedPage(lPageBeginOfSheet, lPageEndOfSheet) Then
'        PrinterKillDoc
'        Exit Sub
'    End If
'    PrinterEndDoc
    
    Exit Sub
ErrHandle:
    
    If Err.Number = 482 Then 'Printer error
        DisplayMessage "0057", msOKOnly, miCriticalError
        PrinterEndDoc
    End If
    
    SaveErrorLog "frmReportData", "PrintSheetAC", Err.Number, Err.Description
    
End Sub
