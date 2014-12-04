VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02TAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 30
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 32
Const NGAY_NHAN_TO_KHAI_X = "R"
Const NGUOI_SU_DUNG_Y = 32
Const NGUOI_SU_DUNG_X = "AF"
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "AF"
Const NGAY_NOP_Y = 32
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 32
Const NGAY_QUET_X = "R"
Const TEN_BPQL_X = "AF"
Const TEN_BPQL_Y = 38
Const MA_BPQL_X = "AF"
Const MA_BPQL_Y = 34
Private rs As ADODB.Recordset
' thong tin din vi
Const MA_SO_THUE_X = "H"
Const MA_SO_THUE_Y = 10
Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 8
Const DIA_CHI_X = "H"
Const DIA_CHI_Y = 12
'Const QUAN_HUYEN_X = "F"
'Const QUAN_HUYEN_Y = 11
'Const TINH_TPHO_X = "T"
'Const TINH_TPHO_Y = 11
Const DIEN_THOAI_X = "H"
Const DIEN_THOAI_Y = 16
Const FAX_X = "R"
Const FAX_Y = 16
Const MAIL_X = "AF"
Const MAIL_Y = 16
'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "H"
Const MA_SO_THUE_DL_Y = 20
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 18
Const DIA_CHI_DL_X = "H"
Const DIA_CHI_DL_Y = 22
'Const QUAN_HUYEN_DL_X = "F"
'Const QUAN_HUYEN_DL_Y = 17
'Const TINH_TPHO_DL_X = "T"
'Const TINH_TPHO_DL_Y = 17
Const DIEN_THOAI_DL_X = "H"
Const DIEN_THOAI_DL_Y = 26
Const FAX_DL_X = "R"
Const FAX_DL_Y = 26
Const MAIL_DL_X = "AF"
Const MAIL_DL_Y = 26
Const SO_HDDL_DL_X = "H"
Const SO_HDDL_DL_Y = 28
Const NGAY_HDDL_DL_X = "R"
Const NGAY_HDDL_DL_Y = 28
'end
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Private LAN_QUET As Variant
Const STT_X = "R"
Const STT_Y = 23
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long

'Longvh
Public strMST As String
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public strMail As String
Public strQHuyen As String
Public strTPho As String
'vttoan them thong tin dai ly thue
Public strMST_DLT As String
Public strTen_DLT As String
Public strDchi_DLT As String
Public strQHuyen_DLT As String
Public strTTPho_DLT As String
Public strDthoai_DLT As String
Public strFax_DLT As String
Public strMail_DLT As String
Public strSoHD_DLT As String
Public strNgayHD_DLT As String
'end

Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public TTHTK As String
Public strTenBpql As String
Private tempKyLb As Variant
Private tempNgayNop As Variant
Public hdrId As String
Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String
Public strSoTTTKhai As String
Public isTKTonTai As Boolean
Public checkDLTTonTai As Boolean
Private MaMucKHBS As String
Private strTuKy As Variant
Private strDenKy As Variant


Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
   Dim namQT As Variant
   Dim madtnt As Variant
   Dim MaMuc As Variant
   Dim maTM As String
   Dim ngnop As Variant
   Dim LoaiTK As Variant
   Dim NganhTN As Variant
   Dim SanLuong As Variant
   Dim DVTinh As Variant
   Dim DONGIA As Variant
   Dim GTTT As Variant
   Dim THUESUAT As Variant
   Dim MucThue As Variant
   Dim THUEPS As Variant
   Dim MG As Variant
   Dim THUETN As Variant
    Dim STT As Variant
   Dim STT2 As Variant
   Dim STTIN As Variant
   Dim mact As Variant
   Dim GTTTCL As Variant
   Dim ThueTNCL As Variant
   
   Dim ThueKK As Variant
   Dim THUECLQT As Variant
   'TTHTK

   Dim DANHAN As Variant
   
'   Dim countSTT As Integer
'
'   Dim Section As Integer
   
    sSQLCol = "NAMQT,MADTNT,MAMUC,MATM,NGNOP,LOAITK,NGANHTN,SANLUONG,DVTINH,DONGIA,GTTT,THUESUAT,MUCTHUE,THUEPS" _
             & ",MG,THUETN,TTHTK,STT,STT2,STTIN,MACT,GTTTCL,THUETNCL,THUEKK,THUECLQT,DANHAN,LAN_QUET"

   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If


   With fps
        .Sheet = 1
        namQT = TAX_Utilities_Svr_New.Year
        
   'ma doi tuong nop thue
        .GetText .ColLetterToNumber("H"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
   'Dim MaMuc As Variant
   'Dim maTM As Variant
  'ngay nop
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
   LoaiTK = "'02/TAIN'"
   
   'Dim NganhTN As Variant
   'Dim SanLuong As Variant
   'Dim DVTinh As Variant
   'Dim DONGIA As Variant
   'Dim GTTT As Variant
   'Dim THUESUAT As Variant
   'Dim MucThue As Variant
   'Dim THUEPS As Variant
   'Dim MG As Variant
   'Dim THUETN As Variant
    'Dim STT As Variant
   'Dim STT2 As Variant
   'Dim STTIN As Variant
   mact = "'1000'"
   'Dim GTTTCL As Variant
   'Dim ThueTNCL As Variant
   
   'Dim ThueKK As Variant
   'Dim THUECLQT As Variant
   'TTHTK

DANHAN = "''"
        
         
        .Col = .ColLetterToNumber("B")
        .Row = 55
        Dim lastRow As Variant
        Dim maKhoangSan As Variant
        
        'TODO:
        'MaTM, MaMuc lay theo danh muc, chua chot duoc
        'Tam thoi fix gia tri de chay
        'Sau nay nho sua lai
        'maTM = "'1553'"
        
    Do

        ' I Khoang san do co so tu khai thac: mact=1000
            ' Dinh nghia trong bang CTBVMT.dbf thuoc folder DB_HT
            mact = "'1000'"

            .GetText .ColLetterToNumber("AN"), .Row, maKhoangSan
            ' Tu maKhoangSan lay ra maTM
            DataDM maKhoangSan, , , , , , maTM

            'maTM = "1553"
             'khong duoc phep de trong
            MaMuc = GetMaMuc(maTM)
            If Trim(MaMuc) = vbNullString Then
                MaMuc = "''"
            Else
                MaMuc = "'" & MaMuc & "'"
            End If

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
            'MaMuc = "'1550'"
           
                    
            
            
            ' nghanh tai nguyen
            .GetText .ColLetterToNumber("AN"), .Row, NganhTN

            If Trim(NganhTN) = vbNullString Then

                NganhTN = "''"
            Else
                NganhTN = "'" & NganhTN & "'"
            End If
            
                       
            
            .GetText .ColLetterToNumber("O"), .Row, SanLuong

            If Trim(SanLuong) = vbNullString Then

                SanLuong = "0"
            End If

            ' don vi tinh
            .GetText .ColLetterToNumber("L"), .Row, DVTinh

            If Trim(DVTinh) = vbNullString Then

                DVTinh = "''"
            Else
                DVTinh = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVTinh), UNICODE, TCVN) & "'"
            End If


            ' don vi tinh
            .GetText .ColLetterToNumber("R"), .Row, DONGIA

            If Trim(DONGIA) = vbNullString Then
                DONGIA = "0"
            End If
            ' muc phi

            GTTT = "0"

            .GetText .ColLetterToNumber("W"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If

            .GetText .ColLetterToNumber("Z"), .Row, MucThue

            If Trim(MucThue) = vbNullString Then

                MucThue = "0"
            End If

            'thue phai nop
            .GetText .ColLetterToNumber("AE"), .Row, THUEPS

            If Trim(THUEPS) = vbNullString Then

                THUEPS = "0"
            End If
            
            
            
            'thue phai nop
            .GetText .ColLetterToNumber("AH"), .Row, MG

            If Trim(MG) = vbNullString Then

                MG = "0"
            End If
            
            ' thue ke khai
            .GetText .ColLetterToNumber("AL"), .Row, THUETN

            If Trim(THUETN) = vbNullString Then

                THUETN = "0"
            End If
            
            
            GTTTCL = "0"
            ThueTNCL = THUETN
            
            .GetText .ColLetterToNumber("AP"), .Row, ThueKK

            If Trim(ThueKK) = vbNullString Then

                ThueKK = "0"
            End If
            
            .GetText .ColLetterToNumber("AQ"), .Row, THUECLQT

            If Trim(THUECLQT) = vbNullString Then

                THUECLQT = "0"
            End If
            
            
            'so thu tu:
            If .Row = 55 Then
                STT2 = "0"
                STTIN = "'I'"
                MaMuc = "''"
                maTM = "''"
            Else
                .GetText .ColLetterToNumber("B"), .Row, STT2

                If Trim(STT) = vbNullString Then
                    STT2 = "0"
                End If

                STTIN = "'" & STT2 & "'"
            End If

            STT = "1"
            
            sSQLVal = namQT & "," & madtnt & "," & MaMuc & "," & maTM & "," & ngnop & "," & LoaiTK & "," & NganhTN
        sSQLVal = sSQLVal & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue
        sSQLVal = sSQLVal & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & STT & "," & STT2 & "," & STTIN
        sSQLVal = sSQLVal & "," & mact & "," & GTTTCL & "," & ThueTNCL & "," & ThueKK & "," & THUECLQT & "," & DANHAN & "," & LAN_QUET


        sSQL = "INSERT INTO TMP_QA" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        
             
            If .Row <> 56 Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            lastRow = .Row
    Loop Until .Text = "aa"
             
             
             
    .Col = .ColLetterToNumber("B")
        .Row = lastRow + 1
        maTM = ""
        'TODO:
        

        Do
            
            ' II Khoang san do co so tu khai thac: mact=2000
            ' Dinh nghia trong bang CTBVMT.dbf thuoc folder DB_HT
            mact = "'2000'"

            
            

            .GetText .ColLetterToNumber("AN"), .Row, maKhoangSan
            ' Tu maKhoangSan lay ra maTM
            DataDM maKhoangSan, , , , , , maTM

'
            MaMuc = GetMaMuc(maTM)
            If Trim(MaMuc) = vbNullString Then
                MaMuc = "''"
            Else
                MaMuc = "'" & MaMuc & "'"
            End If
            'maTM = "1553"

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
            'MaMuc = "'1550'"

                    
            
            
            ' nghanh tai nguyen
            .GetText .ColLetterToNumber("AN"), .Row, NganhTN

            If Trim(NganhTN) = vbNullString Then

                NganhTN = "''"
            Else
                NganhTN = "'" & NganhTN & "'"
            End If
            
                       
            
            .GetText .ColLetterToNumber("O"), .Row, SanLuong

            If Trim(SanLuong) = vbNullString Then

                SanLuong = "0"
            End If

            ' don vi tinh
            .GetText .ColLetterToNumber("L"), .Row, DVTinh

            If Trim(DVTinh) = vbNullString Then

                DVTinh = "''"
            Else

                DVTinh = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVTinh), UNICODE, TCVN) & "'"
            End If


            ' don vi tinh
            .GetText .ColLetterToNumber("R"), .Row, DONGIA

            If Trim(DONGIA) = vbNullString Then
                DONGIA = "0"
            End If
            ' muc phi

            GTTT = "0"

            .GetText .ColLetterToNumber("W"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If

            .GetText .ColLetterToNumber("Z"), .Row, MucThue

            If Trim(MucThue) = vbNullString Then

                MucThue = "0"
            End If

            'thue phai nop
            .GetText .ColLetterToNumber("AE"), .Row, THUEPS

            If Trim(THUEPS) = vbNullString Then

                THUEPS = "0"
            End If
            
            
            
            'thue phai nop
            .GetText .ColLetterToNumber("AH"), .Row, MG

            If Trim(MG) = vbNullString Then

                MG = "0"
            End If
            
            ' thue ke khai
            .GetText .ColLetterToNumber("AL"), .Row, THUETN

            If Trim(THUETN) = vbNullString Then

                THUETN = "0"
            End If
            
            
            GTTTCL = "0"
            ThueTNCL = THUETN
            
            .GetText .ColLetterToNumber("AP"), .Row, ThueKK

            If Trim(ThueKK) = vbNullString Then

                ThueKK = "0"
            End If
            
            .GetText .ColLetterToNumber("AQ"), .Row, THUECLQT

            If Trim(THUECLQT) = vbNullString Then

                THUECLQT = "0"
            End If
            
            'so thu tu:
            If .Row = (lastRow + 1) Then
                STT2 = "0"
                STTIN = "'II'"
                MaMuc = "''"
                maTM = "''"
            Else
                .GetText .ColLetterToNumber("B"), .Row, STT2

                If Trim(STT) = vbNullString Then
                    STT2 = "0"
                End If

                STTIN = "'" & STT2 & "'"
            End If

            STT = "2"
           
            sSQLVal = namQT & "," & madtnt & "," & MaMuc & "," & maTM & "," & ngnop & "," & LoaiTK & "," & NganhTN
        sSQLVal = sSQLVal & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue
        sSQLVal = sSQLVal & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & STT & "," & STT2 & "," & STTIN
        sSQLVal = sSQLVal & "," & mact & "," & GTTTCL & "," & ThueTNCL & "," & ThueKK & "," & THUECLQT & "," & DANHAN & "," & LAN_QUET



        sSQL = "INSERT INTO TMP_QA" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                   
            
            If .Row <> (lastRow + 2) Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
    

        
'       ' sua MAMUC & MATM tu nam 2009
'        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
'            maTM = "'01'"
'            MaMuc = "'012'"
'        Else
'            MaMuc = "'1550'"
'        End If
 
       
        End With
        
    clsDAO.Disconnect

   InsertDTL = vbNullString
   
End Function
Private Function GetMaMuc(ByVal MaTieuMuc As Variant) As String
    Dim strDataFileName As String
    Dim pSQL            As String
    Dim strName         As String
    
    Dim rcord           As ADODB.Recordset
    Dim fso             As New FileSystemObject
    
    Dim clsDAOs As New TAX_Utilities_Svr_New.clsADO
    
    strDataFileName = spathVat & "\DB_HT\DMMT2.DBF"

    If fso.FileExists(strDataFileName) Then
        If clsDAOs.Connected = False Then
            ' Lay mamuc theo matm
            clsDAOs.CreateConnectionString spathVat & "\DB_HT\"
            clsDAOs.Connect
        End If
    
        pSQL = "SELECT MAMUC FROM DMMT2 WHERE MATM = '" & MaTieuMuc & "'"
       
        Set rcord = clsDAOs.Execute(pSQL)

        If Not rcord Is Nothing Then
            strName = rcord.Fields("MAMUC")
        End If
    End If
    
    GetMaMuc = strName
    
    clsDAOs.Disconnect
End Function
Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTTN" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QA" & Right(sKyKeKhai, 4) & ".DBF"
    Dim strFileNameBS As String
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    Dim strFileNameBSHDR As String
    strFileNameBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_ND" & sKyKeKhai & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QTTN.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If
    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QA.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    'file bo xung
    If fso.FileExists(strFileNameBS) = False Then
        Dim strTepmauBS As String
        strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmauBS, strFileNameBS, False
    End If
    
    If fso.FileExists(strFileNameBSHDR) = False Then
        Dim strTepmauBSHDR As String
        strTepmauBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_ND.DBF"
        fso.CopyFile strTepmauBSHDR, strFileNameBSHDR, False
    End If
    
    isExistFile = True
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTTN" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND NAMQT =" & TAX_Utilities_Svr_New.Year & _
        " AND LOAITK='02/TAIN' AND TUKY ='" & strTuKy & "' AND DENKY ='" & strDenKy & "'"
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        dNgayNopDB = rs.Fields("NGNOP")
        If Trim(dNgayNopDB) = vbNullString Then
            dNgayNopDB = "CTOD('')"
        Else
            dNgayNopDB = "CTOD('" & Format(dNgayNopDB, "mm/dd/yyyy") & "')"
        End If
        TKTT = True
   Else
        TKTT = False
   End If
   LAN_QUET = TinhSoLanQuet + 1
   clsDAO.Disconnect

End Function

Public Function TKTTNGNOP() As Boolean
   Dim ngnop As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = ""
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
   End With
   
   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTTN" & Right(sKyKeKhai, 4) & _
        " WHERE  MADTNT = " & "'" & strMST & "'" & _
        " AND  LOAITK='02/TAIN'"
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        For i = 1 To rs.RecordCount
            If Not rs.EOF Then
                NGNOPDB = rs.Fields("NGNOP")
                If Trim(NGNOPDB) = vbNullString Then
                    NGNOPDB = ""
                Else
                    NGNOPDB = "CTOD('" & Format(NGNOPDB, "mm/dd/yyyy") & "')"
                End If
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(ngnop) Then
                    clsDAO.Disconnect
                    TKTTNGNOP = False
                    Exit Function
                Else
                    TKTTNGNOP = True
                End If
            End If
            rs.MoveNext
        Next
   Else
        clsDAO.Disconnect
        TKTTNGNOP = True
   End If

End Function

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "UPDATE TMP_QTTN" & Right(sKyKeKhai, 4) & _
        " SET THUETKY2=1" & _
        " WHERE AND MADTNT = " & "'" & strMST & "'" & _
        " AND LOAITK='02/TAIN'"
    
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   sSQL = "DELETE FROM TMP_QTTN" & Right(sKyKeKhai, 4) & _
        " WHERE  MADTNT = " & "'" & strMST & "'" & _
        " AND  LOAITK='02/TAIN'"
    
   clsDAO.ExecuteDLL sSQL
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 9, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
   End With
   sSQL = "DELETE FROM TMP_TA" & sKyKeKhai & _
        " WHERE MATHUE = '04' AND  MAPP = '1' AND MADTNT = " & "'" & strMST & "' AND KYLBO = " & KYLBO & _
        " AND LOAITK='02/TAIN' AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

Public Function InsertDTL_KHBS() As String
Dim madtnt As Variant
    Dim ngnop  As Variant
    Dim KYLBO  As Variant

    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("H"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

    End With

    InsertDTL_KHBS = InsertDTL_KHBS156(fps, madtnt, "02/TAIN", ngnop, KYLBO, spathVat, TTHTK, LAN_QUET, False, sKyKeKhai, MaMucKHBS)
    
    InsertHDR_KHBS
    
    
    InsertDTL_KHBS = vbNullString
   
End Function
Public Function InsertHDR_KHBS() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
    
    Dim madtnt As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
    Dim MATHUE As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
    Dim MAPP As Variant
    Dim ngnop As Variant
    'Dim TTHTK As Variant
    Dim KYLBO As Variant
    Dim songaycn As Variant
    Dim sotiencn As Variant
    Dim tienhoan As Variant
    Dim lenhhoan As Variant
    Dim ngayhoan As Variant
    Dim HANNOP As Variant
    Dim cqt As Variant
    Dim songay As Variant
    Dim SOTIEN As Variant
    Dim lydokhac As Variant
    Dim DANHAN As Variant
    
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("H"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
        MaMuc = "'1550'"
        maTM = "''"
        matkhai = "'02/TAIN'"
        

        
        MATHUE = "'04'"
        MAPP = "'1'"
       

        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        
        'Lay sheet KHBS tren to khai
        .Sheet = .SheetCount - 1
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 28, songaycn
        If Trim(songaycn) = vbNullString Then
            songaycn = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 27, sotiencn
        If Trim(sotiencn) = vbNullString Then
            sotiencn = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 21, tienhoan
        If Trim(tienhoan) = vbNullString Then
            tienhoan = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 19, lenhhoan
        If Trim(lenhhoan) = vbNullString Then
            lenhhoan = "''"
        Else
            lenhhoan = "'" & TAX_Utilities_Svr_New.Convert(CStr(lenhhoan), UNICODE, TCVN) & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), .MaxRows - 17, ngayhoan

        If Trim(ngayhoan) = vbNullString Then
            ngayhoan = "CTOD('')"
        Else
            ngayhoan = ToDate(Trim(ngayhoan), DDMMYYYY)
            ngayhoan = "CTOD('" & Format(ngayhoan, "mm/dd/yyyy") & "')"
        End If
        
        HANNOP = "CTOD('')"
        
        .GetText .ColLetterToNumber("BI"), .MaxRows - 13, cqt
        If Trim(cqt) = vbNullString Then
            cqt = "''"
        Else
            cqt = "'" & cqt & "'"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 11, songay
        If Trim(songay) = vbNullString Then
            songay = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 9, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        
        .GetText .ColLetterToNumber("BD"), .MaxRows - 5, lydokhac
        If Trim(lydokhac) = vbNullString Then
            lydokhac = "''"
        Else
            lydokhac = "'" & TAX_Utilities_Svr_New.Convert(CStr(lydokhac), UNICODE, TCVN) & "'"
        End If
    
        DANHAN = "''"
        
        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If
    
        sSQLCol = "madtnt,kykkhai,matkhai,MATHUE,MaMuc,maTM,MAPP,ngnop,TTHTK,KYLBO,songaycn,sotiencn,tienhoan," & _
            "lenhhoan,ngayhoan,HANNOP,cqt,songay,SOTIEN,lydokhac,DANHAN,LAN_QUET"
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & songaycn & "," & sotiencn & "," & tienhoan & "," & _
            lenhhoan & "," & ngayhoan & "," & HANNOP & "," & cqt & "," & songay & "," & SOTIEN & "," & lydokhac & "," & DANHAN & "," & LAN_QUET
        
        sSQL = "INSERT INTO TMP_ND" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
    End With

    InsertHDR_KHBS = vbNullString
    
    clsDAO.Disconnect
    
End Function
Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    
    Dim TuKy As Variant
    Dim DenKy As Variant
    Dim namQT      As Variant
    Dim madtnt     As Variant
    Dim MaMuc      As Variant
    Dim maTM       As String
    Dim ngnop      As Variant
    Dim NgNhap     As Variant
    Dim LoaiTien   As Variant
    Dim KyDC       As Variant
    Dim ThueTKy    As Variant
    Dim DaGhi      As Variant
    Dim LoaiTK     As Variant
    Dim MaVach     As Variant
    Dim DATHAY     As Variant
    Dim KYLBO      As Variant
    Dim DoanhSo    As Variant
    Dim GTTinhThue As Variant
    Dim STT As Variant
    Dim STT2 As Variant
    Dim STTIN As Variant
    Dim mact As Variant
    
    Dim madlt      As Variant
    Dim LanBS      As Variant
    Dim THUECL     As Variant
    Dim DSoCL      As Variant
    Dim GTTTCL     As Variant
    Dim DANHAN     As Variant
'    Dim MAPP       As Variant
'    Dim TEMP       As Variant
'    Dim TEMP2      As Variant
    Dim bln        As Boolean
    'Dim Section    As Integer
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant
    Dim ThueCQ As Variant
    
    'bien dung xu ly Ma Tieu Muc
    Dim arrMTM()     As String, arrMTMSave() As String
    Dim arrThueTKy() As String, arrThueTKySave() As Double
    Dim arrThueCQ() As String, arrThueCQSave() As Double
    Dim i, j As Integer
    Dim vTemp        As String
    Dim countThueTKy As Double
    Dim countThueCQ As Double
    Dim strMTM       As String
    Dim strThueTKy   As String
    Dim strThueCQ    As String
    Dim ROWNUM       As Variant
    Dim vTemp2       As String

    
'    'test
'    Dim arrMTM()   As String, arrMTMSave() As String
'    Dim arrVitri() As Integer
'    Dim idx, idx1, i, countMTM As Integer
'    Dim vTemp As String, vVitriTrung As String
'    'end test
    
    sSQLCol = "NAMQT,TUKY,DENKY,MADTNT,MAMUC,MATM,NGNOP,NGNHAP,LOAITIEN,KYDC,THUETKY,DAGHI,LOAITK,TTHTK,MAVACH,DATHAY,HANNOP"
    sSQLCol = sSQLCol & ",KYLBO,DOANHSO,GTTINHTHUE,STT,STT2,STTIN,MACT,MADLT,LANBS,THUECL,DSOCL,GTTTCL,DANHAN,LAN_QUET,TENNV,CHUNGCHI,THUECQ"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
       
    With fps
        .Sheet = 1
        
        'Section = 1
        .GetText .ColLetterToNumber("G"), .MaxRows - 3, TENNV
        If Trim(TENNV) <> vbNullString Then
            TENNV = TAX_Utilities_Svr_New.Convert(CStr(TENNV), UNICODE, TCVN)
        End If
        
        .GetText .ColLetterToNumber("G"), .MaxRows - 1, CHUNGCHI
        If Trim(CHUNGCHI) <> vbNullString Then
            CHUNGCHI = TAX_Utilities_Svr_New.Convert(CStr(CHUNGCHI), UNICODE, TCVN)
        End If
        
        'Doan nay tinh ThueTKy theo maTM
        .Col = .ColLetterToNumber("B")
        .Row = 57

        strMTM = "~"
        'I: Tai nguyen khai thac
        Dim maKhoangSan As Variant
        
        Do
            ' Lay ma khoang san
            .GetText .ColLetterToNumber("AN"), .Row, maKhoangSan
            ' Tu maKhoangSan lay ra maTM
            DataDM maKhoangSan, , , , , , maTM
            'maTM = "1553"
            vTemp = "~" & maTM & "~"
            vTemp2 = vTemp2 & maTM & "~"
            ' Kiem tra neu co maTM moi thi bo sung vao --> strMTM se luu cac maTM duy nhat
            If InStr(1, strMTM, vTemp, vbTextCompare) = 0 Then
                strMTM = strMTM & maTM & "~"
            End If
            
            'Ghep cac gia tri = Tong so thue phai nop
            .Col = .ColLetterToNumber("AL")
            strThueTKy = strThueTKy & .Text & "~"
            
            'Ghep cac gia tri = Tong so thue ke khai
            .Col = .ColLetterToNumber("AP")
            strThueCQ = strThueCQ & .Text & "~"
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            ROWNUM = .Row
        Loop Until .Text = "aa"
        
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 3

        'II Khoang san do co so thu mua nop thay  nguoi khac
        Do
            ' Lay ma khoang san
            .GetText .ColLetterToNumber("AN"), .Row, maKhoangSan
            ' Tu maKhoangSan lay ra maTM
        DataDM maKhoangSan, , , , , , maTM
            'maTM = "1553"
            
            vTemp = "~" & maTM & "~"
            vTemp2 = vTemp2 & maTM & "~"
            ' Kiem tra neu co maTM moi thi bo sung vao --> strMTM se luu cac maTM duy nhat
            If InStr(1, strMTM, vTemp, vbTextCompare) = 0 Then
                strMTM = strMTM & maTM & "~"
            End If
            
            'Ghep cac gia tri = Tong so thue phai nop
            .Col = .ColLetterToNumber("AL")
            strThueTKy = strThueTKy & .Text & "~"
            
            'Ghep cac gia tri = Tong so thue ke khai
            .Col = .ColLetterToNumber("AP")
            strThueCQ = strThueCQ & .Text & "~"
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        'Mang nay luu gia tri cac maTM duy nhat
        arrMTMSave = Split(strMTM, "~")
        
        'Mang nay luu gia tri tuong ung
        arrThueTKy = Split(strThueTKy, "~")
        arrThueCQ = Split(strThueCQ, "~")
        arrMTM = Split(vTemp2, "~")
        
        'Fix lai length cua mang luu cac tong THUETKY theo tung maTM
        ReDim arrThueTKySave(UBound(arrMTMSave))
        'Fix lai length cua mang luu cac tong THUECQ theo tung maTM
        ReDim arrThueCQSave(UBound(arrMTMSave))
        
        
        'Tinh tong ThueTKy, ThueCQ theo tung maTM
        For i = 1 To UBound(arrMTMSave) - 1
            countThueTKy = 0
            countThueCQ = 0
            For j = 0 To UBound(arrMTM) - 1
                If arrMTMSave(i) = arrMTM(j) Then
                    countThueTKy = countThueTKy + Val(Replace(arrThueTKy(j), ".", ""))
                    countThueCQ = countThueCQ + Val(Replace(arrThueCQ(j), ".", ""))
                End If
            Next
            arrThueTKySave(i) = countThueTKy
            arrThueCQSave(i) = countThueCQ
        Next
        

        'section1
        .GetText .ColLetterToNumber("H"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
       
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
    
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("R"), 32, NgNhap

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("H"), 20, madlt

        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If
                
        .GetText .ColLetterToNumber("AC"), 3, LanBS

        If Trim(LanBS) = vbNullString Then
            LanBS = 0
        End If
                
        namQT = TAX_Utilities_Svr_New.Year


        'KyTTTu
        .GetText .ColLetterToNumber("M"), 49, TuKy
        'KyTTTu
        .GetText .ColLetterToNumber("N"), 49, DenKy
        

        LoaiTien = "''"
        DaGhi = ".T."
        LoaiTK = "'02/TAIN'"
        DANHAN = "''"

'        If Trim(HANNOP) = vbNullString Then
'            HANNOP = "CTOD('')"
'        Else
'            HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
'            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
'        End If

        HANNOP = "CTOD('')"

        KyDC = "''"
        DATHAY = ".F."
        MaVach = ".T."
        LoaiTien = "''"

        DoanhSo = 0
        GTTinhThue = 0
        'end section1

        '        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"

        .GetText .ColLetterToNumber("R"), 32, NgNhap

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
       
'        For i = 0 To UBound(arrMTMSave) - 1
'            Section = 1
'            ThueTKy = 0
'            GTTinhThue = 0
'            vTemp = arrMTMSave(i)
'            maTM = "'" & vTemp & "'"
'            .Col = .ColLetterToNumber("B")
'            .Row = 57
'
'            Do
'                .Col = .ColLetterToNumber("AO")
'
'                If vTemp = Trim(.Text) Then
'                    .Col = .ColLetterToNumber("AL")
'                    TEMP = .Value
'
'                    If Trim(TEMP) = vbNullString Then
'                        TEMP = "0"
'                    End If
'
'                    ThueTKy = ThueTKy + CDbl(TEMP)
'
'                    .Col = .ColLetterToNumber("O")
'                    TEMP = .Value
'
'                    If Trim(TEMP) = vbNullString Then
'                        TEMP = "0"
'                    End If
'
'                    .Col = .ColLetterToNumber("R")
'                    TEMP2 = .Value
'
'                    If Trim(TEMP2) = vbNullString Then
'                        TEMP2 = "0"
'                    End If
'
'                    GTTinhThue = GTTinhThue + CDbl(TEMP) * CDbl(TEMP2)
'
'                End If
'
'                .Col = .ColLetterToNumber("B")
'                .Row = .Row + 1
'
'                If .Text = "aa" And Section = 1 Then
'                    .Row = .Row + 3
'                    Section = 2
'                End If
'
'            Loop Until .Text = "aa"
        
            DoanhSo = GTTinhThue
           
            DaGhi = ".F."
            MaVach = ".T."

            'THUECL = ThueTKy
            DSoCL = "0"
            GTTTCL = "0"
            DANHAN = "''"
            STT = "0"
            STT2 = "0"
            STTIN = "''"
            mact = "''"
  
            'sSQLCol = "NAMQT,TUKY,DENKY,MADTNT,MAMUC,MATM,NGNOP,NGNHAP,LOAITIEN,KYDC,THUETKY,DAGHI,LOAITK,TTHTK,MAVACH,DATHAY,HANNOP"
            'sSQLCol = sSQLCol & "KYLBO,DOANHSO,GTTINHTHUE,STT,STT2,STTIN,MACT,MADLT,LANBS,THUECL,DSOCL,GTTTCL,DA_NHAN,LAN_QUET"
         'Tinh thuetky group by theo matm
        For i = 1 To UBound(arrMTMSave) - 1
            vTemp = arrMTMSave(i)
            maTM = "'" & vTemp & "'"
            ThueTKy = arrThueTKySave(i)
            ThueCQ = arrThueCQSave(i)
            'TODO
            'MaMuc = "'1550'"
            MaMuc = GetMaMuc(vTemp)
            If Trim(MaMuc) = vbNullString Then
                MaMuc = "''"
            Else
                MaMuc = "'" & MaMuc & "'"
            End If
            
            'MaMucKHBS = MaMuc
            
            If Trim(ThueTKy) <> vbNullString Then
                THUECL = ThueTKy
            Else
                THUECL = "0"
                ThueTKy = "0"
            End If
            
            If Trim(ThueCQ) = vbNullString Or Trim(ThueCQ) = "" Then
                ThueCQ = "0"
            End If
            
            If Trim(vTemp) <> "" Or Trim(vTemp) <> "''" Then
            sSQLVal = namQT & ",'" & TuKy & "','" & DenKy & "'," & madtnt & "," & MaMuc & "," & maTM & "," & ngnop & "," & NgNhap
            sSQLVal = sSQLVal & "," & LoaiTien & "," & KyDC & "," & ThueTKy & "," & DaGhi & "," & LoaiTK & "," & TTHTK & "," & MaVach & "," & DATHAY
            sSQLVal = sSQLVal & "," & HANNOP & "," & KYLBO & "," & DoanhSo & "," & GTTinhThue & "," & STT & "," & STT2 & "," & STTIN & "," & mact
            sSQLVal = sSQLVal & "," & madlt & "," & LanBS & "," & THUECL & "," & DSoCL & "," & GTTTCL & "," & DANHAN & "," & LAN_QUET & ",'" & TENNV & "','" & CHUNGCHI & "'," & ThueCQ

'            If clsDAO.Connected = False Then
'                clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'                clsDAO.Connect
'            End If

            sSQL = "INSERT INTO TMP_QTTN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            End If
        Next

    End With
    
    clsDAO.Disconnect
    InsertHDR = vbNullString
    
End Function


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        
'        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X)
'        .Row = HEADER_SO_TT_TRONG_TEP_Y
'        .CellType = CellTypeEdit
'        .TypeEditCharSet = TypeEditCharSetNumeric
'        .TypeMaxEditLen = 10
        
        'Ma so tep
'        .Col = .ColLetterToNumber("Q")
'        .Row = 9
'        .CellType = CellTypeEdit
'        .TypeEditCharSet = TypeEditCharSetAlphanumeric
'        .TypeMaxEditLen = 20
'
        'Ghi chu
        .Col = .ColLetterToNumber("E")
        .Row = 36
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("AF")
        .Row = 34
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 32
        .Col = .ColLetterToNumber("R")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
Dim i As Integer, intIndexCombo As Integer
Dim strLTN As Variant, Col7 As Variant, strLTNCu As Variant, strID As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim varID As Variant, varDVT As Variant, varThueSuat As Variant
Dim iTenTN As String, strDVT As String, strThueSuat As String
Dim iCol As Long, iRow As Long
Dim strIdQLT As String, blnSuaThueSuat As Boolean
Dim strMucThueAnDinh As String
Dim countSheet As Integer
Dim maTM As String
    With fps
        .Sheet = .ActiveSheet
        countSheet = 1
        i = 0
        Do
            strIdQLT = ""
            varID = Empty
            varDVT = Empty
            varThueSuat = Empty
            strDVT = vbNullString
            strThueSuat = vbNullString
            
            'Lay gia tri Ma
            .GetText .ColLetterToNumber("AN"), i + 57, varID
            
'            'Lay gia tri DVT
'            .GetText .ColLetterToNumber("O"), i + 39, varDVT
'
'            'Lay gia tri Thue Suat
'            .Col = .ColLetterToNumber("Z")
'            .Row = i + 39
'             varThueSuat = .Value
'
'             'Lay gia tri Muc thue an dinh
'             .Col = .ColLetterToNumber("AC")
'             .Row = i + 39
'             strMucThueAnDinh = .Value
            
            DataDM varID, strIdQLT, iTenTN, strDVT, strThueSuat, blnSuaThueSuat, maTM
            
            If strIdQLT <> vbNullString Then
                'Kiem tra hieu luc danh muc tai nguyen
                If (Not blnSuaThueSuat) And (Val(strMucThueAnDinh) = 0) _
                    And CStr(varThueSuat) <> "" And strThueSuat <> "" Then
                    If Val(strThueSuat) <> Val(CStr(varThueSuat)) Then
                        DisplayMessage "0082", msOKOnly, miCriticalError
                        Exit Function
                    End If
                End If
               .Col = .ColLetterToNumber("C")
                .Row = i + 57
                .Text = iTenTN
                .Col = .ColLetterToNumber("AO")
                .Text = maTM
                'UpdateCell fps, .Col, .Row, .Text

                
            ElseIf CStr(varID) <> vbNullString Then
                DisplayMessage "0082", msOKOnly, miCriticalError
                Exit Function
            End If
          
             i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + 57
            'set ten TN cho muc II
            If .Text = "aa" And countSheet = 1 Then
                i = i + 3
                .Row = i + 57
                countSheet = countSheet + 1
            End If
         Loop Until .Text = "aa"
    End With
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        PHONG_XU_LY_X1 = PHONG_XU_LY_X
        PHONG_XU_LY_Y1 = PHONG_XU_LY_Y
        If rsPXL.Fields.count > 0 Then
            Do While Not rsPXL.EOF
                'dhdang sua combox ngay xu ly
                'ngay 05-08
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE) + "     {" + TAX_Utilities_Svr_New.Convert(Trim(rsPXL.Fields(0).Value), TCVN, UNICODE) + "}"
                .TextTip = TextTipOff
                i = i + 1
                ReDim Preserve larrPhongXuLy(i)
                larrPhongXuLy(i) = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
                ReDim Preserve larrid(i)
                larrid(i) = rsPXL.Fields(0).Value
                rsPXL.MoveNext
            Loop
        Else
            DisplayMessage "0077", msOKOnly
            Exit Function
        End If
        .TypeComboBoxCurSel = 0
        UpdateCell fps, .Col, .Row, .Text
        lSoPhongXL = i
    End With
    Prepared2 = True
End Function

Private Sub DataDM(ByVal Id As String, _
                   Optional ByRef strIdQLT As String, _
                   Optional ByRef TenTN As String, _
                   Optional ByRef DVT As String, _
                   Optional ByRef strThueSuat As String, _
                   Optional ByRef blnSuaThueSuat As Boolean, _
                   Optional ByRef maTMTN As String, _
                   Optional ByRef maTMMT As String)
    Dim arrDanhsach()   As String
    Dim strDataFileName As String
    Dim xmlDOMdata      As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim blChon As Boolean
  
    If TAX_Utilities_Svr_New.Year = "2008" Then
        strDataFileName = Replace(GetCatalogueFileName, "Catalogue_052008_01", "CatalogueQTTN_2008_01")
    ElseIf TAX_Utilities_Svr_New.Year = "2010" Then
        strDataFileName = Replace(GetCatalogueFileName, "Catalogue_072010_01", "CatalogueQTTN_2010_01")
    ElseIf TAX_Utilities_Svr_New.Year = "2014" Then
        strDataFileName = Replace(GetCatalogueFileName, "Catalogue_022014_01", "CatalogueQTTN_2014_01")
    Else
        strDataFileName = GetCatalogueFileName
    End If
    
    If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
        Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")

        For Each xmlNode In xmlNodeListCell

            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")

                If Id = arrDanhsach(1) Then
                    strIdQLT = arrDanhsach(0)
                    TenTN = arrDanhsach(2)
                    DVT = arrDanhsach(3)
                    strThueSuat = arrDanhsach(4)
                    blnSuaThueSuat = IIf(arrDanhsach(5) = "1", True, False)
                    
                    maTMTN = arrDanhsach(7)
                   
                    Exit Sub
                End If
            End If

        Next

    End If

End Sub

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    With fps
        '.EventEnabled(EventAllEvents) = False
        'Set MaSoTep
        .Sheet = 1
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        .Text = strMaSoTep
'        UpdateCell fps, .Col, .Row, .Text
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
'**************************************************************
'       ThanhDX remove
'       Date: 14/09/2006
'       Reason: Doi Nguoi su dung -> Phong quan ly.

        'Set NguoiSuDung
'        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
'        .Row = NGUOI_SU_DUNG_Y
'        If strNguoiSuDung <> "" Then
'            .Text = strNguoiSuDung
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'**************************************************************

         'set Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        If strPhongXuLy <> "" Then
            .Text = strPhongXuLy
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(MA_SO_THUE_X)
        .Row = MA_SO_THUE_Y
        If strMST <> "" Then
            .Text = strMST
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(KY_LAP_BO_X)
        .Row = KY_LAP_BO_Y
        If Len(Month(Now())) = 1 Then
            .Text = "0" & Month(Now()) & "/" & Year(Now())
            sKyKeKhai = "0" & Month(Now()) & Year(Now())
        Else
            .Text = Month(Now()) & "/" & Year(Now())
            sKyKeKhai = Month(Now()) & Year(Now())
        End If
        UpdateCell fps, .Col, .Row, .Text
        
        .Col = .ColLetterToNumber(NGAY_NOP_X)
        .Row = NGAY_NOP_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
         UpdateCell fps, .Col, .Row, .Text
        .Lock = False
        
        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text
       
        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y
        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y
        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(DIEN_THOAI_X)
        .Row = DIEN_THOAI_Y
        If strDThoai <> "" Then
            .Text = strDThoai
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(FAX_X)
        .Row = FAX_Y
        If strFax <> "" Then
            .Text = strFax
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        
                
        'set thong tin dai ly
        
         .Col = .ColLetterToNumber(TEN_GOI_DL_X)
        .Row = TEN_GOI_DL_Y
        If strTen_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTen_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        'mst
        .Col = .ColLetterToNumber(MA_SO_THUE_DL_X)
        .Row = MA_SO_THUE_DL_Y
        If strMST_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strMST_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        'dia chi
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y
        If strDchi_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        'dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y
        If strDthoai_DLT <> vbNullString Then
            .Text = strDthoai_DLT
        End If
        'fax
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y
        If strFax_DLT <> "" Then
            .Text = strFax_DLT
        End If
        ' mail
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y
        If strMail_DLT <> "" Then
            .Text = strMail_DLT
        End If
        ' so hop dong
        .Col = .ColLetterToNumber(SO_HDDL_DL_X)
        .Row = SO_HDDL_DL_Y
        If strSoHD_DLT <> "" Then
            .Text = strSoHD_DLT
        End If
        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HDDL_DL_X)
        .Row = NGAY_HDDL_DL_Y
        If strNgayHD_DLT <> "" Then
            .Text = strNgayHD_DLT
        End If
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        
        ' lan quet
        .Col = .ColLetterToNumber("R")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
        
        '.EventEnabled(EventAllEvents) = True
    End With
    
    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)
Dim strKyLapBo As String
Dim blnValid As Boolean

    blnValid = True
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .Col = .ColLetterToNumber(KY_LAP_BO_X)
        .Row = KY_LAP_BO_Y
        strKyLapBo = .Text
        
        'Phuc vu Tinh So Lan Quet
        'KyTTTu
        .GetText .ColLetterToNumber("M"), 49, strTuKy
        'KyTTTu
        .GetText .ColLetterToNumber("N"), 49, strDenKy
        
        ' Get MaSoTep
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        strMaSoTep = .Text
        
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        strPhongXuLy = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        strNgayNhanToKhai = .Text
        
        'Go to last sheet (header sheet)
        .Sheet = .SheetCount
        
        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
            blnValid = False
        End If
        
'        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_X)
'        .Row = HEADER_KY_LAP_BO_Y
        
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Function GetLastMonthOfThreeMonth(strPeriod As String) As String
    Select Case strPeriod
        Case "01"
            GetLastMonthOfThreeMonth = "03"
        Case "02"
            GetLastMonthOfThreeMonth = "06"
        Case "03"
            GetLastMonthOfThreeMonth = "09"
        Case "04"
            GetLastMonthOfThreeMonth = "12"
    End Select
End Function

Private Sub fps_Click(ByVal Col As Long, ByVal Row As Long)
    With fps
        .Sheet = .ActiveSheet
        .Col = .ActiveCol
        .Row = .ActiveRow
        If Not (.CellType = CellTypeCheckBox Or .CellType = CellTypeButton) Then
            GetCellSpan fps, Col, Row
            fps.SetActiveCell Col, Row
        End If
    End With
End Sub

'Private Sub fps_KeyPress(KeyAscii As Integer)
''    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
''        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
'End Sub

Public Function CheckValidData() As Boolean
    Dim NGAYNOP As Variant, ngayHienTai As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 32, NGAYNOP
        ngayHienTai = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        ngayHienTai = DateSerial(Int(Mid$(ngayHienTai, 7, 4)), Int(Mid$(ngayHienTai, 4, 2)), Int(Mid$(ngayHienTai, 1, 2)))
        NGAYNOP = DateSerial(Int(Mid$(NGAYNOP, 7, 4)), Int(Mid$(NGAYNOP, 4, 2)), Int(Mid$(NGAYNOP, 1, 2)))
        If NGAYNOP > ngayHienTai Then
            CheckValidData = False
        End If

    End With
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
    Dim varTemp As Variant

    With fps

        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("E") And Row = 32 Then
                .GetText Col, Row, varTemp

                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        
                        'check ngay nop khong lon hon ngay hien tai
'                          If DateDiff("D", Format(varTemp, "dd/mm/yyyy"), Format(ToDate(CStr(Date), DDMMYYYY), "dd/mm/yyyy")) < 0 Then
'                            .Sheet = 1
'                             DisplayMessage "0159", msOKOnly, miCriticalError
'                            .SetActiveCell Col, Row
'                          End If
                    Else
                        .SetActiveCell Col, Row
                    End If

                Else
                    .SetText Col, Row, ""
                End If

                UpdateCell fps, .Col, .Row, .Text
           
            End If
            
            If Col = .ColLetterToNumber("E") And Row = 30 Then
                .GetText Col, Row, varTemp

                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If

                Else
                    .SetActiveCell Col, Row
                End If

                UpdateCell fps, .Col, .Row, .Text
            End If
        End If

    End With

End Sub

Public Function TinhSoLanQuet() As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1 As Variant
    
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
'-----------------------------
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTTN" & Right(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If

            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QTTN" & Right(sKyKeKhai, 4) & _
        " WHERE  MADTNT = " & "'" & strMST & "'" & _
        " AND  LOAITK='02/TAIN' AND NAMQT =" & TAX_Utilities_Svr_New.Year & _
        " AND TUKY ='" & strTuKy & "' AND DENKY ='" & strDenKy & "'"
            Set rs = clsDAO.Execute(sSQL)
            If Not rs Is Nothing Then
                 LAN_QUET1 = rs.Fields("LAN_QUET")
                 TinhSoLanQuet = LAN_QUET1
            Else
                 TinhSoLanQuet = 0
            End If
            clsDAO.Disconnect
    Else
    TinhSoLanQuet = 0
    End If
End Function


'dhdang tinh ma nganh
'ngay 25/10/10
Public Function TinhMaNganh() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim NGANH As Variant
    Dim TEMP As New clsADO

        If TEMP.Connected = False Then
            TEMP.CreateConnectionString spathVat & "\dtnt\"
            TEMP.Connect
         End If
         sSQL = "SELECT nganh FROM dtnt2 WHERE MADTNT =  '" & strMST & "'"
        
        Set rs = TEMP.Execute(sSQL)
        If Not rs Is Nothing Then
            NGANH = rs.Fields("nganh")
               
            TinhMaNganh = NGANH
        Else
            TinhMaNganh = 0
        End If
        TEMP.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vKeKhai As Variant
    Dim vKeKhai_temp As Variant
    Dim vKyLB As Variant
    'them bien de kiem tra su ton tai cua bang dl
    Dim strTestFileName As String
    Dim fso As New FileSystemObject
    
    Dim flag As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = sKyKeKhai
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   Do While flag
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTTN" & Right(vKeKhai_temp, 4) & ".DBF"
        'kiem tra neu ton tai bang thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
             sSQL = "SELECT * FROM TMP_QTTN" & Right(vKeKhai_temp, 4) & _
             " WHERE  MADTNT = " & "'" & strMST & "'" & _
             " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4') AND LoaiTK = '02/TAIN'"
        
             Set rs = clsDAO.Execute(sSQL)
        End If
        If rs Is Nothing Then
             isToKhaiChinhThuc = False
        Else
             isToKhaiChinhThuc = True
             Exit Do
        End If
        
        If vKeKhai < vKyLB Then
            vKyLB = DateAdd("M", -1, vKyLB)
            ' ghep ten bang de kiem tra du lieu
            If Len(Month(vKyLB)) = 1 Then
                vKeKhai_temp = "0" & Month(vKyLB) & Year(vKyLB)
            Else
                vKeKhai_temp = Month(vKyLB) & Year(vKyLB)
            End If
            flag = True
        Else
            flag = False
        End If
   Loop
   clsDAO.Disconnect
End Function

