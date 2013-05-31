VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03TAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 9
Const KY_LAP_BO_X = "I"
Const MA_SO_TEP_Y = 9
Const MA_SO_TEP_X = "T"
Const NGAY_NHAN_TO_KHAI_Y = 11
Const NGAY_NHAN_TO_KHAI_X = "I"
'Const NGUOI_SU_DUNG_Y = 12
'Const NGUOI_SU_DUNG_X = "V"
'Const HEADER_KY_LAP_BO_Y = 13
'Const HEADER_KY_LAP_BO_X = "B"
'Const HEADER_SO_TT_TRONG_TEP_X = "V"
'Const HEADER_SO_TT_TRONG_TEP_Y = 10
'Const HEADER_SO_TT_TRONG_TEP_X = "V"
'Const HEADER_SO_TT_TRONG_TEP_Y = 16 'Dong 16 da dc Hidden
Const PHONG_XU_LY_Y = 9
Const PHONG_XU_LY_X = "AF"

'Longvh
Const MA_SO_THUE_X = "K"
Const MA_SO_THUE_Y = 4
Const NGAY_NOP_Y = 11
Const NGAY_NOP_X = "I"
Const NGAY_QUET_Y = 11
Const NGAY_QUET_X = "T"

Const TEN_GOI_X = "K"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "K"
Const DIA_CHI_Y = 6

Const TEN_BPQL_X = "AF"
Const TEN_BPQL_Y = 11
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 16 'Dong 16 da dc Hidden


Const DIEN_THOAI_X = "I"
Const DIEN_THOAI_Y = 7
Const FAX_X = "T"
Const FAX_Y = 7

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Private LAN_QUET As Variant
'Longvh
Public strMST As String
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public TTHTK As String
Public strTenBpql As String

Public Function InsertDTL_KHBS() As String
'    Dim sSQL As String
'    Dim sSQLCol As String
'    Dim sSQLVal As String
'    Dim bln  As Boolean
'    Dim MADTNT As Variant
'    Dim NGNOP As Variant
'    Dim KYLBO As Variant
'    Dim KYKKHAI As Variant
'    Dim MATKHAI As Variant
'    Dim MATHUE As Variant
'    Dim MAMUC As Variant
'    Dim MATM As Variant
'    Dim MAPP As Variant
'    Dim MACT As Variant
'    Dim MACT2 As Variant
'    Dim MACT3 As Variant
'    Dim STT As Variant
'    Dim STT2 As Variant
'    Dim STTIN As Variant
'    Dim SOKK As Variant
'    Dim SODC As Variant
'    Dim SOCL As Variant
'    Dim i As Integer, iCol As Long, iRow As Long
'    Dim xmlNode As MSXML.IXMLDOMNode
'
'    Dim strFileName As String
'    Dim fso As New FileSystemObject
'
'   'kiem tra ton tai tep *.dbf chua
'    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileName) = False Then
'    Dim strTepmau As String
'        strTepmau = spathVat & "\tepmau\BS.DBF"
'        fso.CopyFile strTepmau, strFileName, False
'    End If
'
'
'   sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & _
'             " SOKK, SODC, SOCL"
'
'   If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("G"), 4, MADTNT
'        If Trim(MADTNT) = vbNullString Then
'            MADTNT = "''"
'        Else
'            MADTNT = "'" & MADTNT & "'"
'        End If
'
'        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        MATKHAI = "'03/TNDN'"
'        MATHUE = "'01'"
'        MAMUC = "'014'"
'        MATM = "'01'"
'        MAPP = "'2'"
'
'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 10, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "CTOD('')"
'        Else
'            KYLBO = "'" & KYLBO & "'"
'        End If
'
'' Insert so dieu chinh tang
'
'        MACT = "'1000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 1
'        STT2 = 0
'        STTIN = "'I'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'        .Sheet = .SheetCount - 1
'        i = 1
'        .Row = 24
'        Do
'                MACT = "''"
'                .GetText .ColLetterToNumber("E"), .Row, MACT2
'                If Trim(MACT2) = vbNullString Then
'                    MACT2 = "''"
'                ElseIf MACT2 = "20" Then
'                    MACT2 = "'062'"
'                ElseIf MACT2 = "21" Then
'                    MACT2 = "'076'"
'                End If
'
'                MACT3 = "'2'"
'                STT = 1
'                STT2 = i
'                STTIN = "'" & CStr(i) & "'"
'
'                .GetText .ColLetterToNumber("F"), .Row, SOKK
'                If Trim(SOKK) = vbNullString Then
'                    SOKK = "0"
'                Else
'                    SOKK = SOKK
'                End If
'
'                .GetText .ColLetterToNumber("G"), .Row, SODC
'                If Trim(SODC) = vbNullString Then
'                    SODC = "0"
'                Else
'                    SODC = SODC
'                End If
'
'                .GetText .ColLetterToNumber("H"), .Row, SOCL
'                If Trim(SOCL) = vbNullString Then
'                    SOCL = "0"
'                Else
'                    SOCL = SOCL
'                End If
'
'                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                          NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                          SOKK & "," & SODC & "," & SOCL
'
'                sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'                If MACT2 <> "''" Then
'                    bln = clsDAO.ExecuteDLL(sSQL)
'                End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 1 + .Row
'            i = i + 1
'        Loop Until .Text = "aa"
'' Insert so dieu chinh giam
'
'        MACT = "'2000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 2
'        STT2 = 0
'        STTIN = "'II'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'        i = 1
'        .Row = .Row + 3
'         Do
'                MACT = "''"
'                .GetText .ColLetterToNumber("E"), .Row, MACT2
'                If Trim(MACT2) = vbNullString Then
'                    MACT2 = "''"
'                ElseIf MACT2 = "20" Then
'                    MACT2 = "'062'"
'                ElseIf MACT2 = "21" Then
'                    MACT2 = "'076'"
'                End If
'
'                MACT3 = "'2'"
'                STT = 2
'                STT2 = i
'                STTIN = "'" & CStr(i) & "'"
'
'                .GetText .ColLetterToNumber("F"), .Row, SOKK
'                If Trim(SOKK) = vbNullString Then
'                    SOKK = "0"
'                Else
'                    SOKK = SOKK
'                End If
'
'                .GetText .ColLetterToNumber("G"), .Row, SODC
'                If Trim(SODC) = vbNullString Then
'                    SODC = "0"
'                Else
'                    SODC = SODC
'                End If
'
'                .GetText .ColLetterToNumber("H"), .Row, SOCL
'                If Trim(SOCL) = vbNullString Then
'                    SOCL = "0"
'                Else
'                    SOCL = SOCL
'                End If
'
'                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                          NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                          SOKK & "," & SODC & "," & SOCL
'
'                sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'                If MACT2 <> "''" Then
'                    bln = clsDAO.ExecuteDLL(sSQL)
'                End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 1 + .Row
'            i = i + 1
'        Loop Until .Text = "bb"
'
'' Insert Tong hop dieu chinh
'
'        MACT = "'3000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 3
'        STT2 = 0
'        STTIN = "'III'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'   End With
'
'   clsDAO.Disconnect
   InsertDTL_KHBS = vbNullString
   
End Function


Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
   Dim NAMQT As Variant
   Dim TUKY As Variant
   Dim DENKY As Variant
   Dim MADTNT As Variant
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim NGNOP As Variant
   Dim NGNHAP As Variant
   Dim LOAITIEN As Variant
   Dim KYDC As Variant
   Dim THUEKY As Variant
   Dim DAGHI As Variant
   Dim LOAITK As Variant
   Dim TEMP As Variant
   Dim KYLBO As Variant
   Dim DATHAY As Variant
   Dim MAVACH As Variant
   Dim DOANHSO As Variant
   Dim GTTINHTHUE As Variant
   
   Dim bln  As Boolean
   Dim NGANH As Variant
    Dim DAGHI1 As Variant
    Dim DAGHI2 As Variant
    Dim AUTO_CAL As Variant
    Dim DANHAN As Variant
   'test
    Dim arrMTM() As String, arrMTMSave() As String
    Dim arrVitri() As Integer
    Dim idx, idx1, i, countMTM As Integer
    Dim vTemp As String, vVitriTrung As String
    'end test
   
    sSQLCol = " NAMQT, TUKY, DENKY, MADTNT, MAMUC, MATM, NGNOP, NGNHAP, LOAITIEN, KYDC, THUETKY, DAGHI, LOAITK, TTHTK, KYLBO, DATHAY, MAVACH, HANNOP, DOANHSO, GTTINHTHUE ,LAN_QUET, DANHAN"
       
    If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
    End If
       
    With fps
        .Sheet = 1
        
       
        
        'Lay ds ma tieu muc
        'dhdang phan lay vao mang
        'ngay 14/09/2010
        .Col = .ColLetterToNumber("B")
        .Row = 38
        Do
            ' Lay ma tieu muc tren cel
            .Col = .ColLetterToNumber("C")
            vTemp = vTemp & .Text & "~"
                     
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
         Loop Until .Text = "aa"
         arrMTM = Split(vTemp, "~")
         vTemp = ""
         ReDim arrMTMSave(1)
         arrMTMSave(0) = arrMTM(0)
         For idx = 1 To UBound(arrMTM) - 1
            countMTM = 0
            vTemp = arrMTM(idx)
            For i = 0 To UBound(arrMTMSave)
                If vTemp = arrMTMSave(i) Then
                    countMTM = countMTM + 1
                End If
            Next
            If countMTM = 0 Then
                ReDim Preserve arrMTMSave(UBound(arrMTMSave) + 1)
                arrMTMSave(UBound(arrMTMSave) - 1) = vTemp
            End If
         Next
        '---------------------------------------------------------
        
        
        .GetText .ColLetterToNumber("K"), 4, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
       
        .GetText .ColLetterToNumber("I"), 9, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
    
    
        .GetText .ColLetterToNumber("I"), 11, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("T"), 11, NGNHAP
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
       
'        NGNHAP = Date
'
'        NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
'        NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        
        NAMQT = TAX_Utilities_Svr_New.Year
        TUKY = "'01/" & TAX_Utilities_Svr_New.Year & "'"
        DENKY = "'12/" & TAX_Utilities_Svr_New.Year & "'"
        
          ' sua MAMUC & MATM tu nam 2009
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            MATM = "'01'"
            MAMUC = "'012'"
        Else
            MAMUC = "'1550'"
            MATM = "''"
        End If

        LOAITIEN = "''"
        DAGHI = ".T."
        LOAITK = "'03/TAIN'"
        DANHAN = "''"
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        
        KYDC = KYLBO
        DATHAY = ".F."
        MAVACH = ".T."
        DOANHSO = 0
        GTTINHTHUE = 0
        
        For i = 0 To UBound(arrMTMSave) - 1
            THUEKY = 0
            vTemp = arrMTMSave(i)
            MATM = "'" & vTemp & "'"
            .Col = .ColLetterToNumber("B")
            .Row = 38
                 Do
                     .Col = .ColLetterToNumber("C")
                    If vTemp = Trim(.Text) Then
                            .Col = .ColLetterToNumber("AM")
                            TEMP = .Value
                            If Trim(TEMP) = vbNullString Then
                               TEMP = "0"
                            End If
                             THUEKY = THUEKY + CDbl(TEMP)
                    End If
                    .Col = .ColLetterToNumber("B")
                    .Row = .Row + 1
                  Loop Until .Text = "aa"
        DAGHI = ".T."
        MAVACH = ".T."
        DAGHI1 = ".F."
        DAGHI2 = ".F."
        AUTO_CAL = ".F."
        NGANH = "''"
            
         sSQLVal = NAMQT & "," & TUKY & "," & DENKY & "," & MADTNT & "," & MAMUC & "," & MATM & "," & NGNOP & _
        "," & NGNHAP & "," & LOAITIEN & "," & KYDC & "," & THUEKY & "," & DAGHI & "," & LOAITK & "," & TTHTK & "," & KYLBO & "," & DATHAY & "," & MAVACH & "," & HANNOP & "," & DOANHSO & "," & GTTINHTHUE & "," & LAN_QUET & "," & DANHAN
           
        sSQL = "INSERT INTO TMP_QTTN" & Right(sKyKeKhai, 4) & _
         "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    Next
    End With
    InsertHDR = vbNullString
    
End Function


Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
   Dim NAMQT As Variant
   Dim MADTNT As Variant
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim NGNOP As Variant
   Dim LOAITK As Variant
   Dim NGANHTN As Variant
   Dim SANLUONG As Variant
   Dim DVTINH As Variant
   Dim DONGIA As Variant
   Dim GTTT As Variant
   Dim THUESUAT As Variant
   Dim MUCTHUE As Variant
   Dim THUEPS As Variant
   Dim MG As Variant
   Dim THUETN As Variant
   Dim NGNHAP As Variant
   
   
   
   
   Dim DANHAN As Variant
  
   
   sSQLCol = "NAMQT, MADTNT, MAMUC, MATM, NGNOP, LOAITK, NGANHTN, SANLUONG, DVTINH, DONGIA, GTTT, THUESUAT, MUCTHUE, THUEPS, MG, THUETN, TTHTK, LAN_QUET, DANHAN"
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("K"), 4, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
        
        .GetText .ColLetterToNumber("I"), 11, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
                    
        NAMQT = TAX_Utilities_Svr_New.Year
        
       
        NGNHAP = Date
                
       ' sua MAMUC & MATM tu nam 2009
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            MATM = "'01'"
            MAMUC = "'012'"
        Else
            MAMUC = "'1550'"
        End If
        
        
        LOAITK = "'03/TAIN'"
        DANHAN = "''"
             .Col = .ColLetterToNumber("B")
             .Row = 38
                 Do
                 
                     .Col = .ColLetterToNumber("C")
                     MATM = .Text
                     If Trim(MATM) = vbNullString Then
                        MATM = "''"
                     Else
                        MATM = "'" & MATM & "'"
                     End If
                        
                     .Col = .ColLetterToNumber("D")
                     NGANHTN = .Text
                     If Trim(NGANHTN) = vbNullString Then
                        NGANHTN = "''"
                     Else
                        NGANHTN = "'" & NGANHTN & "'"
                     End If
                     
                    .Col = .ColLetterToNumber("O")
                     DVTINH = .Text
                     If Trim(DVTINH) = vbNullString Then
                           DVTINH = "''"
                     Else
                           DVTINH = "'" & DVTINH & "'"
                     End If
                     
                    .Col = .ColLetterToNumber("R")
                    SANLUONG = .Value
                    If Trim(SANLUONG) = vbNullString Then
                            SANLUONG = "0"
                    End If
                    
                    .Col = .ColLetterToNumber("U")
                    DONGIA = .Value
                    If Trim(DONGIA) = vbNullString Then
                            DONGIA = "0"
                    End If
                    
                    
                    GTTT = CDbl(SANLUONG) * CDbl(DONGIA)
                    
                     
                    .Col = .ColLetterToNumber("Z")
                    THUESUAT = .Value
                    If Trim(THUESUAT) = vbNullString Then
                           THUESUAT = "0"
                    End If
                     
                     .Col = .ColLetterToNumber("AC")
                     MUCTHUE = .Value
                     If Trim(MUCTHUE) = vbNullString Then
                            MUCTHUE = "0"
                     End If

                     .Col = .ColLetterToNumber("AF")
                     THUEPS = .Value
                     If Trim(THUEPS) = vbNullString Then
                        THUEPS = "0"
                     End If

                     .Col = .ColLetterToNumber("AI")
                     MG = .Value
                     If Trim(MG) = vbNullString Then
                        MG = "0"
                     End If
                     
                     .Col = .ColLetterToNumber("AM")
                     THUETN = .Value
                     If Trim(THUETN) = vbNullString Then
                            THUETN = "0"
                     End If
                    
     
                     
                       sSQLVal = NAMQT & "," & MADTNT & "," & MAMUC & "," & MATM & "," & NGNOP & "," & LOAITK & "," & _
                       NGANHTN & "," & SANLUONG & "," & DVTINH & "," & DONGIA & "," & GTTT & "," & _
                       THUESUAT & "," & MUCTHUE & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & LAN_QUET & "," & DANHAN
                       sSQL = "INSERT INTO TMP_QA" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)

                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
                 
             
        End With
   
   


   clsDAO.Disconnect
   
   InsertDTL = vbNullString
   
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
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & Right(sKyKeKhai, 4) & ".DBF"
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

    isExistFile = True
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTTN" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND LOAITK='03/TAIN' AND NAMQT=" & TAX_Utilities_Svr_New.Year
    
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
   Dim NGNOP As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = ""
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
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
        " WHERE  MATM = '01' AND MADTNT = " & "'" & strMST & "'" & _
        " AND LOAITK='03/TAIN' AND NAMQT=" & TAX_Utilities_Svr_New.Year
    
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
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(NGNOP) Then
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
'
'   sSQL = "UPDATE QTTN" & Right(sKyKeKhai, 4) & _
'        " SET SODC=0" & _
'        " WHERE  MATM = '01' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND LOAITK='03/TAIN' AND NAMQT=" & TAX_Utilities_Svr_New.Year
'
'   clsDAO.ExecuteDLL sSQL
   
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
        " WHERE  MATM = '01' AND MADTNT = " & "'" & strMST & "'" & _
        " AND LOAITK='03/TAIN' AND NAMQT='" & TAX_Utilities_Svr_New.Year & "'"
    
   Set rs = clsDAO.Execute(sSQL)
    
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
   sSQL = "DELETE FROM TMP_QA" & Right(sKyKeKhai, 4) & _
        " WHERE MATM = '01' AND MADTNT = " & "'" & strMST & _
        " AND LOAITK='03/TAIN' AND NAMQT='" & TAX_Utilities_Svr_New.Year & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

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
        .Col = .ColLetterToNumber("T")
        .Row = 9
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        .TypeMaxEditLen = 20
        
        'Ghi chu
        .Col = .ColLetterToNumber("T")
        .Row = 13
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("AF")
        .Row = 9
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ky lap bo
        SetDateFormat fps, 1, 9, .ColLetterToNumber("I"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 11, .ColLetterToNumber("I"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 11
        .Col = .ColLetterToNumber("I")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 11, .ColLetterToNumber("T"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 11
        .Col = .ColLetterToNumber("T")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
Dim i As Integer, intIndexCombo As Integer
Dim strLTN As Variant, Col7 As Variant, strLTNCu As Variant, strId As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim varID As Variant, varDVT As Variant, varThueSuat As Variant
Dim iTenTN As String, strDVT As String, strThueSuat As String
Dim iCol As Long, iRow As Long
Dim strIdQLT As String, blnSuaThueSuat As Boolean
Dim strMucThueAnDinh As String
Dim MATM As String
    With fps
        .Sheet = .ActiveSheet
        i = 0
        Do
            strIdQLT = ""
            varID = Empty
            varDVT = Empty
            varThueSuat = Empty
            strDVT = vbNullString
            strThueSuat = vbNullString
            
            'Lay gia tri Ma
            .GetText .ColLetterToNumber("D"), i + 38, varID
            
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
            
            DataDM varID, strIdQLT, iTenTN, strDVT, strThueSuat, blnSuaThueSuat, MATM
            
            If strIdQLT <> vbNullString Then
                'Kiem tra hieu luc danh muc tai nguyen
                If (Not blnSuaThueSuat) And (Val(strMucThueAnDinh) = 0) _
                    And CStr(varThueSuat) <> "" And strThueSuat <> "" Then
                    If Val(strThueSuat) <> Val(CStr(varThueSuat)) Then
                        DisplayMessage "0082", msOKOnly, miCriticalError
                        Exit Function
                    End If
                End If
               .Col = .ColLetterToNumber("E")
                .Row = i + 38
                .Text = iTenTN
                .Col = .ColLetterToNumber("C")
                .Text = MATM
                'UpdateCell fps, .Col, .Row, .Text

                
            ElseIf CStr(varID) <> vbNullString Then
                DisplayMessage "0082", msOKOnly, miCriticalError
                Exit Function
            End If
          
             i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + 38
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
        If rsPXL.Fields.Count > 0 Then
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

Private Sub DataDM(ByVal Id As String, Optional ByRef strIdQLT As String, Optional ByRef TenTN As String, Optional ByRef DVT As String, Optional ByRef strThueSuat As String, Optional ByRef blnSuaThueSuat As Boolean, Optional ByRef MATM As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode
        
       If TAX_Utilities_Svr_New.Year = "2008" Then
            strDataFileName = Replace(GetCatalogueFileName, "Catalogue_052008_01", "CatalogueQTTN_2008_01")
       ElseIf TAX_Utilities_Svr_New.Year = "2010" Then
            strDataFileName = Replace(GetCatalogueFileName, "Catalogue_072010_01", "CatalogueQTTN_2010_01")
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
                            MATM = arrDanhsach(7)
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
    Dim varCheckValue As Variant, varNoteValue As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 13, varCheckValue
        .GetText .ColLetterToNumber("T"), 13, varNoteValue
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 14, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 14, "1"
        End If
        
    End With
End Function



Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("I") And Row = 11 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("I") And Row = 9 Then
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
            " WHERE MADTNT = " & "'" & strMST & "'" & _
            "AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4') AND LOAITK='03/TAIN' AND NAMQT=" & TAX_Utilities_Svr_New.Year
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

