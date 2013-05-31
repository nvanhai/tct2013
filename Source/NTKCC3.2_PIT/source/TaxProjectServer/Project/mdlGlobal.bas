Attribute VB_Name = "mdlGlobal"
Option Explicit


Public spathVat As String
Public spathQHSCC As String
Public gcnSource  As New ADODB.Connection

Public Sub ReadPathFile(ByVal strdir As String)
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim txt(3) As String
    Dim I As Integer
    Set fs = New FileSystemObject
        If fs.FileExists(strdir) Then
        Set ts = fs.OpenTextFile(strdir)
        I = 0
        Do While Not ts.AtEndOfStream
           
            strFile(I) = GiaiMa(ts.ReadLine)
            I = I + 1
        Loop
        ts.Close
    End If
    
    
End Sub

Public Sub WritePathFile(ByVal strdir As String, ByVal str As String)
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim txt As Variant
    txt = Split(str, ",")
    Set fs = New FileSystemObject
        If fs.FileExists(strdir) Then
            fs.DeleteFile (strdir)
            fs.CreateTextFile (strdir)
            Set ts = fs.OpenTextFile(strdir, ForWriting, True)
            ts.WriteLine MaHoa(txt(0))
            ts.WriteLine MaHoa(txt(1))
            ts.WriteLine MaHoa(txt(2))
            ts.WriteLine MaHoa(txt(3))
            ts.Close
        Else
            fs.CreateTextFile (strdir)
            Set ts = fs.OpenTextFile(strdir, ForWriting, True)
            ts.WriteLine MaHoa(txt(0))
            ts.WriteLine MaHoa(txt(1))
            ts.WriteLine MaHoa(txt(2))
            ts.WriteLine MaHoa(txt(3))
            ts.Close
        End If
    
End Sub
Public Function MaHoa(ByVal sStr) As String
    Dim I As Integer
    Dim temp As String
    temp = ""
    If Trim(sStr) = vbNullString Then
        MaHoa = ""
        Exit Function
    End If
    For I = 1 To Len(sStr)
        temp = temp & Chr(Asc(Mid(sStr, I, 1)) + 5)
    Next
    MaHoa = temp
End Function

Public Function GiaiMa(ByVal sStr) As String
    Dim I As Integer
    Dim temp As String
    temp = ""
    If Trim(sStr) = vbNullString Then
        GiaiMa = ""
        Exit Function
    End If
    For I = 1 To Len(sStr)
        temp = temp & Chr(Asc(Mid(sStr, I, 1)) - 5)
    Next
    GiaiMa = temp
End Function


