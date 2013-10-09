Attribute VB_Name = "DateUtil"
Option Explicit

Public Const DDMMYYYY = "DD/MM/YYYY"
Public Const DDMM = "DD/MM"
Public Const MMYYYY = "MM/YYYY"
Public Const YYYY = "YYYY"
Public Const DD = "DD"
Public Const MM = "MM"

Private Function IsDate(Y As Integer, m As Integer, d As Integer) As Variant
    Dim dDateTemp As Date
    
    IsDate = Null
    
    dDateTemp = DateSerial(Y, m, d)
    
    If d = Day(dDateTemp) And m = Month(dDateTemp) And Year(dDateTemp) Then
        IsDate = dDateTemp
    End If
End Function

Public Function ToDate(strDate As String, strFormat As String) As Variant
    Dim arrDateUnit() As String
    Dim d As Integer
    Dim m As Integer
    Dim Y As Integer
    Dim i As Integer
    
    ToDate = Null
    arrDateUnit = Split(strDate, "/")
    For i = 0 To UBound(arrDateUnit)
        arrDateUnit(i) = Trim(arrDateUnit(i))
    Next
    
    d = 0
    m = 0
    Y = 0
    
    Select Case LCase(strFormat)
        Case LCase(DDMMYYYY)
            If UBound(arrDateUnit) = 2 Then
                d = Val(arrDateUnit(0))
                m = Val(arrDateUnit(1))
                Y = Val(arrDateUnit(2))
            ElseIf UBound(arrDateUnit) = 0 And (Len(strDate) = 6 Or Len(strDate) = 8) Then
                d = Val(Mid(strDate, 1, 2))
                m = Val(Mid(strDate, 3, 2))
                Y = Val(Mid(strDate, 5))
            End If
        Case LCase(DDMM)
            If UBound(arrDateUnit) = 1 Then
                d = Val(arrDateUnit(0))
                m = Val(arrDateUnit(1))
                Y = Year(Now)
            End If
        Case LCase(MMYYYY)
            If UBound(arrDateUnit) = 1 Then
                d = 1
                m = Val(arrDateUnit(0))
                Y = Val(arrDateUnit(1))
                ToDate = IsDate(Y, m, d)
            End If
        Case LCase(YYYY)
            If UBound(arrDateUnit) = 0 Then
                d = 1
                m = 1
                Y = Val(arrDateUnit(0))
            End If
        Case LCase(DD)
            If UBound(arrDateUnit) = 0 Then
                d = Val(arrDateUnit(0))
                m = Month(Now)
                Y = Year(Now)
            End If
        Case LCase(MM)
            If UBound(arrDateUnit) = 0 Then
                d = 1
                m = Val(arrDateUnit(0))
                Y = Year(Now)
            End If
    End Select
    If d <> 0 And m <> 0 Then
        If Y < 100 Then
            Y = Y + 2000
        End If
        ToDate = IsDate(Y, m, d)
    End If
End Function

Public Function ToString(dDate As Variant, strFormat) As Variant
    Dim dUnit As String
    Dim mUnit As String
    Dim yUnit As String
    
    ToString = Null
    dUnit = Trim(Str(Day(dDate)))
    mUnit = Trim(Str(Month(dDate)))
    yUnit = Trim(Str(Year(dDate)))
    
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
        Case LCase(YYYY)
            ToString = yUnit
        Case LCase(DD)
            ToString = dUnit
        Case LCase(MM)
            ToString = mUnit
    End Select
End Function


