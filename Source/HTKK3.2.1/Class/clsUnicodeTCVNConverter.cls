VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnicodeTCVNConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CharacterSet
   UNICODE = 1
   VPS = 2
   VISCII = 3
   VNI = 4
   TCVN = 5
   VIQR = 6
End Enum

Private mDOMTextFile As MSXML.DOMDocument       ' Document Object Model
Private mXMLPath As String                      ' XML filename
Private mPopulate() As String

Private Function GetLocalDirectory() As String
' Obtain the folder where the program resides
    Dim TStr
    
    ' Get folder where the Exe of this program resides
    TStr = GetAbsolutePath("..\References\") ' App.path
    ' Append a backslash if folder does not end with one
    If Right(TStr, 1) <> "\" Then TStr = TStr & "\"
    GetLocalDirectory = TStr ' Return the folder
End Function

Private Function ReadUnicode(TXMLPath)
' Read Unicode text from XML file
    Dim objTextFileRoot As IXMLDOMElement
    
    Set mDOMTextFile = New MSXML.DOMDocument
    ' Remember the XML file name to update later
    mXMLPath = TXMLPath
    ' Read the XML file and create a DOM
    mDOMTextFile.Load mXMLPath
    'start at the root element of the XML
    Set objTextFileRoot = mDOMTextFile.documentElement
    ' Return the root node's text
    ReadUnicode = objTextFileRoot.nodeTypedValue
End Function

Private Sub WriteUnicode(OutText, Optional TXMLPath)
' Update Unicode Text of same XML file or write to another XML file
    If IsMissing(TXMLPath) Then
        ' Assign the Text to write out to the Root node of DOM
        mDOMTextFile.documentElement.Text = OutText
        ' Update the XML file
        mDOMTextFile.save mXMLPath
    Else
        ' Read a dummy file to create a DOM
        ReadUnicode GetLocalDirectory & "Dummy.xml"
        ' Assign the Text to write out to the Root node of DOM
        mDOMTextFile.documentElement.Text = OutText
        ' Write the Text to the given XML file
        mDOMTextFile.save TXMLPath
    End If
End Sub

Private Function ReadTextFile(FileName) As String
   Dim Fs As FileSystemObject
   Dim TS As TextStream
   '  Create a FileSystem Object
   Set Fs = CreateObject("Scripting.FileSystemObject")
   ' Open TextStream for Input
   Set TS = Fs.OpenTextFile(FileName, ForReading, False, TristateUseDefault)
   ReadTextFile = TS.ReadAll  ' Read the whole content of the text file in one stroke
   TS.Close ' Close the Text Stream
   Set Fs = Nothing  ' Dispose FileSystem Object
End Function

Private Sub WriteTextFile(FileName, StrOutText)
   Dim Fs As FileSystemObject
   Dim TS As TextStream
   '  Create a FileSystem Object
   Set Fs = CreateObject("Scripting.FileSystemObject")
   ' Open TextStream for Output, create file if necesssary
   Set TS = Fs.OpenTextFile(FileName, ForWriting, True, TristateUseDefault)
   TS.Write StrOutText  ' Write the whole StrOutText string in one stroke
   TS.Close ' Close the Text Stream
   Set Fs = Nothing  ' Dispose FileSystem Object
End Sub

Private Function GetVowels(pCharSet As CharacterSet, Optional pFlag As Boolean = False) As String
    Select Case pCharSet
        Case TCVN
            GetVowels = ReadTextFile(GetLocalDirectory & "TCVNVowels.txt")
        Case UNICODE
            GetVowels = ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
        Case VIQR
            If pFlag = False Then
                GetVowels = ReadTextFile(GetLocalDirectory & "VIQRVowels.txt")
            Else
                GetVowels = ReadTextFile(GetLocalDirectory & "VIQRVowelMap.txt")
            End If
        Case VISCII
            GetVowels = ReadTextFile(GetLocalDirectory & "VISCIIVowels.txt")
        Case VNI
            If pFlag = False Then
                GetVowels = ReadTextFile(GetLocalDirectory & "VNIVowels.txt")
            Else
                GetVowels = ReadTextFile(GetLocalDirectory & "VNIVowelMap.txt")
            End If
        Case VPS
            GetVowels = ReadTextFile(GetLocalDirectory & "VPSVowels.txt")
    End Select
End Function

Private Function StringToString(pString As String, Vowel1, Vowel2) As String
' Direct one-to-one character mapping from one encoding to another
Dim letter As String
Dim Text1 As String
Dim Text2 As String
Dim i, Pos
    ' Use Text1 to execute  a litle faster than TextBox1(0)
    Text1 = pString
    ' Iterate through each character of the from Text string
    For i = 1 To Len(Text1)
        letter = Mid(Text1, i, 1)
        ' Leave Carriage Return and Line Feed characters as is
        If (letter = vbCr) Then
           Text2 = Text2 & vbCr
        ElseIf (letter = vbLf) Then
           Text2 = Text2 & vbLf
        Else
        ' Find position of character in the vowel list
             Pos = InStr(Vowel1, letter)
             If Pos <= 0 Then
               ' Not found - so do  not map
               Text2 = Text2 & letter
             Else
               ' Found - so pick the corresponding character in the other vowel list
               Text2 = Text2 & Mid(Vowel2, Pos, 1)
             End If
        End If
    Next
    StringToString = Text2
End Function

Private Sub PopulateFromFile(pCharSet As CharacterSet, Optional pFlag As Boolean = False)
    Dim strFile As String, strALine As String
    Dim i As Long
    
    Select Case pCharSet
        Case TCVN
            strFile = GetVowels(TCVN)
        Case UNICODE
            strFile = GetVowels(UNICODE)
        Case VIQR
            strFile = GetVowels(VIQR, pFlag)
        Case VISCII
            strFile = GetVowels(VISCII)
        Case VNI
            strFile = GetVowels(VNI, pFlag)
        Case VPS
            strFile = GetVowels(VPS)
    End Select
    mPopulate = Split(strFile, Chr(13) & Chr(10))
End Sub

Private Function UnicodeToMultichar(pString As String, UVowels As String) As String
' Convert Unicode vowels to an encoding like VNI or VIQR where it might take more than one
'   characters to represent a vowel.
Dim letter As String
Dim Text1 As String
Dim Text2 As String
Dim i, Pos
    ' Assign content of input Textbox to Text1
    Text1 = pString
    ' Iterate through every Unicode character, don't care if it take 1,2 or 3 characters internally
    For i = 1 To Len(Text1)
       ' get a Unicode character
       letter = Mid(Text1, i, 1)
        ' if it's a Carriage return or a LineFeed then just copy across
        If (letter = vbCr) Then
           Text2 = Text2 & vbCr
        ElseIf (letter = vbLf) Then
           Text2 = Text2 & vbLf
        Else
        ' can we find it in the Unicode vowel list?
             Pos = InStr(UVowels, letter)
             If Pos <= 0 Then
               ' if not then just copy across as is
               Text2 = Text2 & letter
             Else
               ' convert to the corresponding multi-character vowel of VNI or VIQR
               Text2 = Text2 & mPopulate(Pos - 1)
             End If
        End If
    Next
    'return the result
    UnicodeToMultichar = Text2
End Function

Private Function MulticharToUnicode(pString As String, UVowels As String) As String
' Convert a multi-character vowel in VNI or VIQR to Unicode
Dim letter As String
Dim Text1 As String
Dim Text2 As String
Dim i, Pos, Item, MapNum, TLen
    ' Assign content of input Textbox to Text1
    Text1 = pString
    ' Replace every multi-character vowel in Text1 with a string like |067 that represents
    ' the 67th Unicode vowel
    ' Iterate through every multi-character vowel
    For i = 0 To UBound(mPopulate) - 1
       ' Get an item from the Vowel listbox
       Item = mPopulate(i)   ' like a^~016
       letter = RTrim(Left(Item, 3))  ' isolate the multi-character vowel eg: a^~
       MapNum = "|" & Right(Item, 3)  ' Prefix the | character to the digit string eg: &016
       Text1 = Replace(Text1, letter, MapNum) ' replace all occurences of the vowel
    Next
    ' Now map the position strings like &016, &114 to 16th and 114th Unicode vowels
    i = 1
    TLen = Len(Text1)
    Do While i <= TLen
       ' get a character
       letter = Mid(Text1, i, 1)
        ' if it's a Carriage return or a LineFeed then just copy across
        If (letter = vbCr) Then
           Text2 = Text2 & vbCr
               i = i + 1
        ElseIf (letter = vbLf) Then
           Text2 = Text2 & vbLf
               i = i + 1
        ElseIf letter <> "|" Then
            ' merely copy across everything else
               Text2 = Text2 & letter
               i = i + 1
        Else
           ' get here if encounter a "&", obtain the position of the Unicode vowel
           ' Note that there'll be a bug if the text string contains genuine "|" character
           Pos = Val(Mid(Text1, i + 1, 3))
           If Pos = 0 Then
               Text2 = Text2 & letter
               i = i + 1
           Else
             ' get the Unicode vowel for output
               Text2 = Text2 & Mid(UVowels, Pos, 1)
               i = i + 4
           End If
        End If
    Loop
    ' Return the result
    MulticharToUnicode = Text2
End Function

Public Function Convert(pString As String, pFrom As CharacterSet, pTo As CharacterSet) As String
    Dim UVowels As String
    Dim Vowels As String
    
    If pString = vbNullString Or pFrom = pTo Then
        Convert = pString
        Exit Function
    End If
    
    UVowels = GetVowels(UNICODE)
    If pFrom = UNICODE Then
        Vowels = GetVowels(pTo)
        Select Case pTo
            Case TCVN
                Convert = StringToString(pString, UVowels, Vowels)
            Case VIQR
                PopulateFromFile VIQR
                Convert = UnicodeToMultichar(pString, UVowels)
            Case VISCII
                Convert = StringToString(pString, UVowels, Vowels)
            Case VNI
                PopulateFromFile VNI
                Convert = UnicodeToMultichar(pString, UVowels)
            Case VPS
                Convert = StringToString(pString, UVowels, Vowels)
        End Select
    Else
        Vowels = GetVowels(pFrom)
        Select Case pFrom
            Case TCVN
                Convert = StringToString(pString, Vowels, UVowels)
            Case VIQR
                PopulateFromFile VIQR, True
                Convert = MulticharToUnicode(pString, UVowels)
            Case VISCII
                Convert = StringToString(pString, Vowels, UVowels)
            Case VNI
                PopulateFromFile VNI, True
                Convert = MulticharToUnicode(pString, UVowels)
            Case VPS
                Convert = StringToString(pString, Vowels, UVowels)
        End Select
    End If
End Function

