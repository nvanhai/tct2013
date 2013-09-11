VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLRequestNuic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim bDatos() As String
    Dim iCont As Integer
    iCont = 1
    
    
'    Dim testXml As New MSXML.DOMDocument
'    Dim testXmlHttp As New MSXML.XMLHTTPRequest

    On Error GoTo Err_PW
    ' Creating the DOMDocument and XMLHTTP Object´s
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    ' Load the XML
    objDom.async = False
    objDom.loadXML XmlBody
    'Open the Web service Connection
    objXmlHttp.open "POST", AsmxUrl, False
    ' creating the XML header´s
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    'objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    
    ' Sending the command to open the XML request
    objXmlHttp.send objDom.xml
    ' Get the result and We Stored in the "strRet" variable
    strRet = objXmlHttp.responseText

'    testXml.loadXML XmlBody
'    testXmlHttp.open "POST", AsmxUrl, False
'    testXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    testXmlHttp.send testXml.xml
'    strRet = testXmlHttp.responseText

    PostWebservice = strRet
    'Set testXmlHttp = Nothing
    ' Closing the object
    Set objXmlHttp = Nothing
Exit Function
'' In case of Error, throw a exception
Err_PW:
    PostWebservice = "Error: " & Err.Number & " - " & Err.Description
End Function

Public Function buscarXML(ByVal tag As String, ByVal xml As String) As String
    Dim tag1, tag2
    tag1 = "<" & UCase(tag) & ">"
    tag2 = "</" & UCase(tag) & ">"
    iPosI = InStr(UCase(xml), tag1) + Len(tag1)
    iPosF = InStr(UCase(xml), tag2)
    If iPosI = 0 Or iPosF = 0 Then
        buscarXML = "N/A"
    Else
        buscarXML = Mid(xml, iPosI, iPosF - iPosI)
    End If
End Function
