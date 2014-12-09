Attribute VB_Name = "mdlVariables"
Option Explicit

Public bolAdjustData() As Boolean
Public xmlData() As MSXML.DOMDocument                              ' xmlNode for data (1 sheet = 1 item of this array)
Public xmlNodeMenu As MSXML.IXMLDOMElement                         ' xmlNode for menu (init from frmSystem)
Public xmlNodeValidity As MSXML.IXMLDOMElement                     ' xmlNode for validity
Public xmlNodeListMessage As MSXML.IXMLDOMNodeList                 ' xmlNode for message box
Public xmlNodeCaption As MSXML.IXMLDOMNode                         ' xmlNode for command button caption
Public xmlNodeListError As MSXML.IXMLDOMNodeList                        ' xmlNode for error define

Public strFinYear As String                                         ' Current finacial year (4 digit number)
Public strMonth As String                                           ' Current month (2 digit number)
Public str3Months As String                                         ' Current 3 months (1 digit number)
Public strFirstDayOfQuarter As String                                           ' Current first day of quarter  (8 digit number)
Public strLastDayOfQuarter As String
Public strFinanceStartDate As String

Public isUsePIT As Boolean ' su dung PIT
Public isQCT As Boolean ' -->QCT
Public strMaVach As String ' -->version Ma vach
