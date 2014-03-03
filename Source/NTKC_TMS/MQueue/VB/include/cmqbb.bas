Attribute VB_Name = "CMQBB"
'**********************************************************************'
'*                                                                    *'
'*                  WebSphere MQ for Windows                          *'
'*                                                                    *'
'*  FILE NAME:      CMQBB                                             *'
'*                                                                    *'
'*  DESCRIPTION:    Declarations for MQ Administration Interface      *'
'*                                                                    *'
'**********************************************************************'
'*  @N_OCO_COPYRIGHT@                                                 *'
'*  Licensed Materials - Property of IBM                              *'
'*                                                                    *'
'*  63H9336                                                           *'
'*  (c) Copyright IBM Corp. 1999, 2005 All Rights Reserved.           *'
'*                                                                    *'
'*  US Government Users Restricted Rights - Use, duplication or       *'
'*  disclosure restricted by GSA ADP Schedule Contract with           *'
'*  IBM Corp.                                                         *'
'*  @NOC_COPYRIGHT@                                                   *'
'**********************************************************************'
'*                                                                    *'
'*  FUNCTION:       This file declares the functions and named        *'
'*                  constants for the MQ administration interface     *'
'*                  (MQAI).                                           *'
'*                                                                    *'
'*  PROCESSOR:      BASIC                                             *'
'*                                                                    *'
'**********************************************************************'

'$$#con$ on'
'********************************************************************'
'*  Values Related to Specific Functions                            *'
'********************************************************************'

'Create-Bag Options for mqCreateBag'
Global Const MQCBO_NONE = &H0
Global Const MQCBO_USER_BAG = &H0
Global Const MQCBO_ADMIN_BAG = &H1
Global Const MQCBO_COMMAND_BAG = &H10
Global Const MQCBO_SYSTEM_BAG = &H20
Global Const MQCBO_GROUP_BAG = &H40
Global Const MQCBO_LIST_FORM_ALLOWED = &H2
Global Const MQCBO_LIST_FORM_INHIBITED = &H0
Global Const MQCBO_REORDER_AS_REQUIRED = &H4
Global Const MQCBO_DO_NOT_REORDER = &H0
Global Const MQCBO_CHECK_SELECTORS = &H8
Global Const MQCBO_DO_NOT_CHECK_SELECTORS = &H0

'Buffer Length for mqAddString and mqSetString'
Global Const MQBL_NULL_TERMINATED = -1

'Item Type for mqInquireItemInfo'
Global Const MQITEM_INTEGER = 1
Global Const MQITEM_STRING = 2
Global Const MQITEM_BAG = 3
Global Const MQITEM_BYTE_STRING = 4
Global Const MQITEM_INTEGER_FILTER = 5
Global Const MQITEM_STRING_FILTER = 6
Global Const MQITEM_INTEGER64 = 7
Global Const MQITEM_BYTE_STRING_FILTER = 8
Global Const MQIT_INTEGER = 1
Global Const MQIT_STRING = 2
Global Const MQIT_BAG = 3

'********************************************************************'
'*  Values Related to Most Functions                                *'
'********************************************************************'

'Integer Selectors for Object Attributes'
'See MQIA_*   values in CMQC'
'See MQIACF_* values in CMQCFC'
'See MQIACH_* values in CMQCFC'
'See MQIAMO_*   values in CMQCFC'
'See MQIAMO64_* values in CMQCFC'
'Character Selectors for Object Attributes'
'See MQCA_*   values in CMQC'
'See MQCACF_* values in CMQCFC'
'See MQCACH_* values in CMQCFC'
'See MQCAMO_* values in CMQCFC'
'Byte String Selectors for Object Attributes'
'See MQBA_*   values in CMQC'
'See MQBACF_* values in CMQCFC'
'Group Selectors for Object Attributes'
'See MQGACF_* values in CMQCFC'
'Handle Selectors'
Global Const MQHA_FIRST = 4001
Global Const MQHA_BAG_HANDLE = 4001
Global Const MQHA_LAST_USED = 4001
Global Const MQHA_LAST = 6000

'Limits for Selectors for Object Attributes'
Global Const MQOA_FIRST = 1
Global Const MQOA_LAST = 9000

'Integer System Selectors'
Global Const MQIASY_FIRST = -1
Global Const MQIASY_CODED_CHAR_SET_ID = -1
Global Const MQIASY_TYPE = -2
Global Const MQIASY_COMMAND = -3
Global Const MQIASY_MSG_SEQ_NUMBER = -4
Global Const MQIASY_CONTROL = -5
Global Const MQIASY_COMP_CODE = -6
Global Const MQIASY_REASON = -7
Global Const MQIASY_BAG_OPTIONS = -8
Global Const MQIASY_VERSION = -9
Global Const MQIASY_LAST_USED = -9
Global Const MQIASY_LAST = -2000

'Special Selector Values'
Global Const MQSEL_ANY_SELECTOR = -30001
Global Const MQSEL_ANY_USER_SELECTOR = -30002
Global Const MQSEL_ANY_SYSTEM_SELECTOR = -30003
Global Const MQSEL_ALL_SELECTORS = -30001
Global Const MQSEL_ALL_USER_SELECTORS = -30002
Global Const MQSEL_ALL_SYSTEM_SELECTORS = -30003

'Special Index Values'
Global Const MQIND_NONE = -1
Global Const MQIND_ALL = -2

'Bag Handles'
Global Const MQHB_UNUSABLE_HBAG = -1
Global Const MQHB_NONE = -2

'********************************************************************'
'*  mqAddBag Function -- Add Nested Bag to Bag                      *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Global Const MQVBDLL = "MQM.DLL"

Declare Sub mqAddBag Lib "MQM.DLL" Alias "mqAddBagstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Global Const MQVBDLL = "MQIC.DLL"

Declare Sub mqAddBag Lib "MQIC.DLL" Alias "mqAddBagstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Global Const MQVBDLL = "MQICXA.DLL"

Declare Sub mqAddBag Lib "MQICXA.DLL" Alias "mqAddBagstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
Global Const MQVBDLL = "NONE"
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddByteString Function -- Add Byte String to Bag              *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddByteString Lib "MQM.DLL" Alias "mqAddByteStringstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddByteString Lib "MQIC.DLL" Alias "mqAddByteStringstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddByteString Lib "MQICXA.DLL" Alias "mqAddByteStringstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddByteStringFilter Function -- Add Byte String Filter to Bag *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddByteStringFilter Lib "MQM.DLL" Alias "mqAddByteStringFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddByteStringFilter Lib "MQIC.DLL" Alias "mqAddByteStringFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddByteStringFilter Lib "MQICXA.DLL" Alias "mqAddByteStringFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddInquiry Function -- Add an Inquiry Item to Bag             *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddInquiry Lib "MQM.DLL" Alias "mqAddInquirystd@16" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddInquiry Lib "MQIC.DLL" Alias "mqAddInquirystd@16" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddInquiry Lib "MQICXA.DLL" Alias "mqAddInquirystd@16" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddInteger Function -- Add Integer to Bag                     *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddInteger Lib "MQM.DLL" Alias "mqAddIntegerstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddInteger Lib "MQIC.DLL" Alias "mqAddIntegerstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddInteger Lib "MQICXA.DLL" Alias "mqAddIntegerstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddIntegerFilter Function -- Add Integer Filter to Bag        *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddIntegerFilter Lib "MQM.DLL" Alias "mqAddIntegerFilterstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddIntegerFilter Lib "MQIC.DLL" Alias "mqAddIntegerFilterstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddIntegerFilter Lib "MQICXA.DLL" Alias "mqAddIntegerFilterstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemValue As Long, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddString Function -- Add String to Bag                       *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddString Lib "MQM.DLL" Alias "mqAddStringstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddString Lib "MQIC.DLL" Alias "mqAddStringstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddString Lib "MQICXA.DLL" Alias "mqAddStringstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqAddStringFilter Function -- Add String Filter to Bag          *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqAddStringFilter Lib "MQM.DLL" Alias "mqAddStringFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqAddStringFilter Lib "MQIC.DLL" Alias "mqAddStringFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqAddStringFilter Lib "MQICXA.DLL" Alias "mqAddStringFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqBagToBuffer Function -- Convert Bag to PCF                    *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqBagToBuffer Lib "MQM.DLL" Alias "mqBagToBufferstd@28" _
 (ByVal OptionsBag As Long, _
  ByVal DataBag As Long, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqBagToBuffer Lib "MQIC.DLL" Alias "mqBagToBufferstd@28" _
 (ByVal OptionsBag As Long, _
  ByVal DataBag As Long, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqBagToBuffer Lib "MQICXA.DLL" Alias "mqBagToBufferstd@28" _
 (ByVal OptionsBag As Long, _
  ByVal DataBag As Long, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqBufferToBag Function -- Convert PCF to Bag                    *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqBufferToBag Lib "MQM.DLL" Alias "mqBufferToBagstd@24" _
 (ByVal OptionsBag As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Any, _
  DataBag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqBufferToBag Lib "MQIC.DLL" Alias "mqBufferToBagstd@24" _
 (ByVal OptionsBag As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Any, _
  DataBag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqBufferToBag Lib "MQICXA.DLL" Alias "mqBufferToBagstd@24" _
 (ByVal OptionsBag As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Any, _
  DataBag As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqClearBag Function -- Delete All Items in Bag                  *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqClearBag Lib "MQM.DLL" Alias "mqClearBagstd@12" _
 (ByVal Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqClearBag Lib "MQIC.DLL" Alias "mqClearBagstd@12" _
 (ByVal Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqClearBag Lib "MQICXA.DLL" Alias "mqClearBagstd@12" _
 (ByVal Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqCountItems Function -- Count Items in Bag                     *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqCountItems Lib "MQM.DLL" Alias "mqCountItemsstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ItemCount As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqCountItems Lib "MQIC.DLL" Alias "mqCountItemsstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ItemCount As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqCountItems Lib "MQICXA.DLL" Alias "mqCountItemsstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ItemCount As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqCreateBag Function -- Create Bag                              *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqCreateBag Lib "MQM.DLL" Alias "mqCreateBagstd@16" _
 (ByVal Options As Long, _
  Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqCreateBag Lib "MQIC.DLL" Alias "mqCreateBagstd@16" _
 (ByVal Options As Long, _
  Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqCreateBag Lib "MQICXA.DLL" Alias "mqCreateBagstd@16" _
 (ByVal Options As Long, _
  Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqDeleteBag Function -- Delete Bag                              *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqDeleteBag Lib "MQM.DLL" Alias "mqDeleteBagstd@12" _
 (Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqDeleteBag Lib "MQIC.DLL" Alias "mqDeleteBagstd@12" _
 (Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqDeleteBag Lib "MQICXA.DLL" Alias "mqDeleteBagstd@12" _
 (Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqDeleteItem Function -- Delete Item in Bag                     *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqDeleteItem Lib "MQM.DLL" Alias "mqDeleteItemstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqDeleteItem Lib "MQIC.DLL" Alias "mqDeleteItemstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqDeleteItem Lib "MQICXA.DLL" Alias "mqDeleteItemstd@20" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqExecute Function -- Send Admin Command and Receive Reponse    *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqExecute Lib "MQM.DLL" Alias "mqExecutestd@36" _
 (ByVal Hconn As Long, _
  ByVal Command As Long, _
  ByVal OptionsBag As Long, _
  ByVal AdminBag As Long, _
  ByVal ResponseBag As Long, _
  ByVal AdminQ As Long, _
  ByVal ResponseQ As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqExecute Lib "MQIC.DLL" Alias "mqExecutestd@36" _
 (ByVal Hconn As Long, _
  ByVal Command As Long, _
  ByVal OptionsBag As Long, _
  ByVal AdminBag As Long, _
  ByVal ResponseBag As Long, _
  ByVal AdminQ As Long, _
  ByVal ResponseQ As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqExecute Lib "MQICXA.DLL" Alias "mqExecutestd@36" _
 (ByVal Hconn As Long, _
  ByVal Command As Long, _
  ByVal OptionsBag As Long, _
  ByVal AdminBag As Long, _
  ByVal ResponseBag As Long, _
  ByVal AdminQ As Long, _
  ByVal ResponseQ As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqGetBag Function -- Receive PCF Message into Bag               *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqGetBag Lib "MQM.DLL" Alias "mqGetBagstd@28" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As Any, _
  GetMsgOpts As Any, _
  Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqGetBag Lib "MQIC.DLL" Alias "mqGetBagstd@28" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As Any, _
  GetMsgOpts As Any, _
  Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqGetBag Lib "MQICXA.DLL" Alias "mqGetBagstd@28" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As Any, _
  GetMsgOpts As Any, _
  Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqInquireBag Function -- Inquire Handle in Bag                  *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireBag Lib "MQM.DLL" Alias "mqInquireBagstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireBag Lib "MQIC.DLL" Alias "mqInquireBagstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireBag Lib "MQICXA.DLL" Alias "mqInquireBagstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'*********************************************************************'
'*  mqInquireByteString Function -- Inquire Byte String in Bag       *'
'*********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireByteString Lib "MQM.DLL" Alias "mqInquireByteStringstd@32" _
 (Bag As Long, _
  Selector As Long, _
  ItemIndex As Long, _
  BufferLength As Long, _
  Buffer As Byte, _
  ByteStringLength As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireByteString Lib "MQIC.DLL" Alias "mqInquireByteStringstd@32" _
 (Bag As Long, _
  Selector As Long, _
  ItemIndex As Long, _
  BufferLength As Long, _
  Buffer As Byte, _
  ByteStringLength As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireByteString Lib "MQICXA.DLL" Alias "mqInquireByteStringstd@32" _
 (Bag As Long, _
  Selector As Long, _
  ItemIndex As Long, _
  BufferLength As Long, _
  Buffer As Byte, _
  ByteStringLength As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'*********************************************************************'
'*  mqInquireByteStringFilter Function --                            *'
'*                                Inquire Byte String Filter in Bag  *'
'*********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireByteStringFilter Lib "MQM.DLL" Alias "mqInquireByteStringFilterstd@36" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  Buffer As Byte, _
  ByteStringLength As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireByteStringFilter Lib "MQIC.DLL" Alias "mqInquireByteStringFilterstd@36" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  Buffer As Byte, _
  ByteStringLength As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireByteStringFilter Lib "MQICXA.DLL" Alias "mqInquireByteStringFilterstd@36" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  Buffer As Byte, _
  ByteStringLength As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqInquireInteger Function -- Inquire Integer in Bag             *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireInteger Lib "MQM.DLL" Alias "mqInquireIntegerstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireInteger Lib "MQIC.DLL" Alias "mqInquireIntegerstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireInteger Lib "MQICXA.DLL" Alias "mqInquireIntegerstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'* mqInquireIntegerFilter Function -- Inquire Integer Filter in Bag *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireIntegerFilter Lib "MQM.DLL" Alias "mqInquireIntegerFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireIntegerFilter Lib "MQIC.DLL" Alias "mqInquireIntegerFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireIntegerFilter Lib "MQICXA.DLL" Alias "mqInquireIntegerFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ItemValue As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqInquireItemInfo Function -- Inquire Attributes of Item in Bag *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireItemInfo Lib "MQM.DLL" Alias "mqInquireItemInfostd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  OutSelector As Long, _
  ItemType As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireItemInfo Lib "MQIC.DLL" Alias "mqInquireItemInfostd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  OutSelector As Long, _
  ItemType As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireItemInfo Lib "MQICXA.DLL" Alias "mqInquireItemInfostd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  OutSelector As Long, _
  ItemType As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqInquireString Function -- Inquire String in Bag               *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireString Lib "MQM.DLL" Alias "mqInquireStringstd@36" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  StringLength As Long, _
  CodedCharSetId As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireString Lib "MQIC.DLL" Alias "mqInquireStringstd@36" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  StringLength As Long, _
  CodedCharSetId As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireString Lib "MQICXA.DLL" Alias "mqInquireStringstd@36" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  StringLength As Long, _
  CodedCharSetId As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqInquireStringFilter Function -- Inquire String Filter in Bag  *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqInquireStringFilter Lib "MQM.DLL" Alias "mqInquireStringFilterstd@40" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  StringLength As Long, _
  CodedCharSetId As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqInquireStringFilter Lib "MQIC.DLL" Alias "mqInquireStringFilterstd@40" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  StringLength As Long, _
  CodedCharSetId As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqInquireStringFilter Lib "MQICXA.DLL" Alias "mqInquireStringFilterstd@40" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  StringLength As Long, _
  CodedCharSetId As Long, _
  Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqPutBag Function -- Send Bag as PCF Message                    *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqPutBag Lib "MQM.DLL" Alias "mqPutBagstd@28" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As Any, _
  PutMsgOpts As Any, _
  ByVal Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqPutBag Lib "MQIC.DLL" Alias "mqPutBagstd@28" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As Any, _
  PutMsgOpts As Any, _
  ByVal Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqPutBag Lib "MQICXA.DLL" Alias "mqPutBagstd@28" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As Any, _
  PutMsgOpts As Any, _
  ByVal Bag As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqSetByteString Function -- Modify Byte String in Bag           *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqSetByteString Lib "MQM.DLL" Alias "mqSetByteStringstd@28" _
 (Bag As Long, _
  Selector As Long, _
  ItemIndex As Long, _
  BufferLength As Long, _
  Buffer As Byte, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqSetByteString Lib "MQIC.DLL" Alias "mqSetByteStringstd@28" _
 (Bag As Long, _
  Selector As Long, _
  ItemIndex As Long, _
  BufferLength As Long, _
  Buffer As Byte, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqSetByteString Lib "MQICXA.DLL" Alias "mqSetByteStringstd@28" _
 (Bag As Long, _
  Selector As Long, _
  ItemIndex As Long, _
  BufferLength As Long, _
  Buffer As Byte, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqSetByteStringFilter Function --                               *'
'*                           Modify Byte String Filter in Bag       *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqSetByteStringFilter Lib "MQM.DLL" Alias "mqSetByteStringFilterstd@32" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqSetByteStringFilter Lib "MQIC.DLL" Alias "mqSetByteStringFilterstd@32" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqSetByteStringFilter Lib "MQICXA.DLL" Alias "mqSetByteStringFilterstd@32" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As Byte, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqSetInteger Function -- Modify Integer in Bag                  *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqSetInteger Lib "MQM.DLL" Alias "mqSetIntegerstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqSetInteger Lib "MQIC.DLL" Alias "mqSetIntegerstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqSetInteger Lib "MQICXA.DLL" Alias "mqSetIntegerstd@24" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal ItemValue As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqSetIntegerFilter Function -- Modify Integer Filter in Bag     *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqSetIntegerFilter Lib "MQM.DLL" Alias "mqSetIntegerFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal ItemValue As Long, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqSetIntegerFilter Lib "MQIC.DLL" Alias "mqSetIntegerFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal ItemValue As Long, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqSetIntegerFilter Lib "MQICXA.DLL" Alias "mqSetIntegerFilterstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal ItemValue As Long, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqSetString Function -- Modify String in Bag                    *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqSetString Lib "MQM.DLL" Alias "mqSetStringstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqSetString Lib "MQIC.DLL" Alias "mqSetStringstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqSetString Lib "MQICXA.DLL" Alias "mqSetStringstd@28" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'********************************************************************'
'*  mqSetStringFilter Function -- Modify String Filter in Bag       *'
'********************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub mqSetStringFilter Lib "MQM.DLL" Alias "mqSetStringFilterstd@32" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub mqSetStringFilter Lib "MQIC.DLL" Alias "mqSetStringFilterstd@32" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub mqSetStringFilter Lib "MQICXA.DLL" Alias "mqSetStringFilterstd@32" _
 (ByVal Bag As Long, _
  ByVal Selector As Long, _
  ByVal ItemIndex As Long, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  ByVal Operator As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If
'****************************************************************'
'****************************************************************'

'****************************************************************'
'*  End of CMQBB                                                *'
'****************************************************************'

