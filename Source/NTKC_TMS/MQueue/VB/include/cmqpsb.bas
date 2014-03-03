Attribute VB_Name = "CMQPSB"
'**********************************************************************'
'*                                                                    *'
'*                  WebSphere MQ for Windows                          *'
'*                                                                    *'
'*  FILE NAME:      CMQPSB                                            *'
'*                                                                    *'
'*  DESCRIPTION:    Declarations for Publish/Subscribe                *'
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
'*  FUNCTION:       This file declares the named constants for        *'
'*                  publish/subscribe.                                *'
'*                                                                    *'
'*  PROCESSOR:      BASIC                                             *'
'*                                                                    *'
'**********************************************************************'

'$$#con$ on'
'****************************************************************'
'*  Publish/Subscribe Tags                                      *'
'****************************************************************'
'Tags as strings'
Global Const MQPS_COMMAND = "MQPSCommand"
Global Const MQPS_COMP_CODE = "MQPSCompCode"
Global Const MQPS_CORREL_ID = "MQPSCorrelId"
Global Const MQPS_DELETE_OPTIONS = "MQPSDelOpts"
Global Const MQPS_ERROR_ID = "MQPSErrorId"
Global Const MQPS_ERROR_POS = "MQPSErrorPos"
Global Const MQPS_INTEGER_DATA = "MQPSIntData"
Global Const MQPS_PARAMETER_ID = "MQPSParmId"
Global Const MQPS_PUBLICATION_OPTIONS = "MQPSPubOpts"
Global Const MQPS_PUBLISH_TIMESTAMP = "MQPSPubTime"
Global Const MQPS_Q_MGR_NAME = "MQPSQMgrName"
Global Const MQPS_Q_NAME = "MQPSQName"
Global Const MQPS_REASON = "MQPSReason"
Global Const MQPS_REASON_TEXT = "MQPSReasonText"
Global Const MQPS_REGISTRATION_OPTIONS = "MQPSRegOpts"
Global Const MQPS_SEQUENCE_NUMBER = "MQPSSeqNum"
Global Const MQPS_STREAM_NAME = "MQPSStreamName"
Global Const MQPS_STRING_DATA = "MQPSStringData"
Global Const MQPS_SUBSCRIPTION_IDENTITY = "MQPSSubIdentity"
Global Const MQPS_SUBSCRIPTION_NAME = "MQPSSubName"
Global Const MQPS_SUBSCRIPTION_USER_DATA = "MQPSSubUserData"
Global Const MQPS_TOPIC = "MQPSTopic"
Global Const MQPS_USER_ID = "MQPSUserId"

'Tags as blank-enclosed strings'
Global Const MQPS_COMMAND_B = " MQPSCommand "
Global Const MQPS_COMP_CODE_B = " MQPSCompCode "
Global Const MQPS_CORREL_ID_B = " MQPSCorrelId "
Global Const MQPS_DELETE_OPTIONS_B = " MQPSDelOpts "
Global Const MQPS_ERROR_ID_B = " MQPSErrorId "
Global Const MQPS_ERROR_POS_B = " MQPSErrorPos "
Global Const MQPS_INTEGER_DATA_B = " MQPSIntData "
Global Const MQPS_PARAMETER_ID_B = " MQPSParmId "
Global Const MQPS_PUBLICATION_OPTIONS_B = " MQPSPubOpts "
Global Const MQPS_PUBLISH_TIMESTAMP_B = " MQPSPubTime "
Global Const MQPS_Q_MGR_NAME_B = " MQPSQMgrName "
Global Const MQPS_Q_NAME_B = " MQPSQName "
Global Const MQPS_REASON_B = " MQPSReason "
Global Const MQPS_REASON_TEXT_B = " MQPSReasonText "
Global Const MQPS_REGISTRATION_OPTIONS_B = " MQPSRegOpts "
Global Const MQPS_SEQUENCE_NUMBER_B = " MQPSSeqNum "
Global Const MQPS_STREAM_NAME_B = " MQPSStreamName "
Global Const MQPS_STRING_DATA_B = " MQPSStringData "
Global Const MQPS_SUBSCRIPTION_IDENTITY_B = " MQPSSubIdentity "
Global Const MQPS_SUBSCRIPTION_NAME_B = " MQPSSubName "
Global Const MQPS_SUBSCRIPTION_USER_DATA_B = " MQPSSubUserData "
Global Const MQPS_TOPIC_B = " MQPSTopic "
Global Const MQPS_USER_ID_B = " MQPSUserId "

'****************************************************************'
'*  Values for MQPS_COMMAND Tag                                 *'
'****************************************************************'
'Values as strings'
Global Const MQPS_DELETE_PUBLICATION = "DeletePub"
Global Const MQPS_DEREGISTER_PUBLISHER = "DeregPub"
Global Const MQPS_DEREGISTER_SUBSCRIBER = "DeregSub"
Global Const MQPS_PUBLISH = "Publish"
Global Const MQPS_REGISTER_PUBLISHER = "RegPub"
Global Const MQPS_REGISTER_SUBSCRIBER = "RegSub"
Global Const MQPS_REQUEST_UPDATE = "ReqUpdate"

'Values as blank-enclosed strings'
Global Const MQPS_DELETE_PUBLICATION_B = " DeletePub "
Global Const MQPS_DEREGISTER_PUBLISHER_B = " DeregPub "
Global Const MQPS_DEREGISTER_SUBSCRIBER_B = " DeregSub "
Global Const MQPS_PUBLISH_B = " Publish "
Global Const MQPS_REGISTER_PUBLISHER_B = " RegPub "
Global Const MQPS_REGISTER_SUBSCRIBER_B = " RegSub "
Global Const MQPS_REQUEST_UPDATE_B = " ReqUpdate "

'****************************************************************'
'*  Values for following tags:                                  *'
'*    MQPS_DELETE_OPTIONS                                       *'
'*    MQPS_PUBLICATION_OPTIONS                                  *'
'*    MQPS_REGISTRATION_OPTIONS                                 *'
'****************************************************************'
'Values as strings'
Global Const MQPS_ADD_NAME = "AddName"
Global Const MQPS_ANONYMOUS = "Anon"
Global Const MQPS_CORREL_ID_AS_IDENTITY = "CorrelAsId"
Global Const MQPS_DEREGISTER_ALL = "DeregAll"
Global Const MQPS_DIRECT_REQUESTS = "DirectReq"
Global Const MQPS_DUPLICATES_OK = "DupsOK"
Global Const MQPS_FULL_RESPONSE = "FullResp"
Global Const MQPS_INCLUDE_STREAM_NAME = "InclStreamName"
Global Const MQPS_INFORM_IF_RETAINED = "InformIfRet"
Global Const MQPS_IS_RETAINED_PUBLICATION = "IsRetainedPub"
Global Const MQPS_JOIN_EXCLUSIVE = "JoinExcl"
Global Const MQPS_JOIN_SHARED = "JoinShared"
Global Const MQPS_LEAVE_ONLY = "LeaveOnly"
Global Const MQPS_LOCAL = "Local"
Global Const MQPS_LOCKED = "Locked"
Global Const MQPS_NEW_PUBLICATIONS_ONLY = "NewPubsOnly"
Global Const MQPS_NO_ALTERATION = "NoAlter"
Global Const MQPS_NO_REGISTRATION = "NoReg"
Global Const MQPS_NON_PERSISTENT = "NonPers"
Global Const MQPS_NONE = "None"
Global Const MQPS_OTHER_SUBSCRIBERS_ONLY = "OtherSubsOnly"
Global Const MQPS_PERSISTENT = "Pers"
Global Const MQPS_PERSISTENT_AS_PUBLISH = "PersAsPub"
Global Const MQPS_PERSISTENT_AS_Q = "PersAsQueue"
Global Const MQPS_PUBLISH_ON_REQUEST_ONLY = "PubOnReqOnly"
Global Const MQPS_RETAIN_PUBLICATION = "RetainPub"
Global Const MQPS_VARIABLE_USER_ID = "VariableUserId"

'Values as blank-enclosed strings'
Global Const MQPS_ADD_NAME_B = " AddName "
Global Const MQPS_ANONYMOUS_B = " Anon "
Global Const MQPS_CORREL_ID_AS_IDENTITY_B = " CorrelAsId "
Global Const MQPS_DEREGISTER_ALL_B = " DeregAll "
Global Const MQPS_DIRECT_REQUESTS_B = " DirectReq "
Global Const MQPS_DUPLICATES_OK_B = " DupsOK "
Global Const MQPS_FULL_RESPONSE_B = " FullResp "
Global Const MQPS_INCLUDE_STREAM_NAME_B = " InclStreamName "
Global Const MQPS_INFORM_IF_RETAINED_B = " InformIfRet "
Global Const MQPS_IS_RETAINED_PUBLICATION_B = " IsRetainedPub "
Global Const MQPS_JOIN_EXCLUSIVE_B = " JoinExcl "
Global Const MQPS_JOIN_SHARED_B = " JoinShared "
Global Const MQPS_LEAVE_ONLY_B = " LeaveOnly "
Global Const MQPS_LOCAL_B = " Local "
Global Const MQPS_LOCKED_B = " Locked "
Global Const MQPS_NEW_PUBLICATIONS_ONLY_B = " NewPubsOnly "
Global Const MQPS_NO_ALTERATION_B = " NoAlter "
Global Const MQPS_NO_REGISTRATION_B = " NoReg "
Global Const MQPS_NON_PERSISTENT_B = " NonPers "
Global Const MQPS_NONE_B = " None "
Global Const MQPS_OTHER_SUBSCRIBERS_ONLY_B = " OtherSubsOnly "
Global Const MQPS_PERSISTENT_B = " Pers "
Global Const MQPS_PERSISTENT_AS_PUBLISH_B = " PersAsPub "
Global Const MQPS_PERSISTENT_AS_Q_B = " PersAsQueue "
Global Const MQPS_PUBLISH_ON_REQUEST_ONLY_B = " PubOnReqOnly "
Global Const MQPS_RETAIN_PUBLICATION_B = " RetainPub "
Global Const MQPS_VARIABLE_USER_ID_B = " VariableUserId "
      '****************************************************************'
      '*  End of CMQPSB                                               *'
      '****************************************************************'

'****************************************************************'
'*  End of CMQPSB                                               *'
'****************************************************************'

