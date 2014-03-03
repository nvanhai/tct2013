Attribute VB_Name = "CMQXB"
'**********************************************************************'
'*                                                                    *'
'*                  WebSphere MQ for Windows                          *'
'*                                                                    *'
'*  FILE NAME:      CMQXB                                             *'
'*                                                                    *'
'*  DESCRIPTION:    Structures and Constants for MQCD and MQCNOCD     *'
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
'*  FUNCTION:       This file declares the structures and             *'
'*                  named constants for MQCD and MQCNOCD.             *'
'*                                                                    *'
'*  PROCESSOR:      BASIC                                             *'
'*                                                                    *'
'**********************************************************************'

'****************************************************************'
'*  Values Related to MQACH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQACH_STRUC_ID = "ACH "

'Structure Version Number'
Global Const MQACH_VERSION_1 = 1
Global Const MQACH_CURRENT_VERSION = 1

'Structure Length'
Global Const MQACH_LENGTH_1 = 68
Global Const MQACH_CURRENT_LENGTH = 68

'****************************************************************'
'*  Values Related to MQAXC Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQAXC_STRUC_ID = "AXC "

'Structure Version Number'
Global Const MQAXC_VERSION_1 = 1
Global Const MQAXC_CURRENT_VERSION = 1

'Environments'
Global Const MQXE_OTHER = 0
Global Const MQXE_MCA = 1
Global Const MQXE_MCA_SVRCONN = 2
Global Const MQXE_COMMAND_SERVER = 3
Global Const MQXE_MQSC = 4

'****************************************************************'
'*  Values Related to MQAXP Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQAXP_STRUC_ID = "AXP "

'Structure Version Number'
Global Const MQAXP_VERSION_1 = 1
Global Const MQAXP_CURRENT_VERSION = 1

'API Caller Types'
Global Const MQXACT_EXTERNAL = 1
Global Const MQXACT_INTERNAL = 2

'Problem Determination Area'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQXPDA_NONE As MQBYTE48

'API Function Identifiers'
Global Const MQXF_INIT = 1
Global Const MQXF_TERM = 2
Global Const MQXF_CONN = 3
Global Const MQXF_CONNX = 4
Global Const MQXF_DISC = 5
Global Const MQXF_OPEN = 6
Global Const MQXF_CLOSE = 7
Global Const MQXF_PUT1 = 8
Global Const MQXF_PUT = 9
Global Const MQXF_GET = 10
Global Const MQXF_DATA_CONV_ON_GET = 11
Global Const MQXF_INQ = 12
Global Const MQXF_SET = 13
Global Const MQXF_BEGIN = 14
Global Const MQXF_CMIT = 15
Global Const MQXF_BACK = 16

'****************************************************************'
'*  Values Related to MQCD Structure                            *'
'****************************************************************'
'Structure Version Number'
Global Const MQCD_VERSION_1 = 1
Global Const MQCD_VERSION_2 = 2
Global Const MQCD_VERSION_3 = 3
Global Const MQCD_VERSION_4 = 4
Global Const MQCD_VERSION_5 = 5
Global Const MQCD_VERSION_6 = 6
Global Const MQCD_VERSION_7 = 7
Global Const MQCD_VERSION_8 = 8
Global Const MQCD_CURRENT_VERSION = 8

'Structure Length'
Global Const MQCD_LENGTH_4 = 1540
Global Const MQCD_LENGTH_5 = 1552
Global Const MQCD_LENGTH_6 = 1648
Global Const MQCD_LENGTH_7 = 1748
Global Const MQCD_LENGTH_8 = 1840
Global Const MQCD_CURRENT_LENGTH = 1840

'Channel Types'
Global Const MQCHT_SENDER = 1
Global Const MQCHT_SERVER = 2
Global Const MQCHT_RECEIVER = 3
Global Const MQCHT_REQUESTER = 4
Global Const MQCHT_ALL = 5
Global Const MQCHT_CLNTCONN = 6
Global Const MQCHT_SVRCONN = 7
Global Const MQCHT_CLUSRCVR = 8
Global Const MQCHT_CLUSSDR = 9

'Channel Compression'
Global Const MQCOMPRESS_NOT_AVAILABLE = -1
Global Const MQCOMPRESS_NONE = 0
Global Const MQCOMPRESS_RLE = 1
Global Const MQCOMPRESS_ZLIBFAST = 2
Global Const MQCOMPRESS_ZLIBHIGH = 4
Global Const MQCOMPRESS_SYSTEM = 8
Global Const MQCOMPRESS_ANY = &HFFFFFFF

'Transport Types'
Global Const MQXPT_ALL = -1
Global Const MQXPT_LOCAL = 0
Global Const MQXPT_LU62 = 1
Global Const MQXPT_TCP = 2
Global Const MQXPT_NETBIOS = 3
Global Const MQXPT_SPX = 4
Global Const MQXPT_DECNET = 5
Global Const MQXPT_UDP = 6

'Put Authority'
Global Const MQPA_DEFAULT = 1
Global Const MQPA_CONTEXT = 2
Global Const MQPA_ONLY_MCA = 3
Global Const MQPA_ALTERNATE_OR_MCA = 4

'Channel Data Conversion'
Global Const MQCDC_SENDER_CONVERSION = 1
Global Const MQCDC_NO_SENDER_CONVERSION = 0

'MCA Types'
Global Const MQMCAT_PROCESS = 1
Global Const MQMCAT_THREAD = 2

'NonPersistent-Message Speeds'
Global Const MQNPMS_NORMAL = 1
Global Const MQNPMS_FAST = 2

'SSL Client Authentication'
Global Const MQSCA_REQUIRED = 0
Global Const MQSCA_OPTIONAL = 1

'KeepAlive Interval'
Global Const MQKAI_AUTO = -1

'****************************************************************'
'*  Values Related to MQCXP Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQCXP_STRUC_ID = "CXP "

'Structure Version Number'
Global Const MQCXP_VERSION_1 = 1
Global Const MQCXP_VERSION_2 = 2
Global Const MQCXP_VERSION_3 = 3
Global Const MQCXP_VERSION_4 = 4
Global Const MQCXP_VERSION_5 = 5
Global Const MQCXP_VERSION_6 = 6
Global Const MQCXP_CURRENT_VERSION = 6

'Exit Response 2'
Global Const MQXR2_PUT_WITH_DEF_ACTION = 0
Global Const MQXR2_PUT_WITH_DEF_USERID = 1
Global Const MQXR2_PUT_WITH_MSG_USERID = 2
Global Const MQXR2_USE_AGENT_BUFFER = 0
Global Const MQXR2_USE_EXIT_BUFFER = 4
Global Const MQXR2_DEFAULT_CONTINUATION = 0
Global Const MQXR2_CONTINUE_CHAIN = 8
Global Const MQXR2_SUPPRESS_CHAIN = 16
Global Const MQXR2_STATIC_CACHE = 0
Global Const MQXR2_DYNAMIC_CACHE = 32

'Capability Flags'
Global Const MQCF_NONE = &H0
Global Const MQCF_DIST_LISTS = &H1

'****************************************************************'
'*  Values Related to MQDXP Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQDXP_STRUC_ID = "DXP "

'Structure Version Number'
Global Const MQDXP_VERSION_1 = 1
Global Const MQDXP_CURRENT_VERSION = 1

'Exit Response'
Global Const MQXDR_OK = 0
Global Const MQXDR_CONVERSION_FAILED = 1

'****************************************************************'
'*  Values Related to MQPXP Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQPXP_STRUC_ID = "PXP "

'Structure Version Number'
Global Const MQPXP_VERSION_1 = 1
Global Const MQPXP_CURRENT_VERSION = 1

'Destination Types'
Global Const MQDT_APPL = 1
Global Const MQDT_BROKER = 2

'****************************************************************'
'*  Values Related to MQWDR Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQWDR_STRUC_ID = "WDR "

'Structure Version Number'
Global Const MQWDR_VERSION_1 = 1
Global Const MQWDR_VERSION_2 = 2
Global Const MQWDR_CURRENT_VERSION = 2

'Structure Length'
Global Const MQWDR_LENGTH_1 = 124
Global Const MQWDR_LENGTH_2 = 136
Global Const MQWDR_CURRENT_LENGTH = 136

'Queue Manager Flags'
Global Const MQQMF_REPOSITORY_Q_MGR = &H2
Global Const MQQMF_CLUSSDR_USER_DEFINED = &H8
Global Const MQQMF_CLUSSDR_AUTO_DEFINED = &H10
Global Const MQQMF_AVAILABLE = &H20

'****************************************************************'
'*  Values Related to MQWQR Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQWQR_STRUC_ID = "WQR "

'Structure Version Number'
Global Const MQWQR_VERSION_1 = 1
Global Const MQWQR_VERSION_2 = 2
Global Const MQWQR_CURRENT_VERSION = 2

'Structure Length'
Global Const MQWQR_LENGTH_1 = 200
Global Const MQWQR_LENGTH_2 = 208
Global Const MQWQR_CURRENT_LENGTH = 208

'Queue Flags'
Global Const MQQF_LOCAL_Q = &H1
Global Const MQQF_CLWL_USEQ_ANY = &H40
Global Const MQQF_CLWL_USEQ_LOCAL = &H80

'****************************************************************'
'*  Values Related to MQWXP Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQWXP_STRUC_ID = "WXP "

'Structure Version Number'
Global Const MQWXP_VERSION_1 = 1
Global Const MQWXP_VERSION_2 = 2
Global Const MQWXP_VERSION_3 = 3
Global Const MQWXP_CURRENT_VERSION = 3

'Cluster Workload Flags'
Global Const MQWXP_PUT_BY_CLUSTER_CHL = &H2

'Cluster Cache Types'
Global Const MQCLCT_STATIC = 0
Global Const MQCLCT_DYNAMIC = 1

'****************************************************************'
'*  General Values Related to Exits                             *'
'****************************************************************'

'Exit Identifiers'
Global Const MQXT_API_CROSSING_EXIT = 1
Global Const MQXT_API_EXIT = 2
Global Const MQXT_CHANNEL_SEC_EXIT = 11
Global Const MQXT_CHANNEL_MSG_EXIT = 12
Global Const MQXT_CHANNEL_SEND_EXIT = 13
Global Const MQXT_CHANNEL_RCV_EXIT = 14
Global Const MQXT_CHANNEL_MSG_RETRY_EXIT = 15
Global Const MQXT_CHANNEL_AUTO_DEF_EXIT = 16
Global Const MQXT_CLUSTER_WORKLOAD_EXIT = 20
Global Const MQXT_PUBSUB_ROUTING_EXIT = 21

'Exit Reasons'
Global Const MQXR_BEFORE = 1
Global Const MQXR_AFTER = 2
Global Const MQXR_CONNECTION = 3
Global Const MQXR_INIT = 11
Global Const MQXR_TERM = 12
Global Const MQXR_MSG = 13
Global Const MQXR_XMIT = 14
Global Const MQXR_SEC_MSG = 15
Global Const MQXR_INIT_SEC = 16
Global Const MQXR_RETRY = 17
Global Const MQXR_AUTO_CLUSSDR = 18
Global Const MQXR_AUTO_RECEIVER = 19
Global Const MQXR_CLWL_OPEN = 20
Global Const MQXR_CLWL_PUT = 21
Global Const MQXR_CLWL_MOVE = 22
Global Const MQXR_CLWL_REPOS = 23
Global Const MQXR_CLWL_REPOS_MOVE = 24
Global Const MQXR_AUTO_SVRCONN = 27
Global Const MQXR_AUTO_CLUSRCVR = 28
Global Const MQXR_SEC_PARMS = 29

'Exit Responses'
Global Const MQXCC_OK = 0
Global Const MQXCC_SUPPRESS_FUNCTION = -1
Global Const MQXCC_SKIP_FUNCTION = -2
Global Const MQXCC_SEND_AND_REQUEST_SEC_MSG = -3
Global Const MQXCC_SEND_SEC_MSG = -4
Global Const MQXCC_SUPPRESS_EXIT = -5
Global Const MQXCC_CLOSE_CHANNEL = -6
Global Const MQXCC_FAILED = -8

'Exit User Area Value'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQXUA_NONE As MQBYTE16

'****************************************************************'
'*  Values Related to MQXCNVC Function                          *'
'****************************************************************'

'Conversion Options'
Global Const MQDCC_DEFAULT_CONVERSION = &H1
Global Const MQDCC_FILL_TARGET_BUFFER = &H2
Global Const MQDCC_INT_DEFAULT_CONVERSION = &H4
Global Const MQDCC_SOURCE_ENC_NATIVE = &H20
Global Const MQDCC_SOURCE_ENC_NORMAL = &H10
Global Const MQDCC_SOURCE_ENC_REVERSED = &H20
Global Const MQDCC_SOURCE_ENC_UNDEFINED = &H0
Global Const MQDCC_TARGET_ENC_NATIVE = &H200
Global Const MQDCC_TARGET_ENC_NORMAL = &H100
Global Const MQDCC_TARGET_ENC_REVERSED = &H200
Global Const MQDCC_TARGET_ENC_UNDEFINED = &H0
Global Const MQDCC_NONE = &H0

'Conversion Options Masks and Factors'
Global Const MQDCC_SOURCE_ENC_MASK = &HF0
Global Const MQDCC_TARGET_ENC_MASK = &HF00
Global Const MQDCC_SOURCE_ENC_FACTOR = 16
Global Const MQDCC_TARGET_ENC_FACTOR = 256


'****************************************************************'
'*  MQCD Structure -- Channel Definition                        *'
'****************************************************************'

Type MQCD
  ChannelName As String * 20 'Channel definition name'
  Version As Long 'Structure version number'
  ChannelType As Long 'Channel type'
  TransportType As Long 'Transport type'
  Desc As String * 64 'Channel description'
  QMgrName As String * 48 'Queue-manager name'
  XmitQName As String * 48 'Transmission queue name'
  ShortConnectionName As String * 20 'First 20 bytes of connection name'
  MCAName As String * 20 'Reserved'
  ModeName As String * 8 'LU 6.2 Mode name'
  TpName As String * 64 'LU 6.2 transaction program name'
  BatchSize As Long 'Batch size'
  DiscInterval As Long 'Disconnect interval'
  ShortRetryCount As Long 'Short retry count'
  ShortRetryInterval As Long 'Short retry wait interval'
  LongRetryCount As Long 'Long retry count'
  LongRetryInterval As Long 'Long retry wait interval'
  SecurityExit As String * 128 'Channel security exit name'
  MsgExit As String * 128 'Channel message exit name'
  SendExit As String * 128 'Channel send exit name'
  ReceiveExit As String * 128 'Channel receive exit name'
  SeqNumberWrap As Long 'Highest allowable message sequence number'
  MaxMsgLength As Long 'Maximum message length'
  PutAuthority As Long 'Put authority'
  DataConversion As Long 'Data conversion'
  SecurityUserData As String * 32 'Channel security exit user data'
  MsgUserData As String * 32 'Channel message exit user data'
  SendUserData As String * 32 'Channel send exit user data'
  ReceiveUserData As String * 32 'Channel receive exit user data'
  UserIdentifier As String * 12 'User identifier'
  Password As String * 12 'Password'
  MCAUserIdentifier As String * 12 'First 12 bytes of MCA user identifier'
  MCAType As Long 'Message channel agent type'
  ConnectionName As String * 264 'Connection name'
  RemoteUserIdentifier As String * 12 'First 12 bytes of user identifier from partner'
  RemotePassword As String * 12 'Password from partner'
  MsgRetryExit As String * 128 'Channel message retry exit name'
  MsgRetryUserData As String * 32 'Channel message retry exit user data'
  MsgRetryCount As Long 'Number of times MCA will try to put the message, after first attempt has failed'
  MsgRetryInterval As Long 'Minimum interval in milliseconds after which the open or put operation will be retried'
  HeartbeatInterval As Long 'Time in seconds between heartbeat flows'
  BatchInterval As Long 'Batch duration'
  NonPersistentMsgSpeed As Long 'Speed at which nonpersistent messages are sent'
  StrucLength As Long 'Length of MQCD structure'
  ExitNameLength As Long 'Length of exit name'
  ExitDataLength As Long 'Length of exit user data'
  MsgExitsDefined As Long 'Number of message exits defined'
  SendExitsDefined As Long 'Number of send exits defined'
  ReceiveExitsDefined As Long 'Number of receive exits defined'
  MsgExitPtr As MQPTR 'Address of first MsgExit field'
  MsgUserDataPtr As MQPTR 'Address of first MsgUserData field'
  SendExitPtr As MQPTR 'Address of first SendExit field'
  SendUserDataPtr As MQPTR 'Address of first SendUserData field'
  ReceiveExitPtr As MQPTR 'Address of first ReceiveExit field'
  ReceiveUserDataPtr As MQPTR 'Address of first ReceiveUserData field'
  ClusterPtr As MQPTR 'Address of a list of cluster names'
  ClustersDefined As Long 'Number of clusters to which the channel belongs'
  NetworkPriority As Long 'Network priority'
  LongMCAUserIdLength As Long 'Length of long MCA user identifier'
  LongRemoteUserIdLength As Long 'Length of long remote user identifier'
  LongMCAUserIdPtr As MQPTR 'Address of long MCA user identifier'
  LongRemoteUserIdPtr As MQPTR 'Address of long remote user identifier'
  MCASecurityId As MQBYTE40 'MCA security identifier'
  RemoteSecurityId As MQBYTE40 'Remote security identifier'
  SSLCipherSpec As String * 32 'SSL CipherSpec'
  SSLPeerNamePtr As MQPTR 'Address of SSL peer name'
  SSLPeerNameLength As Long 'Length of SSL peer name'
  SSLClientAuth As Long 'Whether SSL client authentication is required'
  KeepAliveInterval As Long 'Keepalive interval'
  LocalAddress As String * 48 'Local communications address'
  BatchHeartbeat As Long 'Batch heartbeat interval'
  HdrCompList(0 To 1) As Long 'Header data compression list'
  MsgCompList(0 To 15) As Long 'Message data compression list'
  CLWLChannelRank As Long 'Channel rank'
  CLWLChannelPriority As Long 'Channel priority'
  CLWLChannelWeight As Long 'Channel weight'
  ChannelMonitoring As Long 'Channel monitoring'
  ChannelStatistics As Long 'Channel statistics'
End Type

'Default Instance of MQCD Structure'
Global MQCD_DEFAULT As MQCD


'********************************************************************'
'*  MQCNOCD Structure -- Connect Options Plus Channel Definition    *'
'*                                                                  *'
'*  Use this for the "ConnectOpts" parameter of the MQCONNXAny call *'
'*  to specify the channel parameters for an MQ client application. *'
'********************************************************************'

Type MQCNOCD
  ConnectOpts As MQCNO 'Options that control the action of MQCONNX'
  ChannelDef As MQCD 'Channel definition for client connection'
End Type

'Default Instance of MQCNOCD Structure'
Global MQCNOCD_DEFAULT As MQCNOCD


'*********************************************************************'
'*  MQ_SETDEFAULTS_X Subroutine -- Set Defaults                      *'
'*********************************************************************'

'****************************************************************'
'*  End of CMQXB                                                *'
'****************************************************************'

Sub MQCD_DEFAULTS(Struc As MQCD)
  Struc.ChannelName = ""
  Struc.Version = MQCD_VERSION_6
  Struc.ChannelType = MQCHT_SENDER
  Struc.TransportType = MQXPT_LU62
  Struc.Desc = ""
  Struc.QMgrName = ""
  Struc.XmitQName = ""
  Struc.ShortConnectionName = ""
  Struc.MCAName = ""
  Struc.ModeName = ""
  Struc.TpName = ""
  Struc.BatchSize = 50
  Struc.DiscInterval = 6000
  Struc.ShortRetryCount = 10
  Struc.ShortRetryInterval = 60
  Struc.LongRetryCount = 999999999
  Struc.LongRetryInterval = 1200
  Struc.SecurityExit = ""
  Struc.MsgExit = ""
  Struc.SendExit = ""
  Struc.ReceiveExit = ""
  Struc.SeqNumberWrap = 999999999
  Struc.MaxMsgLength = 4194304
  Struc.PutAuthority = MQPA_DEFAULT
  Struc.DataConversion = MQCDC_NO_SENDER_CONVERSION
  Struc.SecurityUserData = ""
  Struc.MsgUserData = ""
  Struc.SendUserData = ""
  Struc.ReceiveUserData = ""
  Struc.UserIdentifier = ""
  Struc.Password = ""
  Struc.MCAUserIdentifier = ""
  Struc.MCAType = MQMCAT_PROCESS
  Struc.ConnectionName = ""
  Struc.RemoteUserIdentifier = ""
  Struc.RemotePassword = ""
  Struc.MsgRetryExit = ""
  Struc.MsgRetryUserData = ""
  Struc.MsgRetryCount = 10
  Struc.MsgRetryInterval = 1000
  Struc.HeartbeatInterval = 300
  Struc.BatchInterval = 0
  Struc.NonPersistentMsgSpeed = MQNPMS_FAST
  Struc.StrucLength = MQCD_CURRENT_LENGTH
  Struc.ExitNameLength = MQ_EXIT_NAME_LENGTH
  Struc.ExitDataLength = MQ_EXIT_DATA_LENGTH
  Struc.MsgExitsDefined = 0
  Struc.SendExitsDefined = 0
  Struc.ReceiveExitsDefined = 0
  Dim TempMsgExitPtr As MQPTR
  MQPTR_DEFAULTS TempMsgExitPtr
  Struc.MsgExitPtr = TempMsgExitPtr
  Dim TempMsgUserDataPtr As MQPTR
  MQPTR_DEFAULTS TempMsgUserDataPtr
  Struc.MsgUserDataPtr = TempMsgUserDataPtr
  Dim TempSendExitPtr As MQPTR
  MQPTR_DEFAULTS TempSendExitPtr
  Struc.SendExitPtr = TempSendExitPtr
  Dim TempSendUserDataPtr As MQPTR
  MQPTR_DEFAULTS TempSendUserDataPtr
  Struc.SendUserDataPtr = TempSendUserDataPtr
  Dim TempReceiveExitPtr As MQPTR
  MQPTR_DEFAULTS TempReceiveExitPtr
  Struc.ReceiveExitPtr = TempReceiveExitPtr
  Dim TempReceiveUserDataPtr As MQPTR
  MQPTR_DEFAULTS TempReceiveUserDataPtr
  Struc.ReceiveUserDataPtr = TempReceiveUserDataPtr
  Dim TempClusterPtr As MQPTR
  MQPTR_DEFAULTS TempClusterPtr
  Struc.ClusterPtr = TempClusterPtr
  Struc.ClustersDefined = 0
  Struc.NetworkPriority = 0
  Struc.LongMCAUserIdLength = 0
  Struc.LongRemoteUserIdLength = 0
  Dim TempLongMCAUserIdPtr As MQPTR
  MQPTR_DEFAULTS TempLongMCAUserIdPtr
  Struc.LongMCAUserIdPtr = TempLongMCAUserIdPtr
  Dim TempLongRemoteUserIdPtr As MQPTR
  MQPTR_DEFAULTS TempLongRemoteUserIdPtr
  Struc.LongRemoteUserIdPtr = TempLongRemoteUserIdPtr
  Dim TempMCASecurityId As MQBYTE40
  MQBYTE40_DEFAULTS TempMCASecurityId
  Struc.MCASecurityId = TempMCASecurityId
  Dim TempRemoteSecurityId As MQBYTE40
  MQBYTE40_DEFAULTS TempRemoteSecurityId
  Struc.RemoteSecurityId = TempRemoteSecurityId
  Struc.SSLCipherSpec = ""
  Dim TempSSLPeerNamePtr As MQPTR
  MQPTR_DEFAULTS TempSSLPeerNamePtr
  Struc.SSLPeerNamePtr = TempSSLPeerNamePtr
  Struc.SSLPeerNameLength = 0
  Struc.SSLClientAuth = MQSCA_REQUIRED
  Struc.KeepAliveInterval = MQKAI_AUTO
  Struc.LocalAddress = ""
  Struc.BatchHeartbeat = 0
  Struc.HdrCompList(0) = MQCOMPRESS_NONE
  Struc.HdrCompList(1) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(0) = MQCOMPRESS_NONE
  Struc.MsgCompList(1) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(2) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(3) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(4) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(5) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(6) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(7) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(8) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(9) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(10) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(11) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(12) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(13) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(14) = MQCOMPRESS_NOT_AVAILABLE
  Struc.MsgCompList(15) = MQCOMPRESS_NOT_AVAILABLE
  Struc.CLWLChannelRank = 0
  Struc.CLWLChannelPriority = 0
  Struc.CLWLChannelWeight = 50
  Struc.ChannelMonitoring = MQMON_OFF
  Struc.ChannelStatistics = MQMON_OFF
End Sub

Sub MQCNOCD_DEFAULTS(Struc As MQCNOCD)
  Dim TempConnectOpts As MQCNO
  MQCNO_DEFAULTS TempConnectOpts
  Struc.ConnectOpts = TempConnectOpts
  Dim TempChannelDef As MQCD
  MQCD_DEFAULTS TempChannelDef
  Struc.ChannelDef = TempChannelDef
End Sub

Sub MQ_SETDEFAULTS_X()

  'Set byte-string constants'
  MQBYTE48_DEFAULTS MQXPDA_NONE
  MQBYTE16_DEFAULTS MQXUA_NONE

  'Set default structures'
  MQCD_DEFAULTS MQCD_DEFAULT
  MQCNOCD_DEFAULTS MQCNOCD_DEFAULT

End Sub

'****************************************************************'
'*  End of CMQXB                                                *'
'****************************************************************'

