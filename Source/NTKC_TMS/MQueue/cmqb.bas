Attribute VB_Name = "CMQB"
'**********************************************************************'
'*                                                                    *'
'*                  WebSphere MQ for Windows                          *'
'*                                                                    *'
'*  FILE NAME:      CMQB                                              *'
'*                                                                    *'
'*  DESCRIPTION:    Declarations for Main MQI                         *'
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
'*  FUNCTION:       This file declares the functions, structures,     *'
'*                  and named constants for the main MQI.             *'
'*                                                                    *'
'*  PROCESSOR:      BASIC                                             *'
'*                                                                    *'
'**********************************************************************'

'**********************************************************************'
'*  This file is for use with the following products:                 *'
'*                                                                    *'
'*  o   Windows Visual Basic Version 6.0 (32-bit)                     *'
'*                                                                    *'
'*  This file can be used with either the WebSphere MQ server or the  *'
'*  WebSphere MQ client, depending on the conditional compilation     *'
'*  argument MqType:                                                  *'
'*                                                                    *'
'*  o   To set MqType:                                                *'
'*                                                                    *'
'*      1.  Select the menu option "Project - xx Properties" (where   *'
'*          xx is the name of the project).                           *'
'*      2.  On the tab labelled "Make", edit the field "Conditional   *'
'*          Compilation".                                             *'
'*                                                                    *'
'*  o   In the "Conditional Compiliation" field, enter (without the   *'
'*      quotes):                                                      *'
'*                                                                    *'
'*    "MqType = 1" for the WebSphere MQ server                        *'
'*    "MqType = 2" for the WebSphere MQ client                        *'
'*    "MqType = 3" for the WebSphere MQ extended transactional client *'
'*                                                                    *'
'*  The selection of client/server controls the selection of          *'
'*  WebSphere MQ DLL. The appropriate DLL must be installed and the   *'
'*  queue manager running for the Visual Basic application to         *'
'*  operate correctly.                                                *'
'*                                                                    *'
'*  To ensure that various default constants are setup properly, the  *'
'*  procedure MQ_SETDEFAULTS should be called before any other        *'
'*  WebSphere MQ calls. A good place to put this call is in the Load  *'
'*  procedure of the startup form. See the sample programs for an     *'
'*  example.                                                          *'
'*                                                                    *'
'*  COMMON PROBLEMS:                                                  *'
'*                                                                    *'
'*  ================================================                  *'
'*  Visual Basic error:                                               *'
'*  o   Run-time error '48': Error in loading DLL                     *'
'*  o   Run-time error '53': File not found                           *'
'*  ================================================                  *'
'*  PROBLEM:                                                          *'
'*  o   Incorrect Mqtype setting                                      *'
'*  SOLUTION:                                                         *'
'*  o   Change MqType setting -- see comments above                   *'
'*                                                                    *'
'*  ================================================                  *'
'*  Run-time error: CompCode = 2, Reason = 2059                       *'
'*  ================================================                  *'
'*  PROBLEM:                                                          *'
'*  o   MQRC_Q_MGR_NOT_AVAILABLE: Queue Manager not available         *'
'*  SOLUTION:                                                         *'
'*  o   Check MQSERVER environment variable:                          *'
'*        Under Windows NT/2000                                       *'
'*          Click: Start->Settings->Control Panel                     *'
'*          Double Click: System                                      *'
'*          Click tab for ENVIRONMENT                                 *'
'*          Add/Change System variable:                               *'
'*            MQSERVER=channel_name/tcp/tcp_name                      *'
'*        Note: channel_name is case sensitive                        *'
'*  o   Verify that the server can be pinged:                         *'
'*        ping tcp_name                                               *'
'*  o   Verify C samples work:                                        *'
'*        \MQM\TOOLS\SAMPLES\C\BIN\AMQSPUTC.EXE                       *'
'*        \MQM\TOOLS\SAMPLES\C\BIN\AMQSGETC.EXE                       *'
'*                                                                    *'
'**********************************************************************'

'****************************************************************'
'*  Values Related to MQAIR Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQAIR_STRUC_ID = "AIR "

'Structure Version Number'
Global Const MQAIR_VERSION_1 = 1
Global Const MQAIR_CURRENT_VERSION = 1

'Authentication Information Type'
Global Const MQAIT_CRL_LDAP = 1

'****************************************************************'
'*  Values Related to MQBO Structure                            *'
'****************************************************************'

'Structure Identifier'
Global Const MQBO_STRUC_ID = "BO  "

'Structure Version Number'
Global Const MQBO_VERSION_1 = 1
Global Const MQBO_CURRENT_VERSION = 1

'Begin Options'
Global Const MQBO_NONE = &H0

'****************************************************************'
'*  Values Related to MQCIH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQCIH_STRUC_ID = "CIH "

'Structure Version Number'
Global Const MQCIH_VERSION_1 = 1
Global Const MQCIH_VERSION_2 = 2
Global Const MQCIH_CURRENT_VERSION = 2

'Structure Length'
Global Const MQCIH_LENGTH_1 = 164
Global Const MQCIH_LENGTH_2 = 180
Global Const MQCIH_CURRENT_LENGTH = 180

'Flags'
Global Const MQCIH_NONE = &H0
Global Const MQCIH_PASS_EXPIRATION = &H1
Global Const MQCIH_UNLIMITED_EXPIRATION = &H0
Global Const MQCIH_REPLY_WITHOUT_NULLS = &H2
Global Const MQCIH_REPLY_WITH_NULLS = &H0
Global Const MQCIH_SYNC_ON_RETURN = &H4
Global Const MQCIH_NO_SYNC_ON_RETURN = &H0

'Return Codes'
Global Const MQCRC_OK = 0
Global Const MQCRC_CICS_EXEC_ERROR = 1
Global Const MQCRC_MQ_API_ERROR = 2
Global Const MQCRC_BRIDGE_ERROR = 3
Global Const MQCRC_BRIDGE_ABEND = 4
Global Const MQCRC_APPLICATION_ABEND = 5
Global Const MQCRC_SECURITY_ERROR = 6
Global Const MQCRC_PROGRAM_NOT_AVAILABLE = 7
Global Const MQCRC_BRIDGE_TIMEOUT = 8
Global Const MQCRC_TRANSID_NOT_AVAILABLE = 9

'Unit-of-Work Controls'
Global Const MQCUOWC_ONLY = &H111
Global Const MQCUOWC_CONTINUE = &H10000
Global Const MQCUOWC_FIRST = &H11
Global Const MQCUOWC_MIDDLE = &H10
Global Const MQCUOWC_LAST = &H110
Global Const MQCUOWC_COMMIT = &H100
Global Const MQCUOWC_BACKOUT = &H1100

'Get Wait Interval'
Global Const MQCGWI_DEFAULT = -2

'Link Types'
Global Const MQCLT_PROGRAM = 1
Global Const MQCLT_TRANSACTION = 2

'Output Data Length'
Global Const MQCODL_AS_INPUT = -1

'ADS Descriptors'
Global Const MQCADSD_NONE = &H0
Global Const MQCADSD_SEND = &H1
Global Const MQCADSD_RECV = &H10
Global Const MQCADSD_MSGFORMAT = &H100

'Conversational Task Options'
Global Const MQCCT_YES = &H1
Global Const MQCCT_NO = &H0

'Task End Status'
Global Const MQCTES_NOSYNC = &H0
Global Const MQCTES_COMMIT = &H100
Global Const MQCTES_BACKOUT = &H1100
Global Const MQCTES_ENDTASK = &H10000

'Facility'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQCFAC_NONE As MQBYTE8

'Functions'
Global Const MQCFUNC_MQCONN = "CONN"
Global Const MQCFUNC_MQGET = "GET "
Global Const MQCFUNC_MQINQ = "INQ "
Global Const MQCFUNC_MQOPEN = "OPEN"
Global Const MQCFUNC_MQPUT = "PUT "
Global Const MQCFUNC_MQPUT1 = "PUT1"
Global Const MQCFUNC_NONE = "    "

'Start Codes'
Global Const MQCSC_START = "S   "
Global Const MQCSC_STARTDATA = "SD  "
Global Const MQCSC_TERMINPUT = "TD  "
Global Const MQCSC_NONE = "    "

'****************************************************************'
'*  Values Related to MQCNO Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQCNO_STRUC_ID = "CNO "

'Structure Version Number'
Global Const MQCNO_VERSION_1 = 1
Global Const MQCNO_VERSION_2 = 2
Global Const MQCNO_VERSION_3 = 3
Global Const MQCNO_VERSION_4 = 4
Global Const MQCNO_VERSION_5 = 5
Global Const MQCNO_CURRENT_VERSION = 5

'Connect Options'
Global Const MQCNO_STANDARD_BINDING = &H0
Global Const MQCNO_FASTPATH_BINDING = &H1
Global Const MQCNO_SERIALIZE_CONN_TAG_Q_MGR = &H2
Global Const MQCNO_SERIALIZE_CONN_TAG_QSG = &H4
Global Const MQCNO_RESTRICT_CONN_TAG_Q_MGR = &H8
Global Const MQCNO_RESTRICT_CONN_TAG_QSG = &H10
Global Const MQCNO_HANDLE_SHARE_NONE = &H20
Global Const MQCNO_HANDLE_SHARE_BLOCK = &H40
Global Const MQCNO_HANDLE_SHARE_NO_BLOCK = &H80
Global Const MQCNO_SHARED_BINDING = &H100
Global Const MQCNO_ISOLATED_BINDING = &H200
Global Const MQCNO_ACCOUNTING_MQI_ENABLED = &H1000
Global Const MQCNO_ACCOUNTING_MQI_DISABLED = &H2000
Global Const MQCNO_ACCOUNTING_Q_ENABLED = &H4000
Global Const MQCNO_ACCOUNTING_Q_DISABLED = &H8000
Global Const MQCNO_NONE = &H0

'Queue Manager Connection Tag'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQCT_NONE As MQBYTE128

'Connection Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQCONNID_NONE As MQBYTE24

'****************************************************************'
'*  Values Related to MQCSP Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQCSP_STRUC_ID = "CSP "

'Structure Version Number'
Global Const MQCSP_VERSION_1 = 1
Global Const MQCSP_CURRENT_VERSION = 1

'Authentication Types'
Global Const MQCSP_AUTH_NONE = 0
Global Const MQCSP_AUTH_USER_ID_AND_PWD = 1

'****************************************************************'
'*  Values Related to MQDH Structure                            *'
'****************************************************************'
'Structure Identifier'
Global Const MQDH_STRUC_ID = "DH  "

'Structure Version Number'
Global Const MQDH_VERSION_1 = 1
Global Const MQDH_CURRENT_VERSION = 1

'Flags'
Global Const MQDHF_NEW_MSG_IDS = &H1
Global Const MQDHF_NONE = &H0

'Put Message Record Flags'
'See values for "Put Message Record Fields" under MQPMO'
'****************************************************************'
'*  Values Related to MQDLH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQDLH_STRUC_ID = "DLH "

'Structure Version Number'
Global Const MQDLH_VERSION_1 = 1
Global Const MQDLH_CURRENT_VERSION = 1

'****************************************************************'
'*  Values Related to MQGMO Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQGMO_STRUC_ID = "GMO "

'Structure Version Number'
Global Const MQGMO_VERSION_1 = 1
Global Const MQGMO_VERSION_2 = 2
Global Const MQGMO_VERSION_3 = 3
Global Const MQGMO_CURRENT_VERSION = 3

'Get Message Options'
Global Const MQGMO_WAIT = &H1
Global Const MQGMO_NO_WAIT = &H0
Global Const MQGMO_SET_SIGNAL = &H8
Global Const MQGMO_FAIL_IF_QUIESCING = &H2000
Global Const MQGMO_SYNCPOINT = &H2
Global Const MQGMO_SYNCPOINT_IF_PERSISTENT = &H1000
Global Const MQGMO_NO_SYNCPOINT = &H4
Global Const MQGMO_MARK_SKIP_BACKOUT = &H80
Global Const MQGMO_BROWSE_FIRST = &H10
Global Const MQGMO_BROWSE_NEXT = &H20
Global Const MQGMO_BROWSE_MSG_UNDER_CURSOR = &H800
Global Const MQGMO_MSG_UNDER_CURSOR = &H100
Global Const MQGMO_LOCK = &H200
Global Const MQGMO_UNLOCK = &H400
Global Const MQGMO_ACCEPT_TRUNCATED_MSG = &H40
Global Const MQGMO_CONVERT = &H4000
Global Const MQGMO_LOGICAL_ORDER = &H8000
Global Const MQGMO_COMPLETE_MSG = &H10000
Global Const MQGMO_ALL_MSGS_AVAILABLE = &H20000
Global Const MQGMO_ALL_SEGMENTS_AVAILABLE = &H40000
Global Const MQGMO_NONE = &H0

'Wait Interval'
Global Const MQWI_UNLIMITED = -1

'Signal Values'
Global Const MQEC_MSG_ARRIVED = 2
Global Const MQEC_WAIT_INTERVAL_EXPIRED = 3
Global Const MQEC_WAIT_CANCELED = 4
Global Const MQEC_Q_MGR_QUIESCING = 5
Global Const MQEC_CONNECTION_QUIESCING = 6

'Match Options'
Global Const MQMO_MATCH_MSG_ID = &H1
Global Const MQMO_MATCH_CORREL_ID = &H2
Global Const MQMO_MATCH_GROUP_ID = &H4
Global Const MQMO_MATCH_MSG_SEQ_NUMBER = &H8
Global Const MQMO_MATCH_OFFSET = &H10
Global Const MQMO_MATCH_MSG_TOKEN = &H20
Global Const MQMO_NONE = &H0

'Group Status'
Global Const MQGS_NOT_IN_GROUP = " "
Global Const MQGS_MSG_IN_GROUP = "G"
Global Const MQGS_LAST_MSG_IN_GROUP = "L"

'Segment Status'
Global Const MQSS_NOT_A_SEGMENT = " "
Global Const MQSS_SEGMENT = "S"
Global Const MQSS_LAST_SEGMENT = "L"

'Segmentation'
Global Const MQSEG_INHIBITED = " "
Global Const MQSEG_ALLOWED = "A"

'Message Token'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQMTOK_NONE As MQBYTE16

'Returned Length'
Global Const MQRL_UNDEFINED = -1

'****************************************************************'
'*  Values Related to MQIIH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQIIH_STRUC_ID = "IIH "

'Structure Version Number'
Global Const MQIIH_VERSION_1 = 1
Global Const MQIIH_CURRENT_VERSION = 1

'Structure Length'
Global Const MQIIH_LENGTH_1 = 84

'Flags'
Global Const MQIIH_NONE = &H0
Global Const MQIIH_PASS_EXPIRATION = &H1
Global Const MQIIH_UNLIMITED_EXPIRATION = &H0
Global Const MQIIH_REPLY_FORMAT_NONE = &H8

'Authenticator'
Global Const MQIAUT_NONE = "        "

'Transaction Instance Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQITII_NONE As MQBYTE16

'Transaction States'
Global Const MQITS_IN_CONVERSATION = "C"
Global Const MQITS_NOT_IN_CONVERSATION = " "
Global Const MQITS_ARCHITECTED = "A"

'Commit Modes'
Global Const MQICM_COMMIT_THEN_SEND = "0"
Global Const MQICM_SEND_THEN_COMMIT = "1"

'Security Scopes'
Global Const MQISS_CHECK = "C"
Global Const MQISS_FULL = "F"

'****************************************************************'
'*  Values Related to MQMD Structure                            *'
'****************************************************************'
'Structure Identifier'
Global Const MQMD_STRUC_ID = "MD  "

'Structure Version Number'
Global Const MQMD_VERSION_1 = 1
Global Const MQMD_VERSION_2 = 2
Global Const MQMD_CURRENT_VERSION = 2

'Report Options'
Global Const MQRO_EXCEPTION = &H1000000
Global Const MQRO_EXCEPTION_WITH_DATA = &H3000000
Global Const MQRO_EXCEPTION_WITH_FULL_DATA = &H7000000
Global Const MQRO_EXPIRATION = &H200000
Global Const MQRO_EXPIRATION_WITH_DATA = &H600000
Global Const MQRO_EXPIRATION_WITH_FULL_DATA = &HE00000
Global Const MQRO_COA = &H100
Global Const MQRO_COA_WITH_DATA = &H300
Global Const MQRO_COA_WITH_FULL_DATA = &H700
Global Const MQRO_COD = &H800
Global Const MQRO_COD_WITH_DATA = &H1800
Global Const MQRO_COD_WITH_FULL_DATA = &H3800
Global Const MQRO_PAN = &H1
Global Const MQRO_NAN = &H2
Global Const MQRO_ACTIVITY = &H4
Global Const MQRO_NEW_MSG_ID = &H0
Global Const MQRO_PASS_MSG_ID = &H80
Global Const MQRO_COPY_MSG_ID_TO_CORREL_ID = &H0
Global Const MQRO_PASS_CORREL_ID = &H40
Global Const MQRO_DEAD_LETTER_Q = &H0
Global Const MQRO_DISCARD_MSG = &H8000000
Global Const MQRO_PASS_DISCARD_AND_EXPIRY = &H4000
Global Const MQRO_NONE = &H0

'Report Options Masks'
Global Const MQRO_REJECT_UNSUP_MASK = &H101C0000
Global Const MQRO_ACCEPT_UNSUP_MASK = &HEFE000FF
Global Const MQRO_ACCEPT_UNSUP_IF_XMIT_MASK = &H3FF00

'Message Types'
Global Const MQMT_SYSTEM_FIRST = 1
Global Const MQMT_REQUEST = 1
Global Const MQMT_REPLY = 2
Global Const MQMT_DATAGRAM = 8
Global Const MQMT_REPORT = 4
Global Const MQMT_MQE_FIELDS_FROM_MQE = 112
Global Const MQMT_MQE_FIELDS = 113
Global Const MQMT_SYSTEM_LAST = 65535
Global Const MQMT_APPL_FIRST = 65536
Global Const MQMT_APPL_LAST = 999999999

'Expiry'
Global Const MQEI_UNLIMITED = -1

'Feedback Values'
Global Const MQFB_NONE = 0
Global Const MQFB_SYSTEM_FIRST = 1
Global Const MQFB_QUIT = 256
Global Const MQFB_EXPIRATION = 258
Global Const MQFB_COA = 259
Global Const MQFB_COD = 260
Global Const MQFB_CHANNEL_COMPLETED = 262
Global Const MQFB_CHANNEL_FAIL_RETRY = 263
Global Const MQFB_CHANNEL_FAIL = 264
Global Const MQFB_APPL_CANNOT_BE_STARTED = 265
Global Const MQFB_TM_ERROR = 266
Global Const MQFB_APPL_TYPE_ERROR = 267
Global Const MQFB_STOPPED_BY_MSG_EXIT = 268
Global Const MQFB_ACTIVITY = 269
Global Const MQFB_XMIT_Q_MSG_ERROR = 271
Global Const MQFB_PAN = 275
Global Const MQFB_NAN = 276
Global Const MQFB_STOPPED_BY_CHAD_EXIT = 277
Global Const MQFB_STOPPED_BY_PUBSUB_EXIT = 279
Global Const MQFB_NOT_A_REPOSITORY_MSG = 280
Global Const MQFB_BIND_OPEN_CLUSRCVR_DEL = 281
Global Const MQFB_MAX_ACTIVITIES = 282
Global Const MQFB_NOT_FORWARDED = 283
Global Const MQFB_NOT_DELIVERED = 284
Global Const MQFB_UNSUPPORTED_FORWARDING = 285
Global Const MQFB_UNSUPPORTED_DELIVERY = 286
Global Const MQFB_DATA_LENGTH_ZERO = 291
Global Const MQFB_DATA_LENGTH_NEGATIVE = 292
Global Const MQFB_DATA_LENGTH_TOO_BIG = 293
Global Const MQFB_BUFFER_OVERFLOW = 294
Global Const MQFB_LENGTH_OFF_BY_ONE = 295
Global Const MQFB_IIH_ERROR = 296
Global Const MQFB_NOT_AUTHORIZED_FOR_IMS = 298
Global Const MQFB_IMS_ERROR = 300
Global Const MQFB_IMS_FIRST = 301
Global Const MQFB_IMS_LAST = 399
Global Const MQFB_CICS_INTERNAL_ERROR = 401
Global Const MQFB_CICS_NOT_AUTHORIZED = 402
Global Const MQFB_CICS_BRIDGE_FAILURE = 403
Global Const MQFB_CICS_CORREL_ID_ERROR = 404
Global Const MQFB_CICS_CCSID_ERROR = 405
Global Const MQFB_CICS_ENCODING_ERROR = 406
Global Const MQFB_CICS_CIH_ERROR = 407
Global Const MQFB_CICS_UOW_ERROR = 408
Global Const MQFB_CICS_COMMAREA_ERROR = 409
Global Const MQFB_CICS_APPL_NOT_STARTED = 410
Global Const MQFB_CICS_APPL_ABENDED = 411
Global Const MQFB_CICS_DLQ_ERROR = 412
Global Const MQFB_CICS_UOW_BACKED_OUT = 413
Global Const MQFB_SYSTEM_LAST = 65535
Global Const MQFB_APPL_FIRST = 65536
Global Const MQFB_APPL_LAST = 999999999

'Encoding'
Global Const MQENC_NATIVE = &H222

'Encoding Masks'
Global Const MQENC_INTEGER_MASK = &HF
Global Const MQENC_DECIMAL_MASK = &HF0
Global Const MQENC_FLOAT_MASK = &HF00
Global Const MQENC_RESERVED_MASK = &HFFFFF000

'Encodings for Binary Integers'
Global Const MQENC_INTEGER_UNDEFINED = &H0
Global Const MQENC_INTEGER_NORMAL = &H1
Global Const MQENC_INTEGER_REVERSED = &H2

'Encodings for Packed Decimal Integers'
Global Const MQENC_DECIMAL_UNDEFINED = &H0
Global Const MQENC_DECIMAL_NORMAL = &H10
Global Const MQENC_DECIMAL_REVERSED = &H20

'Encodings for Floating Point Numbers'
Global Const MQENC_FLOAT_UNDEFINED = &H0
Global Const MQENC_FLOAT_IEEE_NORMAL = &H100
Global Const MQENC_FLOAT_IEEE_REVERSED = &H200
Global Const MQENC_FLOAT_S390 = &H300
Global Const MQENC_FLOAT_TNS = &H400

'Coded Character Set Identifiers'
Global Const MQCCSI_UNDEFINED = 0
Global Const MQCCSI_DEFAULT = 0
Global Const MQCCSI_Q_MGR = 0
Global Const MQCCSI_INHERIT = -2
Global Const MQCCSI_EMBEDDED = -1

'Formats'
Global Const MQFMT_NONE = "        "
Global Const MQFMT_ADMIN = "MQADMIN "
Global Const MQFMT_CHANNEL_COMPLETED = "MQCHCOM "
Global Const MQFMT_CICS = "MQCICS  "
Global Const MQFMT_COMMAND_1 = "MQCMD1  "
Global Const MQFMT_COMMAND_2 = "MQCMD2  "
Global Const MQFMT_DEAD_LETTER_HEADER = "MQDEAD  "
Global Const MQFMT_DIST_HEADER = "MQHDIST "
Global Const MQFMT_EMBEDDED_PCF = "MQHEPCF "
Global Const MQFMT_EVENT = "MQEVENT "
Global Const MQFMT_IMS = "MQIMS   "
Global Const MQFMT_IMS_VAR_STRING = "MQIMSVS "
Global Const MQFMT_MD_EXTENSION = "MQHMDE  "
Global Const MQFMT_PCF = "MQPCF   "
Global Const MQFMT_REF_MSG_HEADER = "MQHREF  "
Global Const MQFMT_RF_HEADER = "MQHRF   "
Global Const MQFMT_RF_HEADER_1 = "MQHRF   "
Global Const MQFMT_RF_HEADER_2 = "MQHRF2  "
Global Const MQFMT_STRING = "MQSTR   "
Global Const MQFMT_TRIGGER = "MQTRIG  "
Global Const MQFMT_WORK_INFO_HEADER = "MQHWIH  "
Global Const MQFMT_XMIT_Q_HEADER = "MQXMIT  "

'Priority'
Global Const MQPRI_PRIORITY_AS_Q_DEF = -1

'Persistence Values'
Global Const MQPER_NOT_PERSISTENT = 0
Global Const MQPER_PERSISTENT = 1
Global Const MQPER_PERSISTENCE_AS_Q_DEF = 2

'Message Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQMI_NONE As MQBYTE24

'Correlation Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQCI_NONE As MQBYTE24
Public MQCI_NEW_SESSION As MQBYTE24

'Accounting Token'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQACT_NONE As MQBYTE32

'Accounting Token Types'
Global Const MQACTT_UNKNOWN = &H0
Global Const MQACTT_CICS_LUOW_ID = &H1
Global Const MQACTT_OS2_DEFAULT = &H4
Global Const MQACTT_DOS_DEFAULT = &H5
Global Const MQACTT_UNIX_NUMERIC_ID = &H6
Global Const MQACTT_OS400_ACCOUNT_TOKEN = &H8
Global Const MQACTT_WINDOWS_DEFAULT = &H9
Global Const MQACTT_NT_SECURITY_ID = &HB
Global Const MQACTT_USER = &H19

'Put Application Types'
Global Const MQAT_UNKNOWN = -1
Global Const MQAT_NO_CONTEXT = 0
Global Const MQAT_CICS = 1
Global Const MQAT_MVS = 2
Global Const MQAT_OS390 = 2
Global Const MQAT_ZOS = 2
Global Const MQAT_IMS = 3
Global Const MQAT_OS2 = 4
Global Const MQAT_DOS = 5
Global Const MQAT_AIX = 6
Global Const MQAT_UNIX = 6
Global Const MQAT_QMGR = 7
Global Const MQAT_OS400 = 8
Global Const MQAT_WINDOWS = 9
Global Const MQAT_CICS_VSE = 10
Global Const MQAT_WINDOWS_NT = 11
Global Const MQAT_VMS = 12
Global Const MQAT_GUARDIAN = 13
Global Const MQAT_NSK = 13
Global Const MQAT_VOS = 14
Global Const MQAT_IMS_BRIDGE = 19
Global Const MQAT_XCF = 20
Global Const MQAT_CICS_BRIDGE = 21
Global Const MQAT_NOTES_AGENT = 22
Global Const MQAT_USER = 25
Global Const MQAT_BROKER = 26
Global Const MQAT_JAVA = 28
Global Const MQAT_DQM = 29
Global Const MQAT_CHANNEL_INITIATOR = 30
Global Const MQAT_WLM = 31
Global Const MQAT_BATCH = 32
Global Const MQAT_RRS_BATCH = 33
Global Const MQAT_SIB = 34
Global Const MQAT_DEFAULT = 11
Global Const MQAT_USER_FIRST = 65536
Global Const MQAT_USER_LAST = 999999999

'Group Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQGI_NONE As MQBYTE24

'Message Flags'
Global Const MQMF_SEGMENTATION_INHIBITED = &H0
Global Const MQMF_SEGMENTATION_ALLOWED = &H1
Global Const MQMF_MSG_IN_GROUP = &H8
Global Const MQMF_LAST_MSG_IN_GROUP = &H10
Global Const MQMF_SEGMENT = &H2
Global Const MQMF_LAST_SEGMENT = &H4
Global Const MQMF_NONE = &H0

'Message Flags Masks'
Global Const MQMF_REJECT_UNSUP_MASK = &HFFF
Global Const MQMF_ACCEPT_UNSUP_MASK = &HFFF00000
Global Const MQMF_ACCEPT_UNSUP_IF_XMIT_MASK = &HFF000

'Original Length'
Global Const MQOL_UNDEFINED = -1

'****************************************************************'
'*  Values Related to MQMDE Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQMDE_STRUC_ID = "MDE "

'Structure Version Number'
Global Const MQMDE_VERSION_2 = 2
Global Const MQMDE_CURRENT_VERSION = 2

'Structure Length'
Global Const MQMDE_LENGTH_2 = 72

'Flags'
Global Const MQMDEF_NONE = &H0

'****************************************************************'
'*  Values Related to MQOD Structure                            *'
'****************************************************************'
'Structure Identifier'
Global Const MQOD_STRUC_ID = "OD  "

'Structure Version Number'
Global Const MQOD_VERSION_1 = 1
Global Const MQOD_VERSION_2 = 2
Global Const MQOD_VERSION_3 = 3
Global Const MQOD_CURRENT_VERSION = 3

'Structure Length'
Global Const MQOD_CURRENT_LENGTH = 336

'Object Types'
Global Const MQOT_Q = 1
Global Const MQOT_NAMELIST = 2
Global Const MQOT_PROCESS = 3
Global Const MQOT_STORAGE_CLASS = 4
Global Const MQOT_Q_MGR = 5
Global Const MQOT_CHANNEL = 6
Global Const MQOT_AUTH_INFO = 7
Global Const MQOT_CF_STRUC = 10
Global Const MQOT_LISTENER = 11
Global Const MQOT_SERVICE = 12
Global Const MQOT_RESERVED_1 = 999

'Extended Object Types'
Global Const MQOT_ALL = 1001
Global Const MQOT_ALIAS_Q = 1002
Global Const MQOT_MODEL_Q = 1003
Global Const MQOT_LOCAL_Q = 1004
Global Const MQOT_REMOTE_Q = 1005
Global Const MQOT_SENDER_CHANNEL = 1007
Global Const MQOT_SERVER_CHANNEL = 1008
Global Const MQOT_REQUESTER_CHANNEL = 1009
Global Const MQOT_RECEIVER_CHANNEL = 1010
Global Const MQOT_CURRENT_CHANNEL = 1011
Global Const MQOT_SAVED_CHANNEL = 1012
Global Const MQOT_SVRCONN_CHANNEL = 1013
Global Const MQOT_CLNTCONN_CHANNEL = 1014
Global Const MQOT_SHORT_CHANNEL = 1015

'Security Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQSID_NONE As MQBYTE40

'Security Identifier Types'
Global Const MQSIDT_NONE = &H0
Global Const MQSIDT_NT_SECURITY_ID = &H1
Global Const MQSIDT_WAS_SECURITY_ID = &H2

'****************************************************************'
'*  Values Related to MQPMO Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQPMO_STRUC_ID = "PMO "

'Structure Version Number'
Global Const MQPMO_VERSION_1 = 1
Global Const MQPMO_VERSION_2 = 2
Global Const MQPMO_CURRENT_VERSION = 2

'Structure Length'
Global Const MQPMO_CURRENT_LENGTH = 152

'Put Message Options'
Global Const MQPMO_SYNCPOINT = &H2
Global Const MQPMO_NO_SYNCPOINT = &H4
Global Const MQPMO_NEW_MSG_ID = &H40
Global Const MQPMO_NEW_CORREL_ID = &H80
Global Const MQPMO_LOGICAL_ORDER = &H8000
Global Const MQPMO_NO_CONTEXT = &H4000
Global Const MQPMO_DEFAULT_CONTEXT = &H20
Global Const MQPMO_PASS_IDENTITY_CONTEXT = &H100
Global Const MQPMO_PASS_ALL_CONTEXT = &H200
Global Const MQPMO_SET_IDENTITY_CONTEXT = &H400
Global Const MQPMO_SET_ALL_CONTEXT = &H800
Global Const MQPMO_ALTERNATE_USER_AUTHORITY = &H1000
Global Const MQPMO_FAIL_IF_QUIESCING = &H2000
Global Const MQPMO_RESOLVE_LOCAL_Q = &H40000
Global Const MQPMO_NONE = &H0

'Put Message Record Fields'
Global Const MQPMRF_MSG_ID = &H1
Global Const MQPMRF_CORREL_ID = &H2
Global Const MQPMRF_GROUP_ID = &H4
Global Const MQPMRF_FEEDBACK = &H8
Global Const MQPMRF_ACCOUNTING_TOKEN = &H10
Global Const MQPMRF_NONE = &H0

'****************************************************************'
'*  Values Related to MQRFH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQRFH_STRUC_ID = "RFH "

'Structure Version Number'
Global Const MQRFH_VERSION_1 = 1
Global Const MQRFH_VERSION_2 = 2

'Structure Length'
Global Const MQRFH_STRUC_LENGTH_FIXED = 32
Global Const MQRFH_STRUC_LENGTH_FIXED_2 = 36

'Flags'
Global Const MQRFH_NONE = &H0
Global Const MQRFH_NO_FLAGS = 0

'Names for Name/Value String'
Global Const MQNVS_APPL_TYPE = "OPT_APP_GRP "
Global Const MQNVS_MSG_TYPE = "OPT_MSG_TYPE "

'****************************************************************'
'*  Values Related to MQRMH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQRMH_STRUC_ID = "RMH "

'Structure Version Number'
Global Const MQRMH_VERSION_1 = 1
Global Const MQRMH_CURRENT_VERSION = 1

'Flags'
Global Const MQRMHF_LAST = &H1
Global Const MQRMHF_NOT_LAST = &H0

'Object Instance Identifier'
'(call the MQ_SETDEFAULTS subroutine to initialize the following)'
Public MQOII_NONE As MQBYTE24

'****************************************************************'
'*  Values Related to MQSCO Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQSCO_STRUC_ID = "SCO "

'Structure Version Number'
Global Const MQSCO_VERSION_1 = 1
Global Const MQSCO_VERSION_2 = 2
Global Const MQSCO_CURRENT_VERSION = 2

'Key Reset Count'
Global Const MQSCO_RESET_COUNT_DEFAULT = 0

'****************************************************************'
'*  Values Related to MQTM Structure                            *'
'****************************************************************'
'Structure Identifier'
Global Const MQTM_STRUC_ID = "TM  "

'Structure Version Number'
Global Const MQTM_VERSION_1 = 1
Global Const MQTM_CURRENT_VERSION = 1

'****************************************************************'
'*  Values Related to MQTMC2 Structure                          *'
'****************************************************************'
'Structure Identifier'
Global Const MQTMC_STRUC_ID = "TMC "

'Structure Version Number'
Global Const MQTMC_VERSION_1 = "   1"
Global Const MQTMC_VERSION_2 = "   2"
Global Const MQTMC_CURRENT_VERSION = "   2"

'****************************************************************'
'*  Values Related to MQWIH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQWIH_STRUC_ID = "WIH "

'Structure Version Number'
Global Const MQWIH_VERSION_1 = 1
Global Const MQWIH_CURRENT_VERSION = 1

'Structure Length'
Global Const MQWIH_LENGTH_1 = 120
Global Const MQWIH_CURRENT_LENGTH = 120

'Flags'
Global Const MQWIH_NONE = &H0

'****************************************************************'
'*  Values Related to MQXQH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQXQH_STRUC_ID = "XQH "

'Structure Version Number'
Global Const MQXQH_VERSION_1 = 1
Global Const MQXQH_CURRENT_VERSION = 1

'****************************************************************'
'*  Values Related to MQCLOSE Function                          *'
'****************************************************************'

'Object Handle'
Global Const MQHO_UNUSABLE_HOBJ = -1
Global Const MQHO_NONE = 0

'Close Options'
Global Const MQCO_NONE = &H0
Global Const MQCO_DELETE = &H1
Global Const MQCO_DELETE_PURGE = &H2

'****************************************************************'
'*  Values Related to MQINQ Function                            *'
'****************************************************************'

'Byte Attribute Selectors'
Global Const MQBA_FIRST = 6001
Global Const MQBA_LAST = 8000

'Character Attribute Selectors'
Global Const MQCA_ALTERATION_DATE = 2027
Global Const MQCA_ALTERATION_TIME = 2028
Global Const MQCA_APPL_ID = 2001
Global Const MQCA_AUTH_INFO_CONN_NAME = 2053
Global Const MQCA_AUTH_INFO_DESC = 2046
Global Const MQCA_AUTH_INFO_NAME = 2045
Global Const MQCA_AUTO_REORG_CATALOG = 2091
Global Const MQCA_AUTO_REORG_START_TIME = 2090
Global Const MQCA_BACKOUT_REQ_Q_NAME = 2019
Global Const MQCA_BASE_Q_NAME = 2002
Global Const MQCA_BATCH_INTERFACE_ID = 2068
Global Const MQCA_CF_STRUC_DESC = 2052
Global Const MQCA_CF_STRUC_NAME = 2039
Global Const MQCA_CHANNEL_AUTO_DEF_EXIT = 2026
Global Const MQCA_CHINIT_SERVICE_PARM = 2076
Global Const MQCA_CICS_FILE_NAME = 2060
Global Const MQCA_CLUSTER_DATE = 2037
Global Const MQCA_CLUSTER_NAME = 2029
Global Const MQCA_CLUSTER_NAMELIST = 2030
Global Const MQCA_CLUSTER_Q_MGR_NAME = 2031
Global Const MQCA_CLUSTER_TIME = 2038
Global Const MQCA_CLUSTER_WORKLOAD_DATA = 2034
Global Const MQCA_CLUSTER_WORKLOAD_EXIT = 2033
Global Const MQCA_COMMAND_INPUT_Q_NAME = 2003
Global Const MQCA_COMMAND_REPLY_Q_NAME = 2067
Global Const MQCA_CREATION_DATE = 2004
Global Const MQCA_CREATION_TIME = 2005
Global Const MQCA_DEAD_LETTER_Q_NAME = 2006
Global Const MQCA_DEF_XMIT_Q_NAME = 2025
Global Const MQCA_DNS_GROUP = 2071
Global Const MQCA_ENV_DATA = 2007
Global Const MQCA_FIRST = 2001
Global Const MQCA_IGQ_USER_ID = 2041
Global Const MQCA_INITIATION_Q_NAME = 2008
Global Const MQCA_LAST = 4000
Global Const MQCA_LAST_USED = 2091
Global Const MQCA_LDAP_PASSWORD = 2048
Global Const MQCA_LDAP_USER_NAME = 2047
Global Const MQCA_LU_GROUP_NAME = 2072
Global Const MQCA_LU_NAME = 2073
Global Const MQCA_LU62_ARM_SUFFIX = 2074
Global Const MQCA_MONITOR_Q_NAME = 2066
Global Const MQCA_NAMELIST_DESC = 2009
Global Const MQCA_NAMELIST_NAME = 2010
Global Const MQCA_NAMES = 2020
Global Const MQCA_PASS_TICKET_APPL = 2086
Global Const MQCA_PROCESS_DESC = 2011
Global Const MQCA_PROCESS_NAME = 2012
Global Const MQCA_Q_DESC = 2013
Global Const MQCA_Q_MGR_DESC = 2014
Global Const MQCA_Q_MGR_IDENTIFIER = 2032
Global Const MQCA_Q_MGR_NAME = 2015
Global Const MQCA_Q_NAME = 2016
Global Const MQCA_QSG_NAME = 2040
Global Const MQCA_REMOTE_Q_MGR_NAME = 2017
Global Const MQCA_REMOTE_Q_NAME = 2018
Global Const MQCA_REPOSITORY_NAME = 2035
Global Const MQCA_REPOSITORY_NAMELIST = 2036
Global Const MQCA_SERVICE_DESC = 2078
Global Const MQCA_SERVICE_NAME = 2077
Global Const MQCA_SERVICE_START_ARGS = 2080
Global Const MQCA_SERVICE_START_COMMAND = 2079
Global Const MQCA_SERVICE_STOP_ARGS = 2082
Global Const MQCA_SERVICE_STOP_COMMAND = 2081
Global Const MQCA_STDERR_DESTINATION = 2084
Global Const MQCA_STDOUT_DESTINATION = 2083
Global Const MQCA_SSL_CRL_NAMELIST = 2050
Global Const MQCA_SSL_CRYPTO_HARDWARE = 2051
Global Const MQCA_SSL_KEY_LIBRARY = 2069
Global Const MQCA_SSL_KEY_MEMBER = 2070
Global Const MQCA_SSL_KEY_REPOSITORY = 2049
Global Const MQCA_STORAGE_CLASS = 2022
Global Const MQCA_STORAGE_CLASS_DESC = 2042
Global Const MQCA_SYSTEM_LOG_Q_NAME = 2065
Global Const MQCA_TCP_NAME = 2075
Global Const MQCA_TPIPE_NAME = 2085
Global Const MQCA_TRIGGER_CHANNEL_NAME = 2064
Global Const MQCA_TRIGGER_DATA = 2023
Global Const MQCA_TRIGGER_PROGRAM_NAME = 2062
Global Const MQCA_TRIGGER_TERM_ID = 2063
Global Const MQCA_TRIGGER_TRANS_ID = 2061
Global Const MQCA_USER_DATA = 2021
Global Const MQCA_USER_LIST = 4000
Global Const MQCA_XCF_GROUP_NAME = 2043
Global Const MQCA_XCF_MEMBER_NAME = 2044
Global Const MQCA_XMIT_Q_NAME = 2024

'Integer Attribute Selectors'
Global Const MQIA_ACCOUNTING_CONN_OVERRIDE = 136
Global Const MQIA_ACCOUNTING_INTERVAL = 135
Global Const MQIA_ACCOUNTING_MQI = 133
Global Const MQIA_ACCOUNTING_Q = 134
Global Const MQIA_ACTIVE_CHANNELS = 100
Global Const MQIA_ACTIVITY_RECORDING = 138
Global Const MQIA_ADOPTNEWMCA_CHECK = 102
Global Const MQIA_ADOPTNEWMCA_TYPE = 103
Global Const MQIA_ADOPTNEWMCA_INTERVAL = 104
Global Const MQIA_APPL_TYPE = 1
Global Const MQIA_ARCHIVE = 60
Global Const MQIA_AUTH_INFO_TYPE = 66
Global Const MQIA_AUTHORITY_EVENT = 47
Global Const MQIA_AUTO_REORG_INTERVAL = 174
Global Const MQIA_AUTO_REORGANIZATION = 173
Global Const MQIA_BACKOUT_THRESHOLD = 22
Global Const MQIA_BATCH_INTERFACE_AUTO = 86
Global Const MQIA_BRIDGE_EVENT = 74
Global Const MQIA_CF_LEVEL = 70
Global Const MQIA_CF_RECOVER = 71
Global Const MQIA_CHANNEL_AUTO_DEF = 55
Global Const MQIA_CHANNEL_AUTO_DEF_EVENT = 56
Global Const MQIA_CHANNEL_EVENT = 73
Global Const MQIA_CHINIT_ADAPTERS = 101
Global Const MQIA_CHINIT_CONTROL = 119
Global Const MQIA_CHINIT_DISPATCHERS = 105
Global Const MQIA_CHINIT_TRACE_AUTO_START = 117
Global Const MQIA_CHINIT_TRACE_TABLE_SIZE = 118
Global Const MQIA_CLUSTER_Q_TYPE = 59
Global Const MQIA_CLUSTER_WORKLOAD_LENGTH = 58
Global Const MQIA_CLWL_MRU_CHANNELS = 97
Global Const MQIA_CLWL_Q_RANK = 95
Global Const MQIA_CLWL_Q_PRIORITY = 96
Global Const MQIA_CLWL_USEQ = 98
Global Const MQIA_CMD_SERVER_AUTO = 87
Global Const MQIA_CMD_SERVER_CONTROL = 120
Global Const MQIA_CMD_SERVER_CONVERT_MSG = 88
Global Const MQIA_CMD_SERVER_DLQ_MSG = 89
Global Const MQIA_CODED_CHAR_SET_ID = 2
Global Const MQIA_COMMAND_EVENT = 99
Global Const MQIA_COMMAND_LEVEL = 31
Global Const MQIA_CONFIGURATION_EVENT = 51
Global Const MQIA_CPI_LEVEL = 27
Global Const MQIA_CURRENT_Q_DEPTH = 3
Global Const MQIA_DEF_BIND = 61
Global Const MQIA_DEF_INPUT_OPEN_OPTION = 4
Global Const MQIA_DEF_PERSISTENCE = 5
Global Const MQIA_DEF_PRIORITY = 6
Global Const MQIA_DEFINITION_TYPE = 7
Global Const MQIA_DIST_LISTS = 34
Global Const MQIA_DNS_WLM = 106
Global Const MQIA_EXPIRY_INTERVAL = 39
Global Const MQIA_FIRST = 1
Global Const MQIA_HARDEN_GET_BACKOUT = 8
Global Const MQIA_HIGH_Q_DEPTH = 36
Global Const MQIA_IGQ_PUT_AUTHORITY = 65
Global Const MQIA_INDEX_TYPE = 57
Global Const MQIA_INHIBIT_EVENT = 48
Global Const MQIA_INHIBIT_GET = 9
Global Const MQIA_INHIBIT_PUT = 10
Global Const MQIA_INTRA_GROUP_QUEUING = 64
Global Const MQIA_IP_ADDRESS_VERSION = 93
Global Const MQIA_LAST = 2000
Global Const MQIA_LAST_USED = 174
Global Const MQIA_LISTENER_PORT_NUMBER = 85
Global Const MQIA_LISTENER_TIMER = 107
Global Const MQIA_LOGGER_EVENT = 94
Global Const MQIA_LU62_CHANNELS = 108
Global Const MQIA_LOCAL_EVENT = 49
Global Const MQIA_MAX_CHANNELS = 109
Global Const MQIA_MAX_CLIENTS = 172
Global Const MQIA_MAX_GLOBAL_LOCKS = 83
Global Const MQIA_MAX_HANDLES = 11
Global Const MQIA_MAX_LOCAL_LOCKS = 84
Global Const MQIA_MAX_MSG_LENGTH = 13
Global Const MQIA_MAX_OPEN_Q = 80
Global Const MQIA_MAX_PRIORITY = 14
Global Const MQIA_MAX_Q_DEPTH = 15
Global Const MQIA_MAX_Q_TRIGGERS = 90
Global Const MQIA_MAX_RECOVERY_TASKS = 171
Global Const MQIA_MAX_UNCOMMITTED_MSGS = 33
Global Const MQIA_MONITOR_INTERVAL = 81
Global Const MQIA_MONITORING_AUTO_CLUSSDR = 124
Global Const MQIA_MONITORING_CHANNEL = 122
Global Const MQIA_MONITORING_Q = 123
Global Const MQIA_MSG_DELIVERY_SEQUENCE = 16
Global Const MQIA_MSG_DEQ_COUNT = 38
Global Const MQIA_MSG_ENQ_COUNT = 37
Global Const MQIA_NAME_COUNT = 19
Global Const MQIA_NAMELIST_TYPE = 72
Global Const MQIA_NPM_CLASS = 78
Global Const MQIA_OPEN_INPUT_COUNT = 17
Global Const MQIA_OPEN_OUTPUT_COUNT = 18
Global Const MQIA_OUTBOUND_PORT_MAX = 140
Global Const MQIA_OUTBOUND_PORT_MIN = 110
Global Const MQIA_PAGESET_ID = 62
Global Const MQIA_PERFORMANCE_EVENT = 53
Global Const MQIA_PLATFORM = 32
Global Const MQIA_Q_DEPTH_HIGH_EVENT = 43
Global Const MQIA_Q_DEPTH_HIGH_LIMIT = 40
Global Const MQIA_Q_DEPTH_LOW_EVENT = 44
Global Const MQIA_Q_DEPTH_LOW_LIMIT = 41
Global Const MQIA_Q_DEPTH_MAX_EVENT = 42
Global Const MQIA_Q_SERVICE_INTERVAL = 54
Global Const MQIA_Q_SERVICE_INTERVAL_EVENT = 46
Global Const MQIA_Q_TYPE = 20
Global Const MQIA_Q_USERS = 82
Global Const MQIA_QMOPT_CONS_COMMS_MSGS = 155
Global Const MQIA_QMOPT_CONS_CRITICAL_MSGS = 154
Global Const MQIA_QMOPT_CONS_ERROR_MSGS = 153
Global Const MQIA_QMOPT_CONS_INFO_MSGS = 151
Global Const MQIA_QMOPT_CONS_REORG_MSGS = 156
Global Const MQIA_QMOPT_CONS_SYSTEM_MSGS = 157
Global Const MQIA_QMOPT_CONS_WARNING_MSGS = 152
Global Const MQIA_QMOPT_CSMT_ON_ERROR = 150
Global Const MQIA_QMOPT_INTERNAL_DUMP = 170
Global Const MQIA_QMOPT_LOG_COMMS_MSGS = 162
Global Const MQIA_QMOPT_LOG_CRITICAL_MSGS = 161
Global Const MQIA_QMOPT_LOG_ERROR_MSGS = 160
Global Const MQIA_QMOPT_LOG_INFO_MSGS = 158
Global Const MQIA_QMOPT_LOG_REORG_MSGS = 163
Global Const MQIA_QMOPT_LOG_SYSTEM_MSGS = 164
Global Const MQIA_QMOPT_LOG_WARNING_MSGS = 159
Global Const MQIA_QMOPT_TRACE_COMMS = 166
Global Const MQIA_QMOPT_TRACE_CONVERSION = 168
Global Const MQIA_QMOPT_TRACE_REORG = 167
Global Const MQIA_QMOPT_TRACE_MQI_CALLS = 165
Global Const MQIA_QMOPT_TRACE_SYSTEM = 169
Global Const MQIA_QSG_DISP = 63
Global Const MQIA_RECEIVE_TIMEOUT = 111
Global Const MQIA_RECEIVE_TIMEOUT_TYPE = 112
Global Const MQIA_RECEIVE_TIMEOUT_MIN = 113
Global Const MQIA_REMOTE_EVENT = 50
Global Const MQIA_RETENTION_INTERVAL = 21
Global Const MQIA_SCOPE = 45
Global Const MQIA_SERVICE_CONTROL = 139
Global Const MQIA_SERVICE_TYPE = 121
Global Const MQIA_SHAREABILITY = 23
Global Const MQIA_SHARED_Q_Q_MGR_NAME = 77
Global Const MQIA_SSL_EVENT = 75
Global Const MQIA_SSL_FIPS_REQUIRED = 92
Global Const MQIA_SSL_RESET_COUNT = 76
Global Const MQIA_SSL_TASKS = 69
Global Const MQIA_START_STOP_EVENT = 52
Global Const MQIA_STATISTICS_CHANNEL = 129
Global Const MQIA_STATISTICS_AUTO_CLUSSDR = 130
Global Const MQIA_STATISTICS_INTERVAL = 131
Global Const MQIA_STATISTICS_MQI = 127
Global Const MQIA_STATISTICS_Q = 128
Global Const MQIA_SYNCPOINT = 30
Global Const MQIA_TCP_CHANNELS = 114
Global Const MQIA_TCP_KEEP_ALIVE = 115
Global Const MQIA_TCP_STACK_TYPE = 116
Global Const MQIA_TIME_SINCE_RESET = 35
Global Const MQIA_TRACE_ROUTE_RECORDING = 137
Global Const MQIA_TRIGGER_CONTROL = 24
Global Const MQIA_TRIGGER_DEPTH = 29
Global Const MQIA_TRIGGER_INTERVAL = 25
Global Const MQIA_TRIGGER_MSG_PRIORITY = 26
Global Const MQIA_TRIGGER_TYPE = 28
Global Const MQIA_TRIGGER_RESTART = 91
Global Const MQIA_USAGE = 12
Global Const MQIA_USER_LIST = 2000

'Integer Attribute Values'
Global Const MQIAV_NOT_APPLICABLE = -1
Global Const MQIAV_UNDEFINED = -2

'Group Attribute Selectors'
Global Const MQGA_FIRST = 8001
Global Const MQGA_LAST = 9000

'****************************************************************'
'*  Values Related to MQOPEN Function                           *'
'****************************************************************'

'Open Options'
Global Const MQOO_INPUT_AS_Q_DEF = &H1
Global Const MQOO_INPUT_SHARED = &H2
Global Const MQOO_INPUT_EXCLUSIVE = &H4
Global Const MQOO_BROWSE = &H8
Global Const MQOO_OUTPUT = &H10
Global Const MQOO_INQUIRE = &H20
Global Const MQOO_SET = &H40
Global Const MQOO_BIND_ON_OPEN = &H4000
Global Const MQOO_BIND_NOT_FIXED = &H8000
Global Const MQOO_BIND_AS_Q_DEF = &H0
Global Const MQOO_SAVE_ALL_CONTEXT = &H80
Global Const MQOO_PASS_IDENTITY_CONTEXT = &H100
Global Const MQOO_PASS_ALL_CONTEXT = &H200
Global Const MQOO_SET_IDENTITY_CONTEXT = &H400
Global Const MQOO_SET_ALL_CONTEXT = &H800
Global Const MQOO_ALTERNATE_USER_AUTHORITY = &H1000
Global Const MQOO_FAIL_IF_QUIESCING = &H2000
Global Const MQOO_RESOLVE_NAMES = &H10000      'C++ only'

Global Const MQOO_RESOLVE_LOCAL_Q = &H40000

'****************************************************************'
'*  Values Related to All Functions                             *'
'****************************************************************'

'Connection Handles'
Global Const MQHC_DEF_HCONN = 0
Global Const MQHC_UNUSABLE_HCONN = -1

'String Lengths'
Global Const MQ_ABEND_CODE_LENGTH = 4
Global Const MQ_ACCOUNTING_TOKEN_LENGTH = 32
Global Const MQ_APPL_IDENTITY_DATA_LENGTH = 32
Global Const MQ_APPL_NAME_LENGTH = 28
Global Const MQ_APPL_ORIGIN_DATA_LENGTH = 4
Global Const MQ_APPL_TAG_LENGTH = 28
Global Const MQ_ARM_SUFFIX_LENGTH = 2
Global Const MQ_ATTENTION_ID_LENGTH = 4
Global Const MQ_AUTH_INFO_CONN_NAME_LENGTH = 264
Global Const MQ_AUTH_INFO_DESC_LENGTH = 64
Global Const MQ_AUTH_INFO_NAME_LENGTH = 48
Global Const MQ_AUTHENTICATOR_LENGTH = 8
Global Const MQ_AUTO_REORG_CATALOG_LENGTH = 44
Global Const MQ_AUTO_REORG_TIME_LENGTH = 4
Global Const MQ_BATCH_INTERFACE_ID_LENGTH = 8
Global Const MQ_BRIDGE_NAME_LENGTH = 24
Global Const MQ_CANCEL_CODE_LENGTH = 4
Global Const MQ_CF_STRUC_DESC_LENGTH = 64
Global Const MQ_CF_STRUC_NAME_LENGTH = 12
Global Const MQ_CHANNEL_DATE_LENGTH = 12
Global Const MQ_CHANNEL_DESC_LENGTH = 64
Global Const MQ_CHANNEL_NAME_LENGTH = 20
Global Const MQ_CHANNEL_TIME_LENGTH = 8
Global Const MQ_CHINIT_SERVICE_PARM_LENGTH = 32
Global Const MQ_CICS_FILE_NAME_LENGTH = 8
Global Const MQ_CLUSTER_NAME_LENGTH = 48
Global Const MQ_CONN_NAME_LENGTH = 264
Global Const MQ_CONN_TAG_LENGTH = 128
Global Const MQ_CONNECTION_ID_LENGTH = 24
Global Const MQ_CORREL_ID_LENGTH = 24
Global Const MQ_CREATION_DATE_LENGTH = 12
Global Const MQ_CREATION_TIME_LENGTH = 8
Global Const MQ_DATE_LENGTH = 12
Global Const MQ_DISTINGUISHED_NAME_LENGTH = 1024
Global Const MQ_DNS_GROUP_NAME_LENGTH = 18
Global Const MQ_EXIT_DATA_LENGTH = 32
Global Const MQ_EXIT_INFO_NAME_LENGTH = 48
Global Const MQ_EXIT_NAME_LENGTH = 128
Global Const MQ_EXIT_PD_AREA_LENGTH = 48
Global Const MQ_EXIT_USER_AREA_LENGTH = 16
Global Const MQ_FACILITY_LENGTH = 8
Global Const MQ_FACILITY_LIKE_LENGTH = 4
Global Const MQ_FORMAT_LENGTH = 8
Global Const MQ_FUNCTION_LENGTH = 4
Global Const MQ_GROUP_ID_LENGTH = 24
Global Const MQ_LDAP_PASSWORD_LENGTH = 32
Global Const MQ_LISTENER_NAME_LENGTH = 48
Global Const MQ_LISTENER_DESC_LENGTH = 64
Global Const MQ_LOCAL_ADDRESS_LENGTH = 48
Global Const MQ_LTERM_OVERRIDE_LENGTH = 8
Global Const MQ_LU_NAME_LENGTH = 8
Global Const MQ_LUWID_LENGTH = 16
Global Const MQ_MAX_EXIT_NAME_LENGTH = 128
Global Const MQ_MAX_MCA_USER_ID_LENGTH = 64
Global Const MQ_MAX_USER_ID_LENGTH = 64
Global Const MQ_MCA_JOB_NAME_LENGTH = 28
Global Const MQ_MCA_NAME_LENGTH = 20
Global Const MQ_MCA_USER_DATA_LENGTH = 32
Global Const MQ_MCA_USER_ID_LENGTH = 64
Global Const MQ_MFS_MAP_NAME_LENGTH = 8
Global Const MQ_MODE_NAME_LENGTH = 8
Global Const MQ_MSG_HEADER_LENGTH = 4000
Global Const MQ_MSG_ID_LENGTH = 24
Global Const MQ_MSG_TOKEN_LENGTH = 16
Global Const MQ_NAMELIST_DESC_LENGTH = 64
Global Const MQ_NAMELIST_NAME_LENGTH = 48
Global Const MQ_OBJECT_INSTANCE_ID_LENGTH = 24
Global Const MQ_OBJECT_NAME_LENGTH = 48
Global Const MQ_PASS_TICKET_APPL_LENGTH = 8
Global Const MQ_PASSWORD_LENGTH = 12
Global Const MQ_PROCESS_APPL_ID_LENGTH = 256
Global Const MQ_PROCESS_DESC_LENGTH = 64
Global Const MQ_PROCESS_ENV_DATA_LENGTH = 128
Global Const MQ_PROCESS_NAME_LENGTH = 48
Global Const MQ_PROCESS_USER_DATA_LENGTH = 128
Global Const MQ_PROGRAM_NAME_LENGTH = 20
Global Const MQ_PUT_APPL_NAME_LENGTH = 28
Global Const MQ_PUT_DATE_LENGTH = 8
Global Const MQ_PUT_TIME_LENGTH = 8
Global Const MQ_Q_DESC_LENGTH = 64
Global Const MQ_Q_MGR_DESC_LENGTH = 64
Global Const MQ_Q_MGR_IDENTIFIER_LENGTH = 48
Global Const MQ_Q_MGR_NAME_LENGTH = 48
Global Const MQ_Q_NAME_LENGTH = 48
Global Const MQ_QSG_NAME_LENGTH = 4
Global Const MQ_REMOTE_SYS_ID_LENGTH = 4
Global Const MQ_SECURITY_ID_LENGTH = 40
Global Const MQ_SERVICE_ARGS_LENGTH = 255
Global Const MQ_SERVICE_COMMAND_LENGTH = 255
Global Const MQ_SERVICE_DESC_LENGTH = 64
Global Const MQ_SERVICE_NAME_LENGTH = 32
Global Const MQ_SERVICE_PATH_LENGTH = 255
Global Const MQ_SERVICE_STEP_LENGTH = 8
Global Const MQ_SHORT_CONN_NAME_LENGTH = 20
Global Const MQ_SHORT_DNAME_LENGTH = 256
Global Const MQ_SSL_CIPHER_SPEC_LENGTH = 32
Global Const MQ_SSL_CRYPTO_HARDWARE_LENGTH = 256
Global Const MQ_SSL_HANDSHAKE_STAGE_LENGTH = 32
Global Const MQ_SSL_KEY_LIBRARY_LENGTH = 44
Global Const MQ_SSL_KEY_MEMBER_LENGTH = 8
Global Const MQ_SSL_KEY_REPOSITORY_LENGTH = 256
Global Const MQ_SSL_PEER_NAME_LENGTH = 1024
Global Const MQ_SSL_SHORT_PEER_NAME_LENGTH = 256
Global Const MQ_START_CODE_LENGTH = 4
Global Const MQ_STORAGE_CLASS_DESC_LENGTH = 64
Global Const MQ_STORAGE_CLASS_LENGTH = 8
Global Const MQ_SUB_IDENTITY_LENGTH = 128
Global Const MQ_TCP_NAME_LENGTH = 8
Global Const MQ_TIME_LENGTH = 8
Global Const MQ_TOTAL_EXIT_DATA_LENGTH = 999
Global Const MQ_TOTAL_EXIT_NAME_LENGTH = 999
Global Const MQ_TP_NAME_LENGTH = 64
Global Const MQ_TPIPE_NAME_LENGTH = 8
Global Const MQ_TRAN_INSTANCE_ID_LENGTH = 16
Global Const MQ_TRANSACTION_ID_LENGTH = 4
Global Const MQ_TRIGGER_DATA_LENGTH = 64
Global Const MQ_TRIGGER_PROGRAM_NAME_LENGTH = 8
Global Const MQ_TRIGGER_TERM_ID_LENGTH = 4
Global Const MQ_TRIGGER_TRANS_ID_LENGTH = 4
Global Const MQ_USER_ID_LENGTH = 12
Global Const MQ_XCF_GROUP_NAME_LENGTH = 8
Global Const MQ_XCF_MEMBER_NAME_LENGTH = 16

'Completion Codes'
Global Const MQCC_OK = 0
Global Const MQCC_WARNING = 1
Global Const MQCC_FAILED = 2
Global Const MQCC_UNKNOWN = -1

'Reason Codes'
Global Const MQRC_NONE = 0
Global Const MQRC_APPL_FIRST = 900
Global Const MQRC_APPL_LAST = 999
Global Const MQRC_ALIAS_BASE_Q_TYPE_ERROR = 2001
Global Const MQRC_ALREADY_CONNECTED = 2002
Global Const MQRC_BACKED_OUT = 2003
Global Const MQRC_BUFFER_ERROR = 2004
Global Const MQRC_BUFFER_LENGTH_ERROR = 2005
Global Const MQRC_CHAR_ATTR_LENGTH_ERROR = 2006
Global Const MQRC_CHAR_ATTRS_ERROR = 2007
Global Const MQRC_CHAR_ATTRS_TOO_SHORT = 2008
Global Const MQRC_CONNECTION_BROKEN = 2009
Global Const MQRC_DATA_LENGTH_ERROR = 2010
Global Const MQRC_DYNAMIC_Q_NAME_ERROR = 2011
Global Const MQRC_ENVIRONMENT_ERROR = 2012
Global Const MQRC_EXPIRY_ERROR = 2013
Global Const MQRC_FEEDBACK_ERROR = 2014
Global Const MQRC_GET_INHIBITED = 2016
Global Const MQRC_HANDLE_NOT_AVAILABLE = 2017
Global Const MQRC_HCONN_ERROR = 2018
Global Const MQRC_HOBJ_ERROR = 2019
Global Const MQRC_INHIBIT_VALUE_ERROR = 2020
Global Const MQRC_INT_ATTR_COUNT_ERROR = 2021
Global Const MQRC_INT_ATTR_COUNT_TOO_SMALL = 2022
Global Const MQRC_INT_ATTRS_ARRAY_ERROR = 2023
Global Const MQRC_SYNCPOINT_LIMIT_REACHED = 2024
Global Const MQRC_MAX_CONNS_LIMIT_REACHED = 2025
Global Const MQRC_MD_ERROR = 2026
Global Const MQRC_MISSING_REPLY_TO_Q = 2027
Global Const MQRC_MSG_TYPE_ERROR = 2029
Global Const MQRC_MSG_TOO_BIG_FOR_Q = 2030
Global Const MQRC_MSG_TOO_BIG_FOR_Q_MGR = 2031
Global Const MQRC_NO_MSG_AVAILABLE = 2033
Global Const MQRC_NO_MSG_UNDER_CURSOR = 2034
Global Const MQRC_NOT_AUTHORIZED = 2035
Global Const MQRC_NOT_OPEN_FOR_BROWSE = 2036
Global Const MQRC_NOT_OPEN_FOR_INPUT = 2037
Global Const MQRC_NOT_OPEN_FOR_INQUIRE = 2038
Global Const MQRC_NOT_OPEN_FOR_OUTPUT = 2039
Global Const MQRC_NOT_OPEN_FOR_SET = 2040
Global Const MQRC_OBJECT_CHANGED = 2041
Global Const MQRC_OBJECT_IN_USE = 2042
Global Const MQRC_OBJECT_TYPE_ERROR = 2043
Global Const MQRC_OD_ERROR = 2044
Global Const MQRC_OPTION_NOT_VALID_FOR_TYPE = 2045
Global Const MQRC_OPTIONS_ERROR = 2046
Global Const MQRC_PERSISTENCE_ERROR = 2047
Global Const MQRC_PERSISTENT_NOT_ALLOWED = 2048
Global Const MQRC_PRIORITY_EXCEEDS_MAXIMUM = 2049
Global Const MQRC_PRIORITY_ERROR = 2050
Global Const MQRC_PUT_INHIBITED = 2051
Global Const MQRC_Q_DELETED = 2052
Global Const MQRC_Q_FULL = 2053
Global Const MQRC_Q_NOT_EMPTY = 2055
Global Const MQRC_Q_SPACE_NOT_AVAILABLE = 2056
Global Const MQRC_Q_TYPE_ERROR = 2057
Global Const MQRC_Q_MGR_NAME_ERROR = 2058
Global Const MQRC_Q_MGR_NOT_AVAILABLE = 2059
Global Const MQRC_REPORT_OPTIONS_ERROR = 2061
Global Const MQRC_SECOND_MARK_NOT_ALLOWED = 2062
Global Const MQRC_SECURITY_ERROR = 2063
Global Const MQRC_SELECTOR_COUNT_ERROR = 2065
Global Const MQRC_SELECTOR_LIMIT_EXCEEDED = 2066
Global Const MQRC_SELECTOR_ERROR = 2067
Global Const MQRC_SELECTOR_NOT_FOR_TYPE = 2068
Global Const MQRC_SIGNAL_OUTSTANDING = 2069
Global Const MQRC_SIGNAL_REQUEST_ACCEPTED = 2070
Global Const MQRC_STORAGE_NOT_AVAILABLE = 2071
Global Const MQRC_SYNCPOINT_NOT_AVAILABLE = 2072
Global Const MQRC_TRIGGER_CONTROL_ERROR = 2075
Global Const MQRC_TRIGGER_DEPTH_ERROR = 2076
Global Const MQRC_TRIGGER_MSG_PRIORITY_ERR = 2077
Global Const MQRC_TRIGGER_TYPE_ERROR = 2078
Global Const MQRC_TRUNCATED_MSG_ACCEPTED = 2079
Global Const MQRC_TRUNCATED_MSG_FAILED = 2080
Global Const MQRC_UNKNOWN_ALIAS_BASE_Q = 2082
Global Const MQRC_UNKNOWN_OBJECT_NAME = 2085
Global Const MQRC_UNKNOWN_OBJECT_Q_MGR = 2086
Global Const MQRC_UNKNOWN_REMOTE_Q_MGR = 2087
Global Const MQRC_WAIT_INTERVAL_ERROR = 2090
Global Const MQRC_XMIT_Q_TYPE_ERROR = 2091
Global Const MQRC_XMIT_Q_USAGE_ERROR = 2092
Global Const MQRC_NOT_OPEN_FOR_PASS_ALL = 2093
Global Const MQRC_NOT_OPEN_FOR_PASS_IDENT = 2094
Global Const MQRC_NOT_OPEN_FOR_SET_ALL = 2095
Global Const MQRC_NOT_OPEN_FOR_SET_IDENT = 2096
Global Const MQRC_CONTEXT_HANDLE_ERROR = 2097
Global Const MQRC_CONTEXT_NOT_AVAILABLE = 2098
Global Const MQRC_SIGNAL1_ERROR = 2099
Global Const MQRC_OBJECT_ALREADY_EXISTS = 2100
Global Const MQRC_OBJECT_DAMAGED = 2101
Global Const MQRC_RESOURCE_PROBLEM = 2102
Global Const MQRC_ANOTHER_Q_MGR_CONNECTED = 2103
Global Const MQRC_UNKNOWN_REPORT_OPTION = 2104
Global Const MQRC_STORAGE_CLASS_ERROR = 2105
Global Const MQRC_COD_NOT_VALID_FOR_XCF_Q = 2106
Global Const MQRC_XWAIT_CANCELED = 2107
Global Const MQRC_XWAIT_ERROR = 2108
Global Const MQRC_SUPPRESSED_BY_EXIT = 2109
Global Const MQRC_FORMAT_ERROR = 2110
Global Const MQRC_SOURCE_CCSID_ERROR = 2111
Global Const MQRC_SOURCE_INTEGER_ENC_ERROR = 2112
Global Const MQRC_SOURCE_DECIMAL_ENC_ERROR = 2113
Global Const MQRC_SOURCE_FLOAT_ENC_ERROR = 2114
Global Const MQRC_TARGET_CCSID_ERROR = 2115
Global Const MQRC_TARGET_INTEGER_ENC_ERROR = 2116
Global Const MQRC_TARGET_DECIMAL_ENC_ERROR = 2117
Global Const MQRC_TARGET_FLOAT_ENC_ERROR = 2118
Global Const MQRC_NOT_CONVERTED = 2119
Global Const MQRC_CONVERTED_MSG_TOO_BIG = 2120
Global Const MQRC_TRUNCATED = 2120
Global Const MQRC_NO_EXTERNAL_PARTICIPANTS = 2121
Global Const MQRC_PARTICIPANT_NOT_AVAILABLE = 2122
Global Const MQRC_OUTCOME_MIXED = 2123
Global Const MQRC_OUTCOME_PENDING = 2124
Global Const MQRC_BRIDGE_STARTED = 2125
Global Const MQRC_BRIDGE_STOPPED = 2126
Global Const MQRC_ADAPTER_STORAGE_SHORTAGE = 2127
Global Const MQRC_UOW_IN_PROGRESS = 2128
Global Const MQRC_ADAPTER_CONN_LOAD_ERROR = 2129
Global Const MQRC_ADAPTER_SERV_LOAD_ERROR = 2130
Global Const MQRC_ADAPTER_DEFS_ERROR = 2131
Global Const MQRC_ADAPTER_DEFS_LOAD_ERROR = 2132
Global Const MQRC_ADAPTER_CONV_LOAD_ERROR = 2133
Global Const MQRC_BO_ERROR = 2134
Global Const MQRC_DH_ERROR = 2135
Global Const MQRC_MULTIPLE_REASONS = 2136
Global Const MQRC_OPEN_FAILED = 2137
Global Const MQRC_ADAPTER_DISC_LOAD_ERROR = 2138
Global Const MQRC_CNO_ERROR = 2139
Global Const MQRC_CICS_WAIT_FAILED = 2140
Global Const MQRC_DLH_ERROR = 2141
Global Const MQRC_HEADER_ERROR = 2142
Global Const MQRC_SOURCE_LENGTH_ERROR = 2143
Global Const MQRC_TARGET_LENGTH_ERROR = 2144
Global Const MQRC_SOURCE_BUFFER_ERROR = 2145
Global Const MQRC_TARGET_BUFFER_ERROR = 2146
Global Const MQRC_IIH_ERROR = 2148
Global Const MQRC_PCF_ERROR = 2149
Global Const MQRC_DBCS_ERROR = 2150
Global Const MQRC_OBJECT_NAME_ERROR = 2152
Global Const MQRC_OBJECT_Q_MGR_NAME_ERROR = 2153
Global Const MQRC_RECS_PRESENT_ERROR = 2154
Global Const MQRC_OBJECT_RECORDS_ERROR = 2155
Global Const MQRC_RESPONSE_RECORDS_ERROR = 2156
Global Const MQRC_ASID_MISMATCH = 2157
Global Const MQRC_PMO_RECORD_FLAGS_ERROR = 2158
Global Const MQRC_PUT_MSG_RECORDS_ERROR = 2159
Global Const MQRC_CONN_ID_IN_USE = 2160
Global Const MQRC_Q_MGR_QUIESCING = 2161
Global Const MQRC_Q_MGR_STOPPING = 2162
Global Const MQRC_DUPLICATE_RECOV_COORD = 2163
Global Const MQRC_PMO_ERROR = 2173
Global Const MQRC_API_EXIT_NOT_FOUND = 2182
Global Const MQRC_API_EXIT_LOAD_ERROR = 2183
Global Const MQRC_REMOTE_Q_NAME_ERROR = 2184
Global Const MQRC_INCONSISTENT_PERSISTENCE = 2185
Global Const MQRC_GMO_ERROR = 2186
Global Const MQRC_CICS_BRIDGE_RESTRICTION = 2187
Global Const MQRC_STOPPED_BY_CLUSTER_EXIT = 2188
Global Const MQRC_CLUSTER_RESOLUTION_ERROR = 2189
Global Const MQRC_CONVERTED_STRING_TOO_BIG = 2190
Global Const MQRC_TMC_ERROR = 2191
Global Const MQRC_PAGESET_FULL = 2192
Global Const MQRC_STORAGE_MEDIUM_FULL = 2192
Global Const MQRC_PAGESET_ERROR = 2193
Global Const MQRC_NAME_NOT_VALID_FOR_TYPE = 2194
Global Const MQRC_UNEXPECTED_ERROR = 2195
Global Const MQRC_UNKNOWN_XMIT_Q = 2196
Global Const MQRC_UNKNOWN_DEF_XMIT_Q = 2197
Global Const MQRC_DEF_XMIT_Q_TYPE_ERROR = 2198
Global Const MQRC_DEF_XMIT_Q_USAGE_ERROR = 2199
Global Const MQRC_NAME_IN_USE = 2201
Global Const MQRC_CONNECTION_QUIESCING = 2202
Global Const MQRC_CONNECTION_STOPPING = 2203
Global Const MQRC_ADAPTER_NOT_AVAILABLE = 2204
Global Const MQRC_MSG_ID_ERROR = 2206
Global Const MQRC_CORREL_ID_ERROR = 2207
Global Const MQRC_FILE_SYSTEM_ERROR = 2208
Global Const MQRC_NO_MSG_LOCKED = 2209
Global Const MQRC_SOAP_DOTNET_ERROR = 2210
Global Const MQRC_SOAP_AXIS_ERROR = 2211
Global Const MQRC_SOAP_URL_ERROR = 2212
Global Const MQRC_FILE_NOT_AUDITED = 2216
Global Const MQRC_CONNECTION_NOT_AUTHORIZED = 2217
Global Const MQRC_MSG_TOO_BIG_FOR_CHANNEL = 2218
Global Const MQRC_CALL_IN_PROGRESS = 2219
Global Const MQRC_RMH_ERROR = 2220
Global Const MQRC_Q_MGR_ACTIVE = 2222
Global Const MQRC_Q_MGR_NOT_ACTIVE = 2223
Global Const MQRC_Q_DEPTH_HIGH = 2224
Global Const MQRC_Q_DEPTH_LOW = 2225
Global Const MQRC_Q_SERVICE_INTERVAL_HIGH = 2226
Global Const MQRC_Q_SERVICE_INTERVAL_OK = 2227
Global Const MQRC_UNIT_OF_WORK_NOT_STARTED = 2232
Global Const MQRC_CHANNEL_AUTO_DEF_OK = 2233
Global Const MQRC_CHANNEL_AUTO_DEF_ERROR = 2234
Global Const MQRC_CFH_ERROR = 2235
Global Const MQRC_CFIL_ERROR = 2236
Global Const MQRC_CFIN_ERROR = 2237
Global Const MQRC_CFSL_ERROR = 2238
Global Const MQRC_CFST_ERROR = 2239
Global Const MQRC_INCOMPLETE_GROUP = 2241
Global Const MQRC_INCOMPLETE_MSG = 2242
Global Const MQRC_INCONSISTENT_CCSIDS = 2243
Global Const MQRC_INCONSISTENT_ENCODINGS = 2244
Global Const MQRC_INCONSISTENT_UOW = 2245
Global Const MQRC_INVALID_MSG_UNDER_CURSOR = 2246
Global Const MQRC_MATCH_OPTIONS_ERROR = 2247
Global Const MQRC_MDE_ERROR = 2248
Global Const MQRC_MSG_FLAGS_ERROR = 2249
Global Const MQRC_MSG_SEQ_NUMBER_ERROR = 2250
Global Const MQRC_OFFSET_ERROR = 2251
Global Const MQRC_ORIGINAL_LENGTH_ERROR = 2252
Global Const MQRC_SEGMENT_LENGTH_ZERO = 2253
Global Const MQRC_UOW_NOT_AVAILABLE = 2255
Global Const MQRC_WRONG_GMO_VERSION = 2256
Global Const MQRC_WRONG_MD_VERSION = 2257
Global Const MQRC_GROUP_ID_ERROR = 2258
Global Const MQRC_INCONSISTENT_BROWSE = 2259
Global Const MQRC_XQH_ERROR = 2260
Global Const MQRC_SRC_ENV_ERROR = 2261
Global Const MQRC_SRC_NAME_ERROR = 2262
Global Const MQRC_DEST_ENV_ERROR = 2263
Global Const MQRC_DEST_NAME_ERROR = 2264
Global Const MQRC_TM_ERROR = 2265
Global Const MQRC_CLUSTER_EXIT_ERROR = 2266
Global Const MQRC_CLUSTER_EXIT_LOAD_ERROR = 2267
Global Const MQRC_CLUSTER_PUT_INHIBITED = 2268
Global Const MQRC_CLUSTER_RESOURCE_ERROR = 2269
Global Const MQRC_NO_DESTINATIONS_AVAILABLE = 2270
Global Const MQRC_CONN_TAG_IN_USE = 2271
Global Const MQRC_PARTIALLY_CONVERTED = 2272
Global Const MQRC_CONNECTION_ERROR = 2273
Global Const MQRC_OPTION_ENVIRONMENT_ERROR = 2274
Global Const MQRC_CD_ERROR = 2277
Global Const MQRC_CLIENT_CONN_ERROR = 2278
Global Const MQRC_CHANNEL_STOPPED_BY_USER = 2279
Global Const MQRC_HCONFIG_ERROR = 2280
Global Const MQRC_FUNCTION_ERROR = 2281
Global Const MQRC_CHANNEL_STARTED = 2282
Global Const MQRC_CHANNEL_STOPPED = 2283
Global Const MQRC_CHANNEL_CONV_ERROR = 2284
Global Const MQRC_SERVICE_NOT_AVAILABLE = 2285
Global Const MQRC_INITIALIZATION_FAILED = 2286
Global Const MQRC_TERMINATION_FAILED = 2287
Global Const MQRC_UNKNOWN_Q_NAME = 2288
Global Const MQRC_SERVICE_ERROR = 2289
Global Const MQRC_Q_ALREADY_EXISTS = 2290
Global Const MQRC_USER_ID_NOT_AVAILABLE = 2291
Global Const MQRC_UNKNOWN_ENTITY = 2292
Global Const MQRC_UNKNOWN_AUTH_ENTITY = 2293
Global Const MQRC_UNKNOWN_REF_OBJECT = 2294
Global Const MQRC_CHANNEL_ACTIVATED = 2295
Global Const MQRC_CHANNEL_NOT_ACTIVATED = 2296
Global Const MQRC_UOW_CANCELED = 2297
Global Const MQRC_FUNCTION_NOT_SUPPORTED = 2298
Global Const MQRC_SELECTOR_TYPE_ERROR = 2299
Global Const MQRC_COMMAND_TYPE_ERROR = 2300
Global Const MQRC_MULTIPLE_INSTANCE_ERROR = 2301
Global Const MQRC_SYSTEM_ITEM_NOT_ALTERABLE = 2302
Global Const MQRC_BAG_CONVERSION_ERROR = 2303
Global Const MQRC_SELECTOR_OUT_OF_RANGE = 2304
Global Const MQRC_SELECTOR_NOT_UNIQUE = 2305
Global Const MQRC_INDEX_NOT_PRESENT = 2306
Global Const MQRC_STRING_ERROR = 2307
Global Const MQRC_ENCODING_NOT_SUPPORTED = 2308
Global Const MQRC_SELECTOR_NOT_PRESENT = 2309
Global Const MQRC_OUT_SELECTOR_ERROR = 2310
Global Const MQRC_STRING_TRUNCATED = 2311
Global Const MQRC_SELECTOR_WRONG_TYPE = 2312
Global Const MQRC_INCONSISTENT_ITEM_TYPE = 2313
Global Const MQRC_INDEX_ERROR = 2314
Global Const MQRC_SYSTEM_BAG_NOT_ALTERABLE = 2315
Global Const MQRC_ITEM_COUNT_ERROR = 2316
Global Const MQRC_FORMAT_NOT_SUPPORTED = 2317
Global Const MQRC_SELECTOR_NOT_SUPPORTED = 2318
Global Const MQRC_ITEM_VALUE_ERROR = 2319
Global Const MQRC_HBAG_ERROR = 2320
Global Const MQRC_PARAMETER_MISSING = 2321
Global Const MQRC_CMD_SERVER_NOT_AVAILABLE = 2322
Global Const MQRC_STRING_LENGTH_ERROR = 2323
Global Const MQRC_INQUIRY_COMMAND_ERROR = 2324
Global Const MQRC_NESTED_BAG_NOT_SUPPORTED = 2325
Global Const MQRC_BAG_WRONG_TYPE = 2326
Global Const MQRC_ITEM_TYPE_ERROR = 2327
Global Const MQRC_SYSTEM_BAG_NOT_DELETABLE = 2328
Global Const MQRC_SYSTEM_ITEM_NOT_DELETABLE = 2329
Global Const MQRC_CODED_CHAR_SET_ID_ERROR = 2330
Global Const MQRC_MSG_TOKEN_ERROR = 2331
Global Const MQRC_MISSING_WIH = 2332
Global Const MQRC_WIH_ERROR = 2333
Global Const MQRC_RFH_ERROR = 2334
Global Const MQRC_RFH_STRING_ERROR = 2335
Global Const MQRC_RFH_COMMAND_ERROR = 2336
Global Const MQRC_RFH_PARM_ERROR = 2337
Global Const MQRC_RFH_DUPLICATE_PARM = 2338
Global Const MQRC_RFH_PARM_MISSING = 2339
Global Const MQRC_CHAR_CONVERSION_ERROR = 2340
Global Const MQRC_UCS2_CONVERSION_ERROR = 2341
Global Const MQRC_DB2_NOT_AVAILABLE = 2342
Global Const MQRC_OBJECT_NOT_UNIQUE = 2343
Global Const MQRC_CONN_TAG_NOT_RELEASED = 2344
Global Const MQRC_CF_NOT_AVAILABLE = 2345
Global Const MQRC_CF_STRUC_IN_USE = 2346
Global Const MQRC_CF_STRUC_LIST_HDR_IN_USE = 2347
Global Const MQRC_CF_STRUC_AUTH_FAILED = 2348
Global Const MQRC_CF_STRUC_ERROR = 2349
Global Const MQRC_CONN_TAG_NOT_USABLE = 2350
Global Const MQRC_GLOBAL_UOW_CONFLICT = 2351
Global Const MQRC_LOCAL_UOW_CONFLICT = 2352
Global Const MQRC_HANDLE_IN_USE_FOR_UOW = 2353
Global Const MQRC_UOW_ENLISTMENT_ERROR = 2354
Global Const MQRC_UOW_MIX_NOT_SUPPORTED = 2355
Global Const MQRC_WXP_ERROR = 2356
Global Const MQRC_CURRENT_RECORD_ERROR = 2357
Global Const MQRC_NEXT_OFFSET_ERROR = 2358
Global Const MQRC_NO_RECORD_AVAILABLE = 2359
Global Const MQRC_OBJECT_LEVEL_INCOMPATIBLE = 2360
Global Const MQRC_NEXT_RECORD_ERROR = 2361
Global Const MQRC_BACKOUT_THRESHOLD_REACHED = 2362
Global Const MQRC_MSG_NOT_MATCHED = 2363
Global Const MQRC_JMS_FORMAT_ERROR = 2364
Global Const MQRC_SEGMENTS_NOT_SUPPORTED = 2365
Global Const MQRC_WRONG_CF_LEVEL = 2366
Global Const MQRC_CONFIG_CREATE_OBJECT = 2367
Global Const MQRC_CONFIG_CHANGE_OBJECT = 2368
Global Const MQRC_CONFIG_DELETE_OBJECT = 2369
Global Const MQRC_CONFIG_REFRESH_OBJECT = 2370
Global Const MQRC_CHANNEL_SSL_ERROR = 2371
Global Const MQRC_CF_STRUC_FAILED = 2373
Global Const MQRC_API_EXIT_ERROR = 2374
Global Const MQRC_API_EXIT_INIT_ERROR = 2375
Global Const MQRC_API_EXIT_TERM_ERROR = 2376
Global Const MQRC_EXIT_REASON_ERROR = 2377
Global Const MQRC_RESERVED_VALUE_ERROR = 2378
Global Const MQRC_NO_DATA_AVAILABLE = 2379
Global Const MQRC_SCO_ERROR = 2380
Global Const MQRC_KEY_REPOSITORY_ERROR = 2381
Global Const MQRC_CRYPTO_HARDWARE_ERROR = 2382
Global Const MQRC_AUTH_INFO_REC_COUNT_ERROR = 2383
Global Const MQRC_AUTH_INFO_REC_ERROR = 2384
Global Const MQRC_AIR_ERROR = 2385
Global Const MQRC_AUTH_INFO_TYPE_ERROR = 2386
Global Const MQRC_AUTH_INFO_CONN_NAME_ERROR = 2387
Global Const MQRC_LDAP_USER_NAME_ERROR = 2388
Global Const MQRC_LDAP_USER_NAME_LENGTH_ERR = 2389
Global Const MQRC_LDAP_PASSWORD_ERROR = 2390
Global Const MQRC_SSL_ALREADY_INITIALIZED = 2391
Global Const MQRC_SSL_CONFIG_ERROR = 2392
Global Const MQRC_SSL_INITIALIZATION_ERROR = 2393
Global Const MQRC_Q_INDEX_TYPE_ERROR = 2394
Global Const MQRC_CFBS_ERROR = 2395
Global Const MQRC_SSL_NOT_ALLOWED = 2396
Global Const MQRC_JSSE_ERROR = 2397
Global Const MQRC_SSL_PEER_NAME_MISMATCH = 2398
Global Const MQRC_SSL_PEER_NAME_ERROR = 2399
Global Const MQRC_UNSUPPORTED_CIPHER_SUITE = 2400
Global Const MQRC_SSL_CERTIFICATE_REVOKED = 2401
Global Const MQRC_SSL_CERT_STORE_ERROR = 2402
Global Const MQRC_CLIENT_EXIT_LOAD_ERROR = 2406
Global Const MQRC_CLIENT_EXIT_ERROR = 2407
Global Const MQRC_SSL_KEY_RESET_ERROR = 2409
Global Const MQRC_UNKNOWN_COMPONENT_NAME = 2410
Global Const MQRC_LOGGER_STATUS = 2411
Global Const MQRC_COMMAND_MQSC = 2412
Global Const MQRC_COMMAND_PCF = 2413
Global Const MQRC_CFIF_ERROR = 2414
Global Const MQRC_CFSF_ERROR = 2415
Global Const MQRC_CFGR_ERROR = 2416
Global Const MQRC_MSG_NOT_ALLOWED_IN_GROUP = 2417
Global Const MQRC_FILTER_OPERATOR_ERROR = 2418
Global Const MQRC_NESTED_SELECTOR_ERROR = 2419
Global Const MQRC_EPH_ERROR = 2420
Global Const MQRC_RFH_FORMAT_ERROR = 2421
Global Const MQRC_CFBF_ERROR = 2422
Global Const MQRC_CLIENT_CHANNEL_CONFLICT = 2423
Global Const MQRC_REOPEN_EXCL_INPUT_ERROR = 6100
Global Const MQRC_REOPEN_INQUIRE_ERROR = 6101
Global Const MQRC_REOPEN_SAVED_CONTEXT_ERR = 6102
Global Const MQRC_REOPEN_TEMPORARY_Q_ERROR = 6103
Global Const MQRC_ATTRIBUTE_LOCKED = 6104
Global Const MQRC_CURSOR_NOT_VALID = 6105
Global Const MQRC_ENCODING_ERROR = 6106
Global Const MQRC_STRUC_ID_ERROR = 6107
Global Const MQRC_NULL_POINTER = 6108
Global Const MQRC_NO_CONNECTION_REFERENCE = 6109
Global Const MQRC_NO_BUFFER = 6110
Global Const MQRC_BINARY_DATA_LENGTH_ERROR = 6111
Global Const MQRC_BUFFER_NOT_AUTOMATIC = 6112
Global Const MQRC_INSUFFICIENT_BUFFER = 6113
Global Const MQRC_INSUFFICIENT_DATA = 6114
Global Const MQRC_DATA_TRUNCATED = 6115
Global Const MQRC_ZERO_LENGTH = 6116
Global Const MQRC_NEGATIVE_LENGTH = 6117
Global Const MQRC_NEGATIVE_OFFSET = 6118
Global Const MQRC_INCONSISTENT_FORMAT = 6119
Global Const MQRC_INCONSISTENT_OBJECT_STATE = 6120
Global Const MQRC_CONTEXT_OBJECT_NOT_VALID = 6121
Global Const MQRC_CONTEXT_OPEN_ERROR = 6122
Global Const MQRC_STRUC_LENGTH_ERROR = 6123
Global Const MQRC_NOT_CONNECTED = 6124
Global Const MQRC_NOT_OPEN = 6125
Global Const MQRC_DISTRIBUTION_LIST_EMPTY = 6126
Global Const MQRC_INCONSISTENT_OPEN_OPTIONS = 6127
Global Const MQRC_WRONG_VERSION = 6128
Global Const MQRC_REFERENCE_ERROR = 6129

'****************************************************************'
'*  Values Related to Queue Attributes                          *'
'****************************************************************'

'Queue Types'
Global Const MQQT_LOCAL = 1
Global Const MQQT_MODEL = 2
Global Const MQQT_ALIAS = 3
Global Const MQQT_REMOTE = 6
Global Const MQQT_CLUSTER = 7

'Cluster Queue Types'
Global Const MQCQT_LOCAL_Q = 1
Global Const MQCQT_ALIAS_Q = 2
Global Const MQCQT_REMOTE_Q = 3
Global Const MQCQT_Q_MGR_ALIAS = 4

'Extended Queue Types'
Global Const MQQT_ALL = 1001

'Queue Definition Types'
Global Const MQQDT_PREDEFINED = 1
Global Const MQQDT_PERMANENT_DYNAMIC = 2
Global Const MQQDT_TEMPORARY_DYNAMIC = 3
Global Const MQQDT_SHARED_DYNAMIC = 4

'Inhibit Get Values'
Global Const MQQA_GET_INHIBITED = 1
Global Const MQQA_GET_ALLOWED = 0

'Inhibit Put Values'
Global Const MQQA_PUT_INHIBITED = 1
Global Const MQQA_PUT_ALLOWED = 0

'Queue Shareability'
Global Const MQQA_SHAREABLE = 1
Global Const MQQA_NOT_SHAREABLE = 0

'Back-Out Hardening'
Global Const MQQA_BACKOUT_HARDENED = 1
Global Const MQQA_BACKOUT_NOT_HARDENED = 0

'Message Delivery Sequence'
Global Const MQMDS_PRIORITY = 0
Global Const MQMDS_FIFO = 1

'Nonpersistent Message Class'
Global Const MQNPM_CLASS_NORMAL = 0
Global Const MQNPM_CLASS_HIGH = 10

'Trigger Controls'
Global Const MQTC_OFF = 0
Global Const MQTC_ON = 1

'Trigger Types'
Global Const MQTT_NONE = 0
Global Const MQTT_FIRST = 1
Global Const MQTT_EVERY = 2
Global Const MQTT_DEPTH = 3

'Trigger Restart'
Global Const MQTRIGGER_RESTART_NO = 0
Global Const MQTRIGGER_RESTART_YES = 1

'Queue Usages'
Global Const MQUS_NORMAL = 0
Global Const MQUS_TRANSMISSION = 1

'Distribution Lists'
Global Const MQDL_SUPPORTED = 1
Global Const MQDL_NOT_SUPPORTED = 0

'Index Types'
Global Const MQIT_NONE = 0
Global Const MQIT_MSG_ID = 1
Global Const MQIT_CORREL_ID = 2
Global Const MQIT_MSG_TOKEN = 4
Global Const MQIT_GROUP_ID = 5

'Default Binds'
Global Const MQBND_BIND_ON_OPEN = 0
Global Const MQBND_BIND_NOT_FIXED = 1

'Queue Sharing Group Dispositions'
Global Const MQQSGD_ALL = -1
Global Const MQQSGD_Q_MGR = 0
Global Const MQQSGD_COPY = 1
Global Const MQQSGD_SHARED = 2
Global Const MQQSGD_GROUP = 3
Global Const MQQSGD_PRIVATE = 4
Global Const MQQSGD_LIVE = 6

'Reorganization Controls'
Global Const MQREORG_DISABLED = 0
Global Const MQREORG_ENABLED = 1

'****************************************************************'
'*  Values Related to Namelist Attributes                       *'
'****************************************************************'

'Name Count'
Global Const MQNC_MAX_NAMELIST_NAME_COUNT = 256

'Namelist Types'
Global Const MQNT_NONE = 0
Global Const MQNT_Q = 1
Global Const MQNT_CLUSTER = 2
Global Const MQNT_AUTH_INFO = 4
Global Const MQNT_ALL = 1001

'****************************************************************'
'*  Values Related to Process-Definition Attributes             *'
'****************************************************************'

'Application Type'
'See values for "Put Application Type" under MQMD'
'****************************************************************'
'*  Values Related to Authentication-Information Attributes     *'
'****************************************************************'

'Authentication Information Type'
'See values for "Authentication Information Type" under MQAIR'
'****************************************************************'
'*  Values Related to CF-Structure Attributes                   *'
'****************************************************************'

'CF Recoverability'
Global Const MQCFR_YES = 1
Global Const MQCFR_NO = 0

'****************************************************************'
'*  Values Related to Service Attributes                        *'
'****************************************************************'

'Service Types'
Global Const MQSVC_TYPE_COMMAND = 0
Global Const MQSVC_TYPE_SERVER = 1

'****************************************************************'
'*  Values Related to Queue-Manager Attributes                  *'
'****************************************************************'

'Adopt New MCA Checks'
Global Const MQADOPT_CHECK_ALL = -1
Global Const MQADOPT_CHECK_NONE = 0
Global Const MQADOPT_CHECK_Q_MGR_NAME = 1
Global Const MQADOPT_CHECK_NET_ADDR = 2

'Adopt New MCA Types'
Global Const MQADOPT_TYPE_NO = 0
Global Const MQADOPT_TYPE_SDR = 1
Global Const MQADOPT_TYPE_SVR = 2
Global Const MQADOPT_TYPE_RCVR = 3
Global Const MQADOPT_TYPE_ALL = 5
Global Const MQADOPT_TYPE_CLUSRCVR = 8

'Autostart'
Global Const MQAUTO_START_NO = 0
Global Const MQAUTO_START_YES = 1

'Channel Auto Definition'
Global Const MQCHAD_DISABLED = 0
Global Const MQCHAD_ENABLED = 1

'Cluster Workload'
Global Const MQCLWL_USEQ_LOCAL = 0
Global Const MQCLWL_USEQ_ANY = 1
Global Const MQCLWL_USEQ_AS_Q_MGR = -3

'Command Levels'
Global Const MQCMDL_LEVEL_1 = 100
Global Const MQCMDL_LEVEL_101 = 101
Global Const MQCMDL_LEVEL_110 = 110
Global Const MQCMDL_LEVEL_114 = 114
Global Const MQCMDL_LEVEL_120 = 120
Global Const MQCMDL_LEVEL_200 = 200
Global Const MQCMDL_LEVEL_201 = 201
Global Const MQCMDL_LEVEL_210 = 210
Global Const MQCMDL_LEVEL_211 = 211
Global Const MQCMDL_LEVEL_220 = 220
Global Const MQCMDL_LEVEL_221 = 221
Global Const MQCMDL_LEVEL_230 = 230
Global Const MQCMDL_LEVEL_320 = 320
Global Const MQCMDL_LEVEL_420 = 420
Global Const MQCMDL_LEVEL_500 = 500
Global Const MQCMDL_LEVEL_510 = 510
Global Const MQCMDL_LEVEL_520 = 520
Global Const MQCMDL_LEVEL_530 = 530
Global Const MQCMDL_LEVEL_531 = 531
Global Const MQCMDL_LEVEL_600 = 600

'Command Server Options'
Global Const MQCSRV_CONVERT_NO = 0
Global Const MQCSRV_CONVERT_YES = 1
Global Const MQCSRV_DLQ_NO = 0
Global Const MQCSRV_DLQ_YES = 1

'Distribution Lists'
'See values for "Distribution Lists" under Queue Attributes'
'DNS WLM'
Global Const MQDNSWLM_NO = 0
Global Const MQDNSWLM_YES = 1

'Expiration Scan Interval'
Global Const MQEXPI_OFF = 0

'Intra-Group Queuing'
Global Const MQIGQ_DISABLED = 0
Global Const MQIGQ_ENABLED = 1

'Intra-Group Queuing Put Authority'
Global Const MQIGQPA_DEFAULT = 1
Global Const MQIGQPA_CONTEXT = 2
Global Const MQIGQPA_ONLY_IGQ = 3
Global Const MQIGQPA_ALTERNATE_OR_IGQ = 4

'IP Address Versions'
Global Const MQIPADDR_IPV4 = 0
Global Const MQIPADDR_IPV6 = 1

'Monitoring Values'
Global Const MQMON_NOT_AVAILABLE = -1
Global Const MQMON_NONE = -1
Global Const MQMON_Q_MGR = -3
Global Const MQMON_OFF = 0
Global Const MQMON_ON = 1
Global Const MQMON_DISABLED = 0
Global Const MQMON_ENABLED = 1
Global Const MQMON_LOW = 17
Global Const MQMON_MEDIUM = 33
Global Const MQMON_HIGH = 65

'Platforms'
Global Const MQPL_MVS = 1
Global Const MQPL_OS390 = 1
Global Const MQPL_ZOS = 1
Global Const MQPL_OS2 = 2
Global Const MQPL_AIX = 3
Global Const MQPL_UNIX = 3
Global Const MQPL_OS400 = 4
Global Const MQPL_WINDOWS = 5
Global Const MQPL_WINDOWS_NT = 11
Global Const MQPL_VMS = 12
Global Const MQPL_NSK = 13
Global Const MQPL_NSS = 13
Global Const MQPL_VSE = 27

'Control Options'
Global Const MQQMOPT_DISABLED = 0
Global Const MQQMOPT_ENABLED = 1
Global Const MQQMOPT_REPLY = 2

'Receive Timeout Types'
Global Const MQRCVTIME_MULTIPLY = 0
Global Const MQRCVTIME_ADD = 1
Global Const MQRCVTIME_EQUAL = 2

'Recording Options'
Global Const MQRECORDING_DISABLED = 0
Global Const MQRECORDING_Q = 1
Global Const MQRECORDING_MSG = 2

'Shared Queue Queue Manager Name'
Global Const MQSQQM_USE = 0
Global Const MQSQQM_IGNORE = 1

'SSL FIPS Requirements'
Global Const MQSSL_FIPS_NO = 0
Global Const MQSSL_FIPS_YES = 1

'Syncpoint Availability'
Global Const MQSP_AVAILABLE = 1
Global Const MQSP_NOT_AVAILABLE = 0

'Service Controls'
Global Const MQSVC_CONTROL_Q_MGR = 0
Global Const MQSVC_CONTROL_Q_MGR_START = 1
Global Const MQSVC_CONTROL_MANUAL = 2

'Service Status'
Global Const MQSVC_STATUS_STOPPED = 0
Global Const MQSVC_STATUS_STARTING = 1
Global Const MQSVC_STATUS_RUNNING = 2
Global Const MQSVC_STATUS_STOPPING = 3
Global Const MQSVC_STATUS_RETRYING = 4

'TCP Keepalive'
Global Const MQTCPKEEP_NO = 0
Global Const MQTCPKEEP_YES = 1

'TCP Stack Types'
Global Const MQTCPSTACK_SINGLE = 0
Global Const MQTCPSTACK_MULTIPLE = 1

'Channel Initiator Trace Autostart'
Global Const MQTRAXSTR_NO = 0
Global Const MQTRAXSTR_YES = 1

'*********************************************************************'
'*  Byte and Pointer Datatypes                                       *'
'*********************************************************************'

Type MQBYTE4
  MQByte(0 To 3) As Byte '4-byte binary string'
End Type

'Default Instance of MQBYTE4 Structure'
Global MQBYTE4_DEFAULT As MQBYTE4

Type MQBYTE8
  MQByte(0 To 7) As Byte '8-byte binary string'
End Type

'Default Instance of MQBYTE8 Structure'
Global MQBYTE8_DEFAULT As MQBYTE8

Type MQBYTE16
  MQByte(0 To 15) As Byte '16-byte binary string'
End Type

'Default Instance of MQBYTE16 Structure'
Global MQBYTE16_DEFAULT As MQBYTE16

Type MQBYTE24
  MQByte(0 To 23) As Byte '24-byte binary string'
End Type

'Default Instance of MQBYTE24 Structure'
Global MQBYTE24_DEFAULT As MQBYTE24

Type MQBYTE32
  MQByte(0 To 31) As Byte '32-byte binary string'
End Type

'Default Instance of MQBYTE32 Structure'
Global MQBYTE32_DEFAULT As MQBYTE32

Type MQBYTE40
  MQByte(0 To 39) As Byte '40-byte binary string'
End Type

'Default Instance of MQBYTE40 Structure'
Global MQBYTE40_DEFAULT As MQBYTE40

Type MQBYTE48
  MQByte(0 To 47) As Byte '48-byte binary string'
End Type

'Default Instance of MQBYTE48 Structure'
Global MQBYTE48_DEFAULT As MQBYTE48

Type MQBYTE128
  MQByte(0 To 127) As Byte '128-byte binary string'
End Type

'Default Instance of MQBYTE128 Structure'
Global MQBYTE128_DEFAULT As MQBYTE128

'Note: the MQPTR datatype is used as a placeholder in structures'
Type MQPTR
  MQByte(0 To 3) As Byte '4-byte pointer'
End Type

'Default Instance of MQPTR Structure'
Global MQPTR_DEFAULT As MQPTR

'****************************************************************'
'*  MQAIR Structure -- Authentication Information Record        *'
'****************************************************************'

Type MQAIR
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  AuthInfoType As Long 'Type of authentication information'
  AuthInfoConnName As String * 264 'Connection name of CRL LDAP server'
  LDAPUserNamePtr As MQPTR 'Address of LDAP user name'
  LDAPUserNameOffset As Long 'Offset of LDAP user name from start of MQAIR structure'
  LDAPUserNameLength As Long 'Length of LDAP user name'
  LDAPPassword As String * 32 'Password to access LDAP server'
End Type

'Default Instance of MQAIR Structure'
Global MQAIR_DEFAULT As MQAIR


'****************************************************************'
'*  MQBO Structure -- Begin Options                             *'
'****************************************************************'

Type MQBO
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Options As Long 'Options that control the action of MQBEGIN'
End Type

'Default Instance of MQBO Structure'
Global MQBO_DEFAULT As MQBO


'****************************************************************'
'*  MQCIH Structure -- CICS Information Header                  *'
'****************************************************************'

Type MQCIH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Length of MQCIH structure'
  Encoding As Long 'Reserved'
  CodedCharSetId As Long 'Reserved'
  Format As String * 8 'MQ format name of data that follows MQCIH'
  Flags As Long 'Flags'
  ReturnCode As Long 'Return code from bridge'
  CompCode As Long 'MQ completion code or CICS EIBRESP'
  Reason As Long 'MQ reason or feedback code, or CICS EIBRESP2'
  UOWControl As Long 'Unit-of-work control'
  GetWaitInterval As Long 'Wait interval for MQGET call issued by bridge task'
  LinkType As Long 'Link type'
  OutputDataLength As Long 'Output COMMAREA data length'
  FacilityKeepTime As Long 'Bridge facility release time'
  ADSDescriptor As Long 'Send/receive ADS descriptor'
  ConversationalTask As Long 'Whether task can be conversational'
  TaskEndStatus As Long 'Status at end of task'
  Facility As MQBYTE8 'Bridge facility token'
  Function As String * 4 'MQ call name or CICS EIBFN function'
  AbendCode As String * 4 'Abend code'
  Authenticator As String * 8 'Password or passticket'
  Reserved1 As String * 8 'Reserved'
  ReplyToFormat As String * 8 'MQ format name of reply message'
  RemoteSysId As String * 4 'Remote CICS system id to use'
  RemoteTransId As String * 4 'CICS RTRANSID to use'
  TransactionId As String * 4 'Transaction to attach'
  FacilityLike As String * 4 'Terminal emulated attributes'
  AttentionId As String * 4 'AID key'
  StartCode As String * 4 'Transaction start code'
  CancelCode As String * 4 'Abend transaction code'
  NextTransactionId As String * 4 'Next transaction to attach'
  Reserved2 As String * 8 'Reserved'
  Reserved3 As String * 8 'Reserved'
  CursorPosition As Long 'Cursor position'
  ErrorOffset As Long 'Offset of error in message'
  InputItem As Long 'Reserved'
  Reserved4 As Long 'Reserved'
End Type

'Default Instance of MQCIH Structure'
Global MQCIH_DEFAULT As MQCIH


'****************************************************************'
'*  MQSCO Structure -- SSL Configuration Options                *'
'****************************************************************'

Type MQSCO
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  KeyRepository As String * 256 'Location of SSL key repository'
  CryptoHardware As String * 256 'Cryptographic hardware configuration string'
  AuthInfoRecCount As Long 'Number of MQAIR records present'
  AuthInfoRecOffset As Long 'Offset of first MQAIR record from start of MQSCO structure'
  AuthInfoRecPtr As MQPTR 'Address of first MQAIR record'
  KeyResetCount As Long 'Number of unencrypted bytes sent/received before secret key is reset'
  FipsRequired As Long 'Mandatory FIPS CipherSpecs?'
End Type

'Default Instance of MQSCO Structure'
Global MQSCO_DEFAULT As MQSCO


'****************************************************************'
'*  MQCSP Structure -- Security Parameters                      *'
'****************************************************************'

Type MQCSP
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  AuthenticationType As Long 'Type of authentication'
  Reserved1 As MQBYTE4 'Reserved'
  CSPUserIdPtr As MQPTR 'Address of user ID'
  CSPUserIdOffset As Long 'Offset of user ID'
  CSPUserIdLength As Long 'Length of user ID'
  Reserved2 As MQBYTE8 'Reserved'
  CSPPasswordPtr As MQPTR 'Address of password'
  CSPPasswordOffset As Long 'Offset of password'
  CSPPasswordLength As Long 'Length of password'
End Type

'Default Instance of MQCSP Structure'
Global MQCSP_DEFAULT As MQCSP


'****************************************************************'
'*  MQCNO Structure -- Connect Options                          *'
'****************************************************************'

Type MQCNO
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Options As Long 'Options that control the action of MQCONNX'
  ClientConnOffset As Long 'Offset of MQCD structure for client connection'
  ClientConnPtr As MQPTR 'Address of MQCD structure for client connection'
  ConnTag As MQBYTE128 'Queue-manager connection tag'
  SSLConfigPtr As MQPTR 'Address of MQSCO structure for client connection'
  SSLConfigOffset As Long 'Offset of MQSCO structure for client connection'
  ConnectionId As MQBYTE24 'Unique Connection Identifier'
  SecurityParmsOffset As Long 'Offset of MQCSP structure'
  SecurityParmsPtr As MQPTR 'Address of MQCSP structure'
End Type

'Default Instance of MQCNO Structure'
Global MQCNO_DEFAULT As MQCNO


'****************************************************************'
'*  MQDH Structure -- Distribution Header                       *'
'****************************************************************'

Type MQDH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Length of MQDH structure plus following MQOR and MQPMR records'
  Encoding As Long 'Numeric encoding of data that follows the MQOR and MQPMR records'
  CodedCharSetId As Long 'Character set identifier of data that follows the MQOR and MQPMR records'
  Format As String * 8 'Format name of data that follows the MQOR and MQPMR records'
  Flags As Long 'General flags'
  PutMsgRecFields As Long 'Flags indicating which MQPMR fields are present'
  RecsPresent As Long 'Number of MQOR records present'
  ObjectRecOffset As Long 'Offset of first MQOR record from start of MQDH'
  PutMsgRecOffset As Long 'Offset of first MQPMR record from start of MQDH'
End Type

'Default Instance of MQDH Structure'
Global MQDH_DEFAULT As MQDH


'****************************************************************'
'*  MQDLH Structure -- Dead Letter Header                       *'
'****************************************************************'

Type MQDLH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Reason As Long 'Reason message arrived on dead-letter (undelivered-message) queue'
  DestQName As String * 48 'Name of original destination queue'
  DestQMgrName As String * 48 'Name of original destination queue manager'
  Encoding As Long 'Numeric encoding of data that follows MQDLH'
  CodedCharSetId As Long 'Character set identifier of data that follows MQDLH'
  Format As String * 8 'Format name of data that follows MQDLH'
  PutApplType As Long 'Type of application that put message on dead-letter (undelivered-message) queue'
  PutApplName As String * 28 'Name of application that put message on dead-letter (undelivered-message) queue'
  PutDate As String * 8 'Date when message was put on dead-letter (undelivered-message) queue'
  PutTime As String * 8 'Time when message was put on the dead-letter (undelivered-message) queue'
End Type

'Default Instance of MQDLH Structure'
Global MQDLH_DEFAULT As MQDLH


'****************************************************************'
'*  MQGMO Structure -- Get Message Options                      *'
'****************************************************************'

Type MQGMO
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Options As Long 'Options that control the action of MQGET'
  WaitInterval As Long 'Wait interval'
  Signal1 As Long 'Signal'
  Signal2 As Long 'Signal identifier'
  ResolvedQName As String * 48 'Resolved name of destination queue'
  MatchOptions As Long 'Options controlling selection criteria used for MQGET'
  GroupStatus As String * 1 'Flag indicating whether message retrieved is in a group'
  SegmentStatus As String * 1 'Flag indicating whether message retrieved is a segment of a logical message'
  Segmentation As String * 1 'Flag indicating whether further segmentation is allowed for the message retrieved'
  Reserved1 As String * 1 'Reserved'
  MsgToken As MQBYTE16 'Message token'
  ReturnedLength As Long 'Length of message data returned (bytes)'
End Type

'Default Instance of MQGMO Structure'
Global MQGMO_DEFAULT As MQGMO


'****************************************************************'
'*  MQIIH Structure -- IMS Information Header                   *'
'****************************************************************'

Type MQIIH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Length of MQIIH structure'
  Encoding As Long 'Reserved'
  CodedCharSetId As Long 'Reserved'
  Format As String * 8 'MQ format name of data that follows MQIIH'
  Flags As Long 'Flags'
  LTermOverride As String * 8 'Logical terminal override'
  MFSMapName As String * 8 'Message format services map name'
  ReplyToFormat As String * 8 'MQ format name of reply message'
  Authenticator As String * 8 'RACF password or passticket'
  TranInstanceId As MQBYTE16 'Transaction instance identifier'
  TranState As String * 1 'Transaction state'
  CommitMode As String * 1 'Commit mode'
  SecurityScope As String * 1 'Security scope'
  Reserved As String * 1 'Reserved'
End Type

'Default Instance of MQIIH Structure'
Global MQIIH_DEFAULT As MQIIH


'****************************************************************'
'*  MQMD Structure -- Message Descriptor                        *'
'****************************************************************'

Type MQMD
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Report As Long 'Options for report messages'
  MsgType As Long 'Message type'
  Expiry As Long 'Message lifetime'
  Feedback As Long 'Feedback or reason code'
  Encoding As Long 'Numeric encoding of message data'
  CodedCharSetId As Long 'Character set identifier of message data'
  Format As String * 8 'Format name of message data'
  Priority As Long 'Message priority'
  Persistence As Long 'Message persistence'
  MsgId As MQBYTE24 'Message identifier'
  CorrelId As MQBYTE24 'Correlation identifier'
  BackoutCount As Long 'Backout counter'
  ReplyToQ As String * 48 'Name of reply queue'
  ReplyToQMgr As String * 48 'Name of reply queue manager'
  UserIdentifier As String * 12 'User identifier'
  AccountingToken As MQBYTE32 'Accounting token'
  ApplIdentityData As String * 32 'Application data relating to identity'
  PutApplType As Long 'Type of application that put the message'
  PutApplName As String * 28 'Name of application that put the message'
  PutDate As String * 8 'Date when message was put'
  PutTime As String * 8 'Time when message was put'
  ApplOriginData As String * 4 'Application data relating to origin'
  GroupId As MQBYTE24 'Group identifier'
  MsgSeqNumber As Long 'Sequence number of logical message within group'
  Offset As Long 'Offset of data in physical message from start of logical message'
  MsgFlags As Long 'Message flags'
  OriginalLength As Long 'Length of original message'
End Type

'Default Instance of MQMD Structure'
Global MQMD_DEFAULT As MQMD


'****************************************************************'
'*  MQMDE Structure -- Message Descriptor Extension             *'
'****************************************************************'

Type MQMDE
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Length of MQMDE structure'
  Encoding As Long 'Numeric encoding of data that follows MQMDE'
  CodedCharSetId As Long 'Character-set identifier of data that follows MQMDE'
  Format As String * 8 'Format name of data that follows MQMDE'
  Flags As Long 'General flags'
  GroupId As MQBYTE24 'Group identifier'
  MsgSeqNumber As Long 'Sequence number of logical message within group'
  Offset As Long 'Offset of data in physical message from start of logical message'
  MsgFlags As Long 'Message flags'
  OriginalLength As Long 'Length of original message'
End Type

'Default Instance of MQMDE Structure'
Global MQMDE_DEFAULT As MQMDE


'****************************************************************'
'*  MQMD1 Structure -- Version-1 Message Descriptor             *'
'****************************************************************'

Type MQMD1
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Report As Long 'Report options'
  MsgType As Long 'Message type'
  Expiry As Long 'Expiry time'
  Feedback As Long 'Feedback or reason code'
  Encoding As Long 'Numeric encoding of message data'
  CodedCharSetId As Long 'Character set identifier of message data'
  Format As String * 8 'Format name of message data'
  Priority As Long 'Message priority'
  Persistence As Long 'Message persistence'
  MsgId As MQBYTE24 'Message identifier'
  CorrelId As MQBYTE24 'Correlation identifier'
  BackoutCount As Long 'Backout counter'
  ReplyToQ As String * 48 'Name of reply-to queue'
  ReplyToQMgr As String * 48 'Name of reply queue manager'
  UserIdentifier As String * 12 'User identifier'
  AccountingToken As MQBYTE32 'Accounting token'
  ApplIdentityData As String * 32 'Application data relating to identity'
  PutApplType As Long 'Type of application that put the message'
  PutApplName As String * 28 'Name of application that put the message'
  PutDate As String * 8 'Date when message was put'
  PutTime As String * 8 'Time when message was put'
  ApplOriginData As String * 4 'Application data relating to origin'
End Type

'Default Instance of MQMD1 Structure'
Global MQMD1_DEFAULT As MQMD1


'****************************************************************'
'*  MQMD2 Structure -- Version-2 Message Descriptor             *'
'****************************************************************'

Type MQMD2
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Report As Long 'Report options'
  MsgType As Long 'Message type'
  Expiry As Long 'Expiry time'
  Feedback As Long 'Feedback or reason code'
  Encoding As Long 'Numeric encoding of message data'
  CodedCharSetId As Long 'Character set identifier of message data'
  Format As String * 8 'Format name of message data'
  Priority As Long 'Message priority'
  Persistence As Long 'Message persistence'
  MsgId As MQBYTE24 'Message identifier'
  CorrelId As MQBYTE24 'Correlation identifier'
  BackoutCount As Long 'Backout counter'
  ReplyToQ As String * 48 'Name of reply-to queue'
  ReplyToQMgr As String * 48 'Name of reply queue manager'
  UserIdentifier As String * 12 'User identifier'
  AccountingToken As MQBYTE32 'Accounting token'
  ApplIdentityData As String * 32 'Application data relating to identity'
  PutApplType As Long 'Type of application that put the message'
  PutApplName As String * 28 'Name of application that put the message'
  PutDate As String * 8 'Date when message was put'
  PutTime As String * 8 'Time when message was put'
  ApplOriginData As String * 4 'Application data relating to origin'
  GroupId As MQBYTE24 'Group identifier'
  MsgSeqNumber As Long 'Sequence number of logical message within group'
  Offset As Long 'Offset of data in physical message from start of logical message'
  MsgFlags As Long 'Message flags'
  OriginalLength As Long 'Length of original message'
End Type

'Default Instance of MQMD2 Structure'
Global MQMD2_DEFAULT As MQMD2


'****************************************************************'
'*  MQOD Structure -- Object Descriptor                         *'
'****************************************************************'

Type MQOD
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  ObjectType As Long 'Object type'
  ObjectName As String * 48 'Object name'
  ObjectQMgrName As String * 48 'Object queue manager name'
  DynamicQName As String * 48 'Dynamic queue name'
  AlternateUserId As String * 12 'Alternate user identifier'
  RecsPresent As Long 'Number of object records present'
  KnownDestCount As Long 'Number of local queues opened successfully'
  UnknownDestCount As Long 'Number of remote queues opened successfully'
  InvalidDestCount As Long 'Number of queues that failed to open'
  ObjectRecOffset As Long 'Offset of first object record from start of MQOD'
  ResponseRecOffset As Long 'Offset of first response record from start of MQOD'
  ObjectRecPtr As MQPTR 'Address of first object record'
  ResponseRecPtr As MQPTR 'Address of first response record'
  AlternateSecurityId As MQBYTE40 'Alternate security identifier'
  ResolvedQName As String * 48 'Resolved queue name'
  ResolvedQMgrName As String * 48 'Resolved queue manager name'
End Type

'Default Instance of MQOD Structure'
Global MQOD_DEFAULT As MQOD


'****************************************************************'
'*  MQOR Structure -- Object Record                             *'
'****************************************************************'

Type MQOR
  ObjectName As String * 48 'Object name'
  ObjectQMgrName As String * 48 'Object queue manager name'
End Type

'Default Instance of MQOR Structure'
Global MQOR_DEFAULT As MQOR


'****************************************************************'
'*  MQPMO Structure -- Put Message Options                      *'
'****************************************************************'

Type MQPMO
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  Options As Long 'Options that control the action of MQPUT and MQPUT1'
  Timeout As Long 'Reserved'
  Context As Long 'Object handle of input queue'
  KnownDestCount As Long 'Number of messages sent successfully to local queues'
  UnknownDestCount As Long 'Number of messages sent successfully to remote queues'
  InvalidDestCount As Long 'Number of messages that could not be sent'
  ResolvedQName As String * 48 'Resolved name of destination queue'
  ResolvedQMgrName As String * 48 'Resolved name of destination queue manager'
  RecsPresent As Long 'Number of put message records or response records present'
  PutMsgRecFields As Long 'Flags indicating which MQPMR fields are present'
  PutMsgRecOffset As Long 'Offset of first put message record from start of MQPMO'
  ResponseRecOffset As Long 'Offset of first response record from start of MQPMO'
  PutMsgRecPtr As MQPTR 'Address of first put message record'
  ResponseRecPtr As MQPTR 'Address of first response record'
End Type

'Default Instance of MQPMO Structure'
Global MQPMO_DEFAULT As MQPMO


'****************************************************************'
'*  MQRFH Structure -- Rules and Formatting Header              *'
'****************************************************************'

Type MQRFH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Total length of MQRFH including NameValueString'
  Encoding As Long 'Numeric encoding of data that follows NameValueString'
  CodedCharSetId As Long 'Character set identifier of data that follows NameValueString'
  Format As String * 8 'Format name of data that follows NameValueString'
  Flags As Long 'Flags'
End Type

'Default Instance of MQRFH Structure'
Global MQRFH_DEFAULT As MQRFH


'****************************************************************'
'*  MQRFH2 Structure -- Rules and Formatting Header 2           *'
'****************************************************************'

Type MQRFH2
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Total length of MQRFH2 including all NameValueLength and NameValueData fields'
  Encoding As Long 'Numeric encoding of data that follows last NameValueData field'
  CodedCharSetId As Long 'Character set identifier of data that follows last NameValueData field'
  Format As String * 8 'Format name of data that follows last NameValueData field'
  Flags As Long 'Flags'
  NameValueCCSID As Long 'Character set identifier of NameValueData'
End Type

'Default Instance of MQRFH2 Structure'
Global MQRFH2_DEFAULT As MQRFH2


'****************************************************************'
'*  MQRMH Structure -- Reference Message Header                 *'
'****************************************************************'

Type MQRMH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Total length of MQRMH, including strings at end of fixed fields, but not the bulk data'
  Encoding As Long 'Numeric encoding of bulk data'
  CodedCharSetId As Long 'Character set identifier of bulk data'
  Format As String * 8 'Format name of bulk data'
  Flags As Long 'Reference message flags'
  ObjectType As String * 8 'Object type'
  ObjectInstanceId As MQBYTE24 'Object instance identifier'
  SrcEnvLength As Long 'Length of source environment data'
  SrcEnvOffset As Long 'Offset of source environment data'
  SrcNameLength As Long 'Length of source object name'
  SrcNameOffset As Long 'Offset of source object name'
  DestEnvLength As Long 'Length of destination environment data'
  DestEnvOffset As Long 'Offset of destination environment data'
  DestNameLength As Long 'Length of destination object name'
  DestNameOffset As Long 'Offset of destination object name'
  DataLogicalLength As Long 'Length of bulk data'
  DataLogicalOffset As Long 'Low offset of bulk data'
  DataLogicalOffset2 As Long 'High offset of bulk data'
End Type

'Default Instance of MQRMH Structure'
Global MQRMH_DEFAULT As MQRMH


'****************************************************************'
'*  MQRR Structure -- Response Record                           *'
'****************************************************************'

Type MQRR
  CompCode As Long 'Completion code for queue'
  Reason As Long 'Reason code for queue'
End Type

'Default Instance of MQRR Structure'
Global MQRR_DEFAULT As MQRR


'****************************************************************'
'*  MQTM Structure -- Trigger Message                           *'
'****************************************************************'

Type MQTM
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  QName As String * 48 'Name of triggered queue'
  ProcessName As String * 48 'Name of process object'
  TriggerData As String * 64 'Trigger data'
  ApplType As Long 'Application type'
  ApplId As String * 256 'Application identifier'
  EnvData As String * 128 'Environment data'
  UserData As String * 128 'User data'
End Type

'Default Instance of MQTM Structure'
Global MQTM_DEFAULT As MQTM


'****************************************************************'
'*  MQTMC2 Structure -- Trigger Message 2 (Character)           *'
'****************************************************************'

Type MQTMC2
  StrucId As String * 4 'Structure identifier'
  Version As String * 4 'Structure version number'
  QName As String * 48 'Name of triggered queue'
  ProcessName As String * 48 'Name of process object'
  TriggerData As String * 64 'Trigger data'
  ApplType As String * 4 'Application type'
  ApplId As String * 256 'Application identifier'
  EnvData As String * 128 'Environment data'
  UserData As String * 128 'User data'
  QMgrName As String * 48 'Queue manager name'
End Type

'Default Instance of MQTMC2 Structure'
Global MQTMC2_DEFAULT As MQTMC2


'****************************************************************'
'*  MQWIH Structure -- Work Information Header                  *'
'****************************************************************'

Type MQWIH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Length of MQWIH structure'
  Encoding As Long 'Numeric encoding of data that follows MQWIH'
  CodedCharSetId As Long 'Character-set identifier of data that follows MQWIH'
  Format As String * 8 'Format name of data that follows MQWIH'
  Flags As Long 'Flags'
  ServiceName As String * 32 'Service name'
  ServiceStep As String * 8 'Service step name'
  MsgToken As MQBYTE16 'Message token'
  Reserved As String * 32 'Reserved'
End Type

'Default Instance of MQWIH Structure'
Global MQWIH_DEFAULT As MQWIH


'****************************************************************'
'*  MQXQH Structure -- Transmission Queue Header                *'
'****************************************************************'

Type MQXQH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  RemoteQName As String * 48 'Name of destination queue'
  RemoteQMgrName As String * 48 'Name of destination queue manager'
  MsgDesc As MQMD1 'Original message descriptor'
End Type

'Default Instance of MQXQH Structure'
Global MQXQH_DEFAULT As MQXQH


'****************************************************************'
'*  Parameter usage in functions                                *'
'*    I:    input                                               *'
'*    IB:   input, data buffer                                  *'
'*    IL:   input, length of data buffer                        *'
'*    IO:   input and output                                    *'
'*    IOB:  input and output, data buffer                       *'
'*    IOL:  input and output, length of data buffer             *'
'*    O:    output                                              *'
'*    OB:   output, data buffer                                 *'
'*    OC:   output, completion code                             *'
'*    OR:   output, reason code                                 *'
'****************************************************************'
'****************************************************************'
'*  MQBACK Function -- Back Out Changes                         *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Global Const MQVBDLL = "MQM.DLL"

Declare Sub MQBACK Lib "MQM.DLL" Alias "MQBACKstd@12" _
 (ByVal Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Global Const MQVBDLL = "MQIC.DLL"

Declare Sub MQBACK Lib "MQIC.DLL" Alias "MQBACKstd@12" _
 (ByVal Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Global Const MQVBDLL = "MQICXA.DLL"

Declare Sub MQBACK Lib "MQICXA.DLL" Alias "MQBACKstd@12" _
 (ByVal Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
Global Const MQVBDLL = "NONE"
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQBEGIN Function -- Begin Unit of Work                      *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQBEGIN Lib "MQM.DLL" Alias "MQBEGINstd@16" _
 (ByVal Hconn As Long, _
  BeginOptions As MQBO, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQBEGIN Lib "MQIC.DLL" Alias "MQBEGINstd@16" _
 (ByVal Hconn As Long, _
  BeginOptions As MQBO, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQBEGIN Lib "MQICXA.DLL" Alias "MQBEGINstd@16" _
 (ByVal Hconn As Long, _
  BeginOptions As MQBO, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQCLOSE Function -- Close Object                            *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQCLOSE Lib "MQM.DLL" Alias "MQCLOSEstd@20" _
 (ByVal Hconn As Long, _
  Hobj As Long, _
  ByVal Options As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQCLOSE Lib "MQIC.DLL" Alias "MQCLOSEstd@20" _
 (ByVal Hconn As Long, _
  Hobj As Long, _
  ByVal Options As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQCLOSE Lib "MQICXA.DLL" Alias "MQCLOSEstd@20" _
 (ByVal Hconn As Long, _
  Hobj As Long, _
  ByVal Options As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQCMIT Function -- Commit Changes                           *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQCMIT Lib "MQM.DLL" Alias "MQCMITstd@12" _
 (ByVal Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQCMIT Lib "MQIC.DLL" Alias "MQCMITstd@12" _
 (ByVal Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQCMIT Lib "MQICXA.DLL" Alias "MQCMITstd@12" _
 (ByVal Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQCONN Function -- Connect Queue Manager                    *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQCONN Lib "MQM.DLL" Alias "MQCONNstd@16" _
 (ByVal QMgrName As String, _
  Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQCONN Lib "MQIC.DLL" Alias "MQCONNstd@16" _
 (ByVal QMgrName As String, _
  Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQCONN Lib "MQICXA.DLL" Alias "MQCONNstd@16" _
 (ByVal QMgrName As String, _
  Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQCONNX Function -- Connect Queue Manager (Extended)        *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQCONNXAny Lib "MQM.DLL" Alias "MQCONNXstd@20" _
 (ByVal QMgrName As String, _
  ConnectOpts As Any, _
  Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQCONNXAny Lib "MQIC.DLL" Alias "MQCONNXstd@20" _
 (ByVal QMgrName As String, _
  ConnectOpts As Any, _
  Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQCONNXAny Lib "MQICXA.DLL" Alias "MQCONNXstd@20" _
 (ByVal QMgrName As String, _
  ConnectOpts As Any, _
  Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQDISC Function -- Disconnect Queue Manager                 *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQDISC Lib "MQM.DLL" Alias "MQDISCstd@12" _
 (Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQDISC Lib "MQIC.DLL" Alias "MQDISCstd@12" _
 (Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQDISC Lib "MQICXA.DLL" Alias "MQDISCstd@12" _
 (Hconn As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQGET Function -- Get Message                               *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQGETAny Lib "MQM.DLL" Alias "MQGETstd@36" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQGETX Lib "MQM.DLL" Alias "MQGETstd@36" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQGETAny Lib "MQIC.DLL" Alias "MQGETstd@36" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQGETX Lib "MQIC.DLL" Alias "MQGETstd@36" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQGETAny Lib "MQICXA.DLL" Alias "MQGETstd@36" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQGETX Lib "MQICXA.DLL" Alias "MQGETstd@36" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If


'****************************************************************'
'*  MQINQ Function -- Inquire Object Attributes                 *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Private Declare Sub MQINQX Lib "MQM.DLL" Alias "MQINQstd@40" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Private Declare Sub MQINQX Lib "MQIC.DLL" Alias "MQINQstd@40" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Private Declare Sub MQINQX Lib "MQICXA.DLL" Alias "MQINQstd@40" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If


'****************************************************************'
'*  MQOPEN Function -- Open Object                              *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQOPEN Lib "MQM.DLL" Alias "MQOPENstd@24" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  ByVal Options As Long, _
  Hobj As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQOPEN Lib "MQIC.DLL" Alias "MQOPENstd@24" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  ByVal Options As Long, _
  Hobj As Long, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQOPEN Lib "MQICXA.DLL" Alias "MQOPENstd@24" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  ByVal Options As Long, _
  Hobj As Long, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If

'****************************************************************'
'*  MQPUT Function -- Put Message                               *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQPUTAny Lib "MQM.DLL" Alias "MQPUTstd@32" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQPUTX Lib "MQM.DLL" Alias "MQPUTstd@32" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQPUTAny Lib "MQIC.DLL" Alias "MQPUTstd@32" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQPUTX Lib "MQIC.DLL" Alias "MQPUTstd@32" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQPUTAny Lib "MQICXA.DLL" Alias "MQPUTstd@32" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQPUTX Lib "MQICXA.DLL" Alias "MQPUTstd@32" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If


'****************************************************************'
'*  MQPUT1 Function -- Put One Message                          *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Declare Sub MQPUT1Any Lib "MQM.DLL" Alias "MQPUT1std@32" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQPUT1X Lib "MQM.DLL" Alias "MQPUT1std@32" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Declare Sub MQPUT1Any Lib "MQIC.DLL" Alias "MQPUT1std@32" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQPUT1X Lib "MQIC.DLL" Alias "MQPUT1std@32" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Declare Sub MQPUT1Any Lib "MQICXA.DLL" Alias "MQPUT1std@32" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  Buffer As Any, _
  CompCode As Long, _
  Reason As Long)

Private Declare Sub MQPUT1X Lib "MQICXA.DLL" Alias "MQPUT1std@32" _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If


'****************************************************************'
'*  MQSET Function -- Set Object Attributes                     *'
'****************************************************************'

#If MqType = 1 Then 'MQ server'
 'Name of dynamic link library'
Private Declare Sub MQSETX Lib "MQM.DLL" Alias "MQSETstd@40" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  ByVal IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  ByVal CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 2 Then 'MQ client'
 'Name of dynamic link library'
Private Declare Sub MQSETX Lib "MQIC.DLL" Alias "MQSETstd@40" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  ByVal IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  ByVal CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)


#ElseIf MqType = 3 Then 'MQ extended transactional client'
 'Name of dynamic link library'
Private Declare Sub MQSETX Lib "MQICXA.DLL" Alias "MQSETstd@40" _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  ByVal IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  ByVal CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)


#Else
 'MqType not set or set wrong'
 'Please see the comments at the top of this file'
#End If


'*********************************************************************'
'*  Subroutines to Set Structures to Default Values                  *'
'*********************************************************************'










'*********************************************************************'
'*  MQ_SETDEFAULTS Subroutine -- Set Defaults                        *'
'*********************************************************************'

'****************************************************************'
'****************************************************************'

Sub MQAIR_DEFAULTS(Struc As MQAIR)
  Struc.StrucId = MQAIR_STRUC_ID
  Struc.Version = MQAIR_VERSION_1
  Struc.AuthInfoType = MQAIT_CRL_LDAP
  Struc.AuthInfoConnName = ""
  Dim TempLDAPUserNamePtr As MQPTR
  MQPTR_DEFAULTS TempLDAPUserNamePtr
  Struc.LDAPUserNamePtr = TempLDAPUserNamePtr
  Struc.LDAPUserNameOffset = 0
  Struc.LDAPUserNameLength = 0
  Struc.LDAPPassword = ""
End Sub

Sub MQBO_DEFAULTS(Struc As MQBO)
  Struc.StrucId = MQBO_STRUC_ID
  Struc.Version = MQBO_VERSION_1
  Struc.Options = MQBO_NONE
End Sub

Sub MQCIH_DEFAULTS(Struc As MQCIH)
  Struc.StrucId = MQCIH_STRUC_ID
  Struc.Version = MQCIH_VERSION_2
  Struc.StrucLength = MQCIH_LENGTH_2
  Struc.Encoding = 0
  Struc.CodedCharSetId = 0
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQCIH_NONE
  Struc.ReturnCode = MQCRC_OK
  Struc.CompCode = MQCC_OK
  Struc.Reason = MQRC_NONE
  Struc.UOWControl = MQCUOWC_ONLY
  Struc.GetWaitInterval = MQCGWI_DEFAULT
  Struc.LinkType = MQCLT_PROGRAM
  Struc.OutputDataLength = MQCODL_AS_INPUT
  Struc.FacilityKeepTime = 0
  Struc.ADSDescriptor = MQCADSD_NONE
  Struc.ConversationalTask = MQCCT_NO
  Struc.TaskEndStatus = MQCTES_NOSYNC
  Dim TempFacility As MQBYTE8
  MQBYTE8_DEFAULTS TempFacility
  Struc.Facility = TempFacility
  Struc.Function = MQCFUNC_NONE
  Struc.AbendCode = ""
  Struc.Authenticator = ""
  Struc.Reserved1 = ""
  Struc.ReplyToFormat = MQFMT_NONE
  Struc.RemoteSysId = ""
  Struc.RemoteTransId = ""
  Struc.TransactionId = ""
  Struc.FacilityLike = ""
  Struc.AttentionId = ""
  Struc.StartCode = MQCSC_NONE
  Struc.CancelCode = ""
  Struc.NextTransactionId = ""
  Struc.Reserved2 = ""
  Struc.Reserved3 = ""
  Struc.CursorPosition = 0
  Struc.ErrorOffset = 0
  Struc.InputItem = 0
  Struc.Reserved4 = 0
End Sub

Sub MQSCO_DEFAULTS(Struc As MQSCO)
  Struc.StrucId = MQSCO_STRUC_ID
  Struc.Version = MQSCO_VERSION_1
  Struc.KeyRepository = ""
  Struc.CryptoHardware = ""
  Struc.AuthInfoRecCount = 0
  Struc.AuthInfoRecOffset = 0
  Dim TempAuthInfoRecPtr As MQPTR
  MQPTR_DEFAULTS TempAuthInfoRecPtr
  Struc.AuthInfoRecPtr = TempAuthInfoRecPtr
  Struc.KeyResetCount = MQSCO_RESET_COUNT_DEFAULT
  Struc.FipsRequired = MQSSL_FIPS_NO
End Sub

Sub MQCSP_DEFAULTS(Struc As MQCSP)
  Struc.StrucId = MQCSP_STRUC_ID
  Struc.Version = MQCSP_VERSION_1
  Struc.AuthenticationType = MQCSP_AUTH_NONE
  Dim TempReserved1 As MQBYTE4
  MQBYTE4_DEFAULTS TempReserved1
  Struc.Reserved1 = TempReserved1
  Dim TempCSPUserIdPtr As MQPTR
  MQPTR_DEFAULTS TempCSPUserIdPtr
  Struc.CSPUserIdPtr = TempCSPUserIdPtr
  Struc.CSPUserIdOffset = 0
  Struc.CSPUserIdLength = 0
  Dim TempReserved2 As MQBYTE8
  MQBYTE8_DEFAULTS TempReserved2
  Struc.Reserved2 = TempReserved2
  Dim TempCSPPasswordPtr As MQPTR
  MQPTR_DEFAULTS TempCSPPasswordPtr
  Struc.CSPPasswordPtr = TempCSPPasswordPtr
  Struc.CSPPasswordOffset = 0
  Struc.CSPPasswordLength = 0
End Sub

Sub MQCNO_DEFAULTS(Struc As MQCNO)
  Struc.StrucId = MQCNO_STRUC_ID
  Struc.Version = MQCNO_VERSION_1
  Struc.Options = MQCNO_NONE
  Struc.ClientConnOffset = 0
  Dim TempClientConnPtr As MQPTR
  MQPTR_DEFAULTS TempClientConnPtr
  Struc.ClientConnPtr = TempClientConnPtr
  Dim TempConnTag As MQBYTE128
  MQBYTE128_DEFAULTS TempConnTag
  Struc.ConnTag = TempConnTag
  Dim TempSSLConfigPtr As MQPTR
  MQPTR_DEFAULTS TempSSLConfigPtr
  Struc.SSLConfigPtr = TempSSLConfigPtr
  Struc.SSLConfigOffset = 0
  Dim TempConnectionId As MQBYTE24
  MQBYTE24_DEFAULTS TempConnectionId
  Struc.ConnectionId = TempConnectionId
  Struc.SecurityParmsOffset = 0
  Dim TempSecurityParmsPtr As MQPTR
  MQPTR_DEFAULTS TempSecurityParmsPtr
  Struc.SecurityParmsPtr = TempSecurityParmsPtr
End Sub

Sub MQDH_DEFAULTS(Struc As MQDH)
  Struc.StrucId = MQDH_STRUC_ID
  Struc.Version = MQDH_VERSION_1
  Struc.StrucLength = 0
  Struc.Encoding = 0
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQDHF_NONE
  Struc.PutMsgRecFields = MQPMRF_NONE
  Struc.RecsPresent = 0
  Struc.ObjectRecOffset = 0
  Struc.PutMsgRecOffset = 0
End Sub

Sub MQDLH_DEFAULTS(Struc As MQDLH)
  Struc.StrucId = MQDLH_STRUC_ID
  Struc.Version = MQDLH_VERSION_1
  Struc.Reason = MQRC_NONE
  Struc.DestQName = ""
  Struc.DestQMgrName = ""
  Struc.Encoding = 0
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.PutApplType = 0
  Struc.PutApplName = ""
  Struc.PutDate = ""
  Struc.PutTime = ""
End Sub

Sub MQGMO_DEFAULTS(Struc As MQGMO)
  Struc.StrucId = MQGMO_STRUC_ID
  Struc.Version = MQGMO_VERSION_1
  Struc.Options = MQGMO_NO_WAIT
  Struc.WaitInterval = 0
  Struc.Signal1 = 0
  Struc.Signal2 = 0
  Struc.ResolvedQName = ""
  Struc.MatchOptions = MQMO_MATCH_MSG_ID + MQMO_MATCH_CORREL_ID
  Struc.GroupStatus = MQGS_NOT_IN_GROUP
  Struc.SegmentStatus = MQSS_NOT_A_SEGMENT
  Struc.Segmentation = MQSEG_INHIBITED
  Struc.Reserved1 = ""
  Dim TempMsgToken As MQBYTE16
  MQBYTE16_DEFAULTS TempMsgToken
  Struc.MsgToken = TempMsgToken
  Struc.ReturnedLength = MQRL_UNDEFINED
End Sub

Sub MQIIH_DEFAULTS(Struc As MQIIH)
  Struc.StrucId = MQIIH_STRUC_ID
  Struc.Version = MQIIH_VERSION_1
  Struc.StrucLength = MQIIH_LENGTH_1
  Struc.Encoding = 0
  Struc.CodedCharSetId = 0
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQIIH_NONE
  Struc.LTermOverride = ""
  Struc.MFSMapName = ""
  Struc.ReplyToFormat = MQFMT_NONE
  Struc.Authenticator = MQIAUT_NONE
  Dim TempTranInstanceId As MQBYTE16
  MQBYTE16_DEFAULTS TempTranInstanceId
  Struc.TranInstanceId = TempTranInstanceId
  Struc.TranState = MQITS_NOT_IN_CONVERSATION
  Struc.CommitMode = MQICM_COMMIT_THEN_SEND
  Struc.SecurityScope = MQISS_CHECK
  Struc.Reserved = ""
End Sub

Sub MQMD_DEFAULTS(Struc As MQMD)
  Struc.StrucId = MQMD_STRUC_ID
  Struc.Version = MQMD_VERSION_1
  Struc.Report = MQRO_NONE
  Struc.MsgType = MQMT_DATAGRAM
  Struc.Expiry = MQEI_UNLIMITED
  Struc.Feedback = MQFB_NONE
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_Q_MGR
  Struc.Format = MQFMT_NONE
  Struc.Priority = MQPRI_PRIORITY_AS_Q_DEF
  Struc.Persistence = MQPER_PERSISTENCE_AS_Q_DEF
  Dim TempMsgId As MQBYTE24
  MQBYTE24_DEFAULTS TempMsgId
  Struc.MsgId = TempMsgId
  Dim TempCorrelId As MQBYTE24
  MQBYTE24_DEFAULTS TempCorrelId
  Struc.CorrelId = TempCorrelId
  Struc.BackoutCount = 0
  Struc.ReplyToQ = ""
  Struc.ReplyToQMgr = ""
  Struc.UserIdentifier = ""
  Dim TempAccountingToken As MQBYTE32
  MQBYTE32_DEFAULTS TempAccountingToken
  Struc.AccountingToken = TempAccountingToken
  Struc.ApplIdentityData = ""
  Struc.PutApplType = MQAT_NO_CONTEXT
  Struc.PutApplName = ""
  Struc.PutDate = ""
  Struc.PutTime = ""
  Struc.ApplOriginData = ""
  Dim TempGroupId As MQBYTE24
  MQBYTE24_DEFAULTS TempGroupId
  Struc.GroupId = TempGroupId
  Struc.MsgSeqNumber = 1
  Struc.Offset = 0
  Struc.MsgFlags = MQMF_NONE
  Struc.OriginalLength = MQOL_UNDEFINED
End Sub

Sub MQMDE_DEFAULTS(Struc As MQMDE)
  Struc.StrucId = MQMDE_STRUC_ID
  Struc.Version = MQMDE_VERSION_2
  Struc.StrucLength = MQMDE_LENGTH_2
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQMDEF_NONE
  Dim TempGroupId As MQBYTE24
  MQBYTE24_DEFAULTS TempGroupId
  Struc.GroupId = TempGroupId
  Struc.MsgSeqNumber = 1
  Struc.Offset = 0
  Struc.MsgFlags = MQMF_NONE
  Struc.OriginalLength = MQOL_UNDEFINED
End Sub

Sub MQMD1_DEFAULTS(Struc As MQMD1)
  Struc.StrucId = MQMD_STRUC_ID
  Struc.Version = MQMD_VERSION_1
  Struc.Report = MQRO_NONE
  Struc.MsgType = MQMT_DATAGRAM
  Struc.Expiry = MQEI_UNLIMITED
  Struc.Feedback = MQFB_NONE
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_Q_MGR
  Struc.Format = MQFMT_NONE
  Struc.Priority = MQPRI_PRIORITY_AS_Q_DEF
  Struc.Persistence = MQPER_PERSISTENCE_AS_Q_DEF
  Dim TempMsgId As MQBYTE24
  MQBYTE24_DEFAULTS TempMsgId
  Struc.MsgId = TempMsgId
  Dim TempCorrelId As MQBYTE24
  MQBYTE24_DEFAULTS TempCorrelId
  Struc.CorrelId = TempCorrelId
  Struc.BackoutCount = 0
  Struc.ReplyToQ = ""
  Struc.ReplyToQMgr = ""
  Struc.UserIdentifier = ""
  Dim TempAccountingToken As MQBYTE32
  MQBYTE32_DEFAULTS TempAccountingToken
  Struc.AccountingToken = TempAccountingToken
  Struc.ApplIdentityData = ""
  Struc.PutApplType = MQAT_NO_CONTEXT
  Struc.PutApplName = ""
  Struc.PutDate = ""
  Struc.PutTime = ""
  Struc.ApplOriginData = ""
End Sub

Sub MQMD2_DEFAULTS(Struc As MQMD2)
  Struc.StrucId = MQMD_STRUC_ID
  Struc.Version = MQMD_VERSION_2
  Struc.Report = MQRO_NONE
  Struc.MsgType = MQMT_DATAGRAM
  Struc.Expiry = MQEI_UNLIMITED
  Struc.Feedback = MQFB_NONE
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_Q_MGR
  Struc.Format = MQFMT_NONE
  Struc.Priority = MQPRI_PRIORITY_AS_Q_DEF
  Struc.Persistence = MQPER_PERSISTENCE_AS_Q_DEF
  Dim TempMsgId As MQBYTE24
  MQBYTE24_DEFAULTS TempMsgId
  Struc.MsgId = TempMsgId
  Dim TempCorrelId As MQBYTE24
  MQBYTE24_DEFAULTS TempCorrelId
  Struc.CorrelId = TempCorrelId
  Struc.BackoutCount = 0
  Struc.ReplyToQ = ""
  Struc.ReplyToQMgr = ""
  Struc.UserIdentifier = ""
  Dim TempAccountingToken As MQBYTE32
  MQBYTE32_DEFAULTS TempAccountingToken
  Struc.AccountingToken = TempAccountingToken
  Struc.ApplIdentityData = ""
  Struc.PutApplType = MQAT_NO_CONTEXT
  Struc.PutApplName = ""
  Struc.PutDate = ""
  Struc.PutTime = ""
  Struc.ApplOriginData = ""
  Dim TempGroupId As MQBYTE24
  MQBYTE24_DEFAULTS TempGroupId
  Struc.GroupId = TempGroupId
  Struc.MsgSeqNumber = 1
  Struc.Offset = 0
  Struc.MsgFlags = MQMF_NONE
  Struc.OriginalLength = MQOL_UNDEFINED
End Sub

Sub MQOD_DEFAULTS(Struc As MQOD)
  Struc.StrucId = MQOD_STRUC_ID
  Struc.Version = MQOD_VERSION_1
  Struc.ObjectType = MQOT_Q
  Struc.ObjectName = ""
  Struc.ObjectQMgrName = ""
  Struc.DynamicQName = "AMQ.*"
  Struc.AlternateUserId = ""
  Struc.RecsPresent = 0
  Struc.KnownDestCount = 0
  Struc.UnknownDestCount = 0
  Struc.InvalidDestCount = 0
  Struc.ObjectRecOffset = 0
  Struc.ResponseRecOffset = 0
  Dim TempObjectRecPtr As MQPTR
  MQPTR_DEFAULTS TempObjectRecPtr
  Struc.ObjectRecPtr = TempObjectRecPtr
  Dim TempResponseRecPtr As MQPTR
  MQPTR_DEFAULTS TempResponseRecPtr
  Struc.ResponseRecPtr = TempResponseRecPtr
  Dim TempAlternateSecurityId As MQBYTE40
  MQBYTE40_DEFAULTS TempAlternateSecurityId
  Struc.AlternateSecurityId = TempAlternateSecurityId
  Struc.ResolvedQName = ""
  Struc.ResolvedQMgrName = ""
End Sub

Sub MQOR_DEFAULTS(Struc As MQOR)
  Struc.ObjectName = ""
  Struc.ObjectQMgrName = ""
End Sub

Sub MQPMO_DEFAULTS(Struc As MQPMO)
  Struc.StrucId = MQPMO_STRUC_ID
  Struc.Version = MQPMO_VERSION_1
  Struc.Options = MQPMO_NONE
  Struc.Timeout = -1
  Struc.Context = 0
  Struc.KnownDestCount = 0
  Struc.UnknownDestCount = 0
  Struc.InvalidDestCount = 0
  Struc.ResolvedQName = ""
  Struc.ResolvedQMgrName = ""
  Struc.RecsPresent = 0
  Struc.PutMsgRecFields = MQPMRF_NONE
  Struc.PutMsgRecOffset = 0
  Struc.ResponseRecOffset = 0
  Dim TempPutMsgRecPtr As MQPTR
  MQPTR_DEFAULTS TempPutMsgRecPtr
  Struc.PutMsgRecPtr = TempPutMsgRecPtr
  Dim TempResponseRecPtr As MQPTR
  MQPTR_DEFAULTS TempResponseRecPtr
  Struc.ResponseRecPtr = TempResponseRecPtr
End Sub

Sub MQRFH_DEFAULTS(Struc As MQRFH)
  Struc.StrucId = MQRFH_STRUC_ID
  Struc.Version = MQRFH_VERSION_1
  Struc.StrucLength = MQRFH_STRUC_LENGTH_FIXED
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQRFH_NONE
End Sub

Sub MQRFH2_DEFAULTS(Struc As MQRFH2)
  Struc.StrucId = MQRFH_STRUC_ID
  Struc.Version = MQRFH_VERSION_2
  Struc.StrucLength = MQRFH_STRUC_LENGTH_FIXED_2
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_INHERIT
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQRFH_NONE
  Struc.NameValueCCSID = 1208
End Sub

Sub MQRMH_DEFAULTS(Struc As MQRMH)
  Struc.StrucId = MQRMH_STRUC_ID
  Struc.Version = MQRMH_VERSION_1
  Struc.StrucLength = 0
  Struc.Encoding = MQENC_NATIVE
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQRMHF_NOT_LAST
  Struc.ObjectType = ""
  Dim TempObjectInstanceId As MQBYTE24
  MQBYTE24_DEFAULTS TempObjectInstanceId
  Struc.ObjectInstanceId = TempObjectInstanceId
  Struc.SrcEnvLength = 0
  Struc.SrcEnvOffset = 0
  Struc.SrcNameLength = 0
  Struc.SrcNameOffset = 0
  Struc.DestEnvLength = 0
  Struc.DestEnvOffset = 0
  Struc.DestNameLength = 0
  Struc.DestNameOffset = 0
  Struc.DataLogicalLength = 0
  Struc.DataLogicalOffset = 0
  Struc.DataLogicalOffset2 = 0
End Sub

Sub MQRR_DEFAULTS(Struc As MQRR)
  Struc.CompCode = MQCC_OK
  Struc.Reason = MQRC_NONE
End Sub

Sub MQTM_DEFAULTS(Struc As MQTM)
  Struc.StrucId = MQTM_STRUC_ID
  Struc.Version = MQTM_VERSION_1
  Struc.QName = ""
  Struc.ProcessName = ""
  Struc.TriggerData = ""
  Struc.ApplType = 0
  Struc.ApplId = ""
  Struc.EnvData = ""
  Struc.UserData = ""
End Sub

Sub MQTMC2_DEFAULTS(Struc As MQTMC2)
  Struc.StrucId = MQTMC_STRUC_ID
  Struc.Version = MQTMC_VERSION_2
  Struc.QName = ""
  Struc.ProcessName = ""
  Struc.TriggerData = ""
  Struc.ApplType = ""
  Struc.ApplId = ""
  Struc.EnvData = ""
  Struc.UserData = ""
  Struc.QMgrName = ""
End Sub

Sub MQWIH_DEFAULTS(Struc As MQWIH)
  Struc.StrucId = MQWIH_STRUC_ID
  Struc.Version = MQWIH_VERSION_1
  Struc.StrucLength = MQWIH_LENGTH_1
  Struc.Encoding = 0
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQWIH_NONE
  Struc.ServiceName = ""
  Struc.ServiceStep = ""
  Dim TempMsgToken As MQBYTE16
  MQBYTE16_DEFAULTS TempMsgToken
  Struc.MsgToken = TempMsgToken
  Struc.Reserved = ""
End Sub

Sub MQXQH_DEFAULTS(Struc As MQXQH)
  Struc.StrucId = MQXQH_STRUC_ID
  Struc.Version = MQXQH_VERSION_1
  Struc.RemoteQName = ""
  Struc.RemoteQMgrName = ""
  Dim TempMsgDesc As MQMD1
  MQMD1_DEFAULTS TempMsgDesc
  Struc.MsgDesc = TempMsgDesc
End Sub

'Safe Definition of MQGET'
Sub MQGET _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  GetMsgOpts As MQGMO, _
  ByVal BufferLength As Long, _
  Buffer As String, _
  DataLength As Long, _
  CompCode As Long, _
  Reason As Long)

 If BufferLength > LenB(Buffer) Then
   Reason = MQCC_FAILED
   CompCode = MQRC_BUFFER_LENGTH_ERROR
 Else
   MQGETX _
    Hconn, _
    Hobj, _
    MsgDesc, _
    GetMsgOpts, _
    BufferLength, _
    Buffer, _
    DataLength, _
    CompCode, _
    Reason
 End If
End Sub

'Safe Definition of MQINQ'
Sub MQINQ _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)

 If CharAttrLength > LenB(CharAttrs) Then
   CompCode = MQCC_FAILED
   Reason = MQRC_BUFFER_LENGTH_ERROR
 Else
   MQINQX _
    Hconn, _
    Hobj, _
    SelectorCount, _
    Selectors, _
    IntAttrCount, _
    IntAttrs, _
    CharAttrLength, _
    CharAttrs, _
    CompCode, _
    Reason
 End If
End Sub

'Safe Definition of MQPUT'
Sub MQPUT _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)

 If BufferLength > LenB(Buffer) Then
   CompCode = MQCC_FAILED
   Reason = MQRC_BUFFER_LENGTH_ERROR
 Else
   MQPUTX _
    Hconn, _
    Hobj, _
    MsgDesc, _
    PutMsgOpts, _
    BufferLength, _
    Buffer, _
    CompCode, _
    Reason
 End If
End Sub

'Safe Definition of MQPUT1'
Sub MQPUT1 _
 (ByVal Hconn As Long, _
  ObjDesc As MQOD, _
  MsgDesc As MQMD, _
  PutMsgOpts As MQPMO, _
  ByVal BufferLength As Long, _
  ByVal Buffer As String, _
  CompCode As Long, _
  Reason As Long)

 If BufferLength > LenB(Buffer) Then
   CompCode = MQCC_FAILED
   Reason = MQRC_BUFFER_LENGTH_ERROR
 Else
   MQPUT1X _
    Hconn, _
    ObjDesc, _
    MsgDesc, _
    PutMsgOpts, _
    BufferLength, _
    Buffer, _
    CompCode, _
    Reason
 End If
End Sub

'Safe Definition of MQSET'
Sub MQSET _
 (ByVal Hconn As Long, _
  ByVal Hobj As Long, _
  ByVal SelectorCount As Long, _
  Selectors As Long, _
  ByVal IntAttrCount As Long, _
  ByVal IntAttrs As Long, _
  ByVal CharAttrLength As Long, _
  ByVal CharAttrs As String, _
  CompCode As Long, _
  Reason As Long)

 If CharAttrLength > LenB(CharAttrs) Then
   CompCode = MQCC_FAILED
   Reason = MQRC_BUFFER_LENGTH_ERROR
 Else
   MQSETX _
    Hconn, _
    Hobj, _
    SelectorCount, _
    Selectors, _
    IntAttrCount, _
    IntAttrs, _
    CharAttrLength, _
    CharAttrs, _
    CompCode, _
    Reason
 End If
End Sub

Sub MQBYTE4_DEFAULTS(Struc As MQBYTE4)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
End Sub

Sub MQBYTE8_DEFAULTS(Struc As MQBYTE8)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
End Sub

Sub MQBYTE16_DEFAULTS(Struc As MQBYTE16)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
  Struc.MQByte(8) = 0
  Struc.MQByte(9) = 0
  Struc.MQByte(10) = 0
  Struc.MQByte(11) = 0
  Struc.MQByte(12) = 0
  Struc.MQByte(13) = 0
  Struc.MQByte(14) = 0
  Struc.MQByte(15) = 0
End Sub

Sub MQBYTE24_DEFAULTS(Struc As MQBYTE24)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
  Struc.MQByte(8) = 0
  Struc.MQByte(9) = 0
  Struc.MQByte(10) = 0
  Struc.MQByte(11) = 0
  Struc.MQByte(12) = 0
  Struc.MQByte(13) = 0
  Struc.MQByte(14) = 0
  Struc.MQByte(15) = 0
  Struc.MQByte(16) = 0
  Struc.MQByte(17) = 0
  Struc.MQByte(18) = 0
  Struc.MQByte(19) = 0
  Struc.MQByte(20) = 0
  Struc.MQByte(21) = 0
  Struc.MQByte(22) = 0
  Struc.MQByte(23) = 0
End Sub

Sub MQBYTE32_DEFAULTS(Struc As MQBYTE32)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
  Struc.MQByte(8) = 0
  Struc.MQByte(9) = 0
  Struc.MQByte(10) = 0
  Struc.MQByte(11) = 0
  Struc.MQByte(12) = 0
  Struc.MQByte(13) = 0
  Struc.MQByte(14) = 0
  Struc.MQByte(15) = 0
  Struc.MQByte(16) = 0
  Struc.MQByte(17) = 0
  Struc.MQByte(18) = 0
  Struc.MQByte(19) = 0
  Struc.MQByte(20) = 0
  Struc.MQByte(21) = 0
  Struc.MQByte(22) = 0
  Struc.MQByte(23) = 0
  Struc.MQByte(24) = 0
  Struc.MQByte(25) = 0
  Struc.MQByte(26) = 0
  Struc.MQByte(27) = 0
  Struc.MQByte(28) = 0
  Struc.MQByte(29) = 0
  Struc.MQByte(30) = 0
  Struc.MQByte(31) = 0
End Sub

Sub MQBYTE40_DEFAULTS(Struc As MQBYTE40)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
  Struc.MQByte(8) = 0
  Struc.MQByte(9) = 0
  Struc.MQByte(10) = 0
  Struc.MQByte(11) = 0
  Struc.MQByte(12) = 0
  Struc.MQByte(13) = 0
  Struc.MQByte(14) = 0
  Struc.MQByte(15) = 0
  Struc.MQByte(16) = 0
  Struc.MQByte(17) = 0
  Struc.MQByte(18) = 0
  Struc.MQByte(19) = 0
  Struc.MQByte(20) = 0
  Struc.MQByte(21) = 0
  Struc.MQByte(22) = 0
  Struc.MQByte(23) = 0
  Struc.MQByte(24) = 0
  Struc.MQByte(25) = 0
  Struc.MQByte(26) = 0
  Struc.MQByte(27) = 0
  Struc.MQByte(28) = 0
  Struc.MQByte(29) = 0
  Struc.MQByte(30) = 0
  Struc.MQByte(31) = 0
  Struc.MQByte(32) = 0
  Struc.MQByte(33) = 0
  Struc.MQByte(34) = 0
  Struc.MQByte(35) = 0
  Struc.MQByte(36) = 0
  Struc.MQByte(37) = 0
  Struc.MQByte(38) = 0
  Struc.MQByte(39) = 0
End Sub

Sub MQBYTE48_DEFAULTS(Struc As MQBYTE48)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
  Struc.MQByte(8) = 0
  Struc.MQByte(9) = 0
  Struc.MQByte(10) = 0
  Struc.MQByte(11) = 0
  Struc.MQByte(12) = 0
  Struc.MQByte(13) = 0
  Struc.MQByte(14) = 0
  Struc.MQByte(15) = 0
  Struc.MQByte(16) = 0
  Struc.MQByte(17) = 0
  Struc.MQByte(18) = 0
  Struc.MQByte(19) = 0
  Struc.MQByte(20) = 0
  Struc.MQByte(21) = 0
  Struc.MQByte(22) = 0
  Struc.MQByte(23) = 0
  Struc.MQByte(24) = 0
  Struc.MQByte(25) = 0
  Struc.MQByte(26) = 0
  Struc.MQByte(27) = 0
  Struc.MQByte(28) = 0
  Struc.MQByte(29) = 0
  Struc.MQByte(30) = 0
  Struc.MQByte(31) = 0
  Struc.MQByte(32) = 0
  Struc.MQByte(33) = 0
  Struc.MQByte(34) = 0
  Struc.MQByte(35) = 0
  Struc.MQByte(36) = 0
  Struc.MQByte(37) = 0
  Struc.MQByte(38) = 0
  Struc.MQByte(39) = 0
  Struc.MQByte(40) = 0
  Struc.MQByte(41) = 0
  Struc.MQByte(42) = 0
  Struc.MQByte(43) = 0
  Struc.MQByte(44) = 0
  Struc.MQByte(45) = 0
  Struc.MQByte(46) = 0
  Struc.MQByte(47) = 0
End Sub

Sub MQBYTE128_DEFAULTS(Struc As MQBYTE128)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
  Struc.MQByte(4) = 0
  Struc.MQByte(5) = 0
  Struc.MQByte(6) = 0
  Struc.MQByte(7) = 0
  Struc.MQByte(8) = 0
  Struc.MQByte(9) = 0
  Struc.MQByte(10) = 0
  Struc.MQByte(11) = 0
  Struc.MQByte(12) = 0
  Struc.MQByte(13) = 0
  Struc.MQByte(14) = 0
  Struc.MQByte(15) = 0
  Struc.MQByte(16) = 0
  Struc.MQByte(17) = 0
  Struc.MQByte(18) = 0
  Struc.MQByte(19) = 0
  Struc.MQByte(20) = 0
  Struc.MQByte(21) = 0
  Struc.MQByte(22) = 0
  Struc.MQByte(23) = 0
  Struc.MQByte(24) = 0
  Struc.MQByte(25) = 0
  Struc.MQByte(26) = 0
  Struc.MQByte(27) = 0
  Struc.MQByte(28) = 0
  Struc.MQByte(29) = 0
  Struc.MQByte(30) = 0
  Struc.MQByte(31) = 0
  Struc.MQByte(32) = 0
  Struc.MQByte(33) = 0
  Struc.MQByte(34) = 0
  Struc.MQByte(35) = 0
  Struc.MQByte(36) = 0
  Struc.MQByte(37) = 0
  Struc.MQByte(38) = 0
  Struc.MQByte(39) = 0
  Struc.MQByte(40) = 0
  Struc.MQByte(41) = 0
  Struc.MQByte(42) = 0
  Struc.MQByte(43) = 0
  Struc.MQByte(44) = 0
  Struc.MQByte(45) = 0
  Struc.MQByte(46) = 0
  Struc.MQByte(47) = 0
  Struc.MQByte(48) = 0
  Struc.MQByte(49) = 0
  Struc.MQByte(50) = 0
  Struc.MQByte(51) = 0
  Struc.MQByte(52) = 0
  Struc.MQByte(53) = 0
  Struc.MQByte(54) = 0
  Struc.MQByte(55) = 0
  Struc.MQByte(56) = 0
  Struc.MQByte(57) = 0
  Struc.MQByte(58) = 0
  Struc.MQByte(59) = 0
  Struc.MQByte(60) = 0
  Struc.MQByte(61) = 0
  Struc.MQByte(62) = 0
  Struc.MQByte(63) = 0
  Struc.MQByte(64) = 0
  Struc.MQByte(65) = 0
  Struc.MQByte(66) = 0
  Struc.MQByte(67) = 0
  Struc.MQByte(68) = 0
  Struc.MQByte(69) = 0
  Struc.MQByte(70) = 0
  Struc.MQByte(71) = 0
  Struc.MQByte(72) = 0
  Struc.MQByte(73) = 0
  Struc.MQByte(74) = 0
  Struc.MQByte(75) = 0
  Struc.MQByte(76) = 0
  Struc.MQByte(77) = 0
  Struc.MQByte(78) = 0
  Struc.MQByte(79) = 0
  Struc.MQByte(80) = 0
  Struc.MQByte(81) = 0
  Struc.MQByte(82) = 0
  Struc.MQByte(83) = 0
  Struc.MQByte(84) = 0
  Struc.MQByte(85) = 0
  Struc.MQByte(86) = 0
  Struc.MQByte(87) = 0
  Struc.MQByte(88) = 0
  Struc.MQByte(89) = 0
  Struc.MQByte(90) = 0
  Struc.MQByte(91) = 0
  Struc.MQByte(92) = 0
  Struc.MQByte(93) = 0
  Struc.MQByte(94) = 0
  Struc.MQByte(95) = 0
  Struc.MQByte(96) = 0
  Struc.MQByte(97) = 0
  Struc.MQByte(98) = 0
  Struc.MQByte(99) = 0
  Struc.MQByte(100) = 0
  Struc.MQByte(101) = 0
  Struc.MQByte(102) = 0
  Struc.MQByte(103) = 0
  Struc.MQByte(104) = 0
  Struc.MQByte(105) = 0
  Struc.MQByte(106) = 0
  Struc.MQByte(107) = 0
  Struc.MQByte(108) = 0
  Struc.MQByte(109) = 0
  Struc.MQByte(110) = 0
  Struc.MQByte(111) = 0
  Struc.MQByte(112) = 0
  Struc.MQByte(113) = 0
  Struc.MQByte(114) = 0
  Struc.MQByte(115) = 0
  Struc.MQByte(116) = 0
  Struc.MQByte(117) = 0
  Struc.MQByte(118) = 0
  Struc.MQByte(119) = 0
  Struc.MQByte(120) = 0
  Struc.MQByte(121) = 0
  Struc.MQByte(122) = 0
  Struc.MQByte(123) = 0
  Struc.MQByte(124) = 0
  Struc.MQByte(125) = 0
  Struc.MQByte(126) = 0
  Struc.MQByte(127) = 0
End Sub

Sub MQPTR_DEFAULTS(Struc As MQPTR)
  Struc.MQByte(0) = 0
  Struc.MQByte(1) = 0
  Struc.MQByte(2) = 0
  Struc.MQByte(3) = 0
End Sub

Sub MQ_SETDEFAULTS()

  'Set byte-string constants'
  MQBYTE8_DEFAULTS MQCFAC_NONE
  MQBYTE128_DEFAULTS MQCT_NONE
  MQBYTE24_DEFAULTS MQCONNID_NONE
  MQBYTE16_DEFAULTS MQMTOK_NONE
  MQBYTE16_DEFAULTS MQITII_NONE
  MQBYTE24_DEFAULTS MQMI_NONE
  MQBYTE24_DEFAULTS MQCI_NONE
  MQBYTE24_DEFAULTS MQCI_NEW_SESSION
  MQBYTE32_DEFAULTS MQACT_NONE
  MQBYTE24_DEFAULTS MQGI_NONE
  MQBYTE40_DEFAULTS MQSID_NONE
  MQBYTE24_DEFAULTS MQOII_NONE
  MQCI_NEW_SESSION.MQByte(0) = &H41
  MQCI_NEW_SESSION.MQByte(1) = &H4D
  MQCI_NEW_SESSION.MQByte(2) = &H51
  MQCI_NEW_SESSION.MQByte(3) = &H21
  MQCI_NEW_SESSION.MQByte(4) = &H4E
  MQCI_NEW_SESSION.MQByte(5) = &H45
  MQCI_NEW_SESSION.MQByte(6) = &H57
  MQCI_NEW_SESSION.MQByte(7) = &H5F
  MQCI_NEW_SESSION.MQByte(8) = &H53
  MQCI_NEW_SESSION.MQByte(9) = &H45
  MQCI_NEW_SESSION.MQByte(10) = &H53
  MQCI_NEW_SESSION.MQByte(11) = &H53
  MQCI_NEW_SESSION.MQByte(12) = &H49
  MQCI_NEW_SESSION.MQByte(13) = &H4F
  MQCI_NEW_SESSION.MQByte(14) = &H4E
  MQCI_NEW_SESSION.MQByte(15) = &H5F
  MQCI_NEW_SESSION.MQByte(16) = &H43
  MQCI_NEW_SESSION.MQByte(17) = &H4F
  MQCI_NEW_SESSION.MQByte(18) = &H52
  MQCI_NEW_SESSION.MQByte(19) = &H52
  MQCI_NEW_SESSION.MQByte(20) = &H45
  MQCI_NEW_SESSION.MQByte(21) = &H4C
  MQCI_NEW_SESSION.MQByte(22) = &H49
  MQCI_NEW_SESSION.MQByte(23) = &H44

  'Set default structures'
  MQBYTE4_DEFAULTS MQBYTE4_DEFAULT
  MQBYTE8_DEFAULTS MQBYTE8_DEFAULT
  MQBYTE16_DEFAULTS MQBYTE16_DEFAULT
  MQBYTE24_DEFAULTS MQBYTE24_DEFAULT
  MQBYTE32_DEFAULTS MQBYTE32_DEFAULT
  MQBYTE40_DEFAULTS MQBYTE40_DEFAULT
  MQBYTE48_DEFAULTS MQBYTE48_DEFAULT
  MQBYTE128_DEFAULTS MQBYTE128_DEFAULT
  MQPTR_DEFAULTS MQPTR_DEFAULT
  MQAIR_DEFAULTS MQAIR_DEFAULT
  MQBO_DEFAULTS MQBO_DEFAULT
  MQCIH_DEFAULTS MQCIH_DEFAULT
  MQSCO_DEFAULTS MQSCO_DEFAULT
  MQCSP_DEFAULTS MQCSP_DEFAULT
  MQCNO_DEFAULTS MQCNO_DEFAULT
  MQDH_DEFAULTS MQDH_DEFAULT
  MQDLH_DEFAULTS MQDLH_DEFAULT
  MQGMO_DEFAULTS MQGMO_DEFAULT
  MQIIH_DEFAULTS MQIIH_DEFAULT
  MQMD_DEFAULTS MQMD_DEFAULT
  MQMDE_DEFAULTS MQMDE_DEFAULT
  MQMD1_DEFAULTS MQMD1_DEFAULT
  MQMD2_DEFAULTS MQMD2_DEFAULT
  MQOD_DEFAULTS MQOD_DEFAULT
  MQOR_DEFAULTS MQOR_DEFAULT
  MQPMO_DEFAULTS MQPMO_DEFAULT
  MQRFH_DEFAULTS MQRFH_DEFAULT
  MQRFH2_DEFAULTS MQRFH2_DEFAULT
  MQRMH_DEFAULTS MQRMH_DEFAULT
  MQRR_DEFAULTS MQRR_DEFAULT
  MQTM_DEFAULTS MQTM_DEFAULT
  MQTMC2_DEFAULTS MQTMC2_DEFAULT
  MQWIH_DEFAULTS MQWIH_DEFAULT
  MQXQH_DEFAULTS MQXQH_DEFAULT

End Sub

'****************************************************************'
'*  End of CMQB                                                 *'
'****************************************************************'

