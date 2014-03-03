Attribute VB_Name = "CMQCFB"
'**********************************************************************'
'*                                                                    *'
'*                  WebSphere MQ for Windows                          *'
'*                                                                    *'
'*  FILE NAME:      CMQCFB                                            *'
'*                                                                    *'
'*  DESCRIPTION:    Declarations for PCF and Events                   *'
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
'*  FUNCTION:       This file declares the structures and named       *'
'*                  constants for PCF and event messages.             *'
'*                                                                    *'
'*  PROCESSOR:      BASIC                                             *'
'*                                                                    *'
'**********************************************************************'

'$$#con$ on'
'****************************************************************'
'*  Values Related to MQCFH Structure                           *'
'****************************************************************'
'Structure Length'
Global Const MQCFH_STRUC_LENGTH = 36

'Structure Version Number'
Global Const MQCFH_VERSION_1 = 1
Global Const MQCFH_VERSION_2 = 2
Global Const MQCFH_VERSION_3 = 3
Global Const MQCFH_CURRENT_VERSION = 3

'Command Codes'
Global Const MQCMD_CHANGE_Q_MGR = 1
Global Const MQCMD_INQUIRE_Q_MGR = 2
Global Const MQCMD_CHANGE_PROCESS = 3
Global Const MQCMD_COPY_PROCESS = 4
Global Const MQCMD_CREATE_PROCESS = 5
Global Const MQCMD_DELETE_PROCESS = 6
Global Const MQCMD_INQUIRE_PROCESS = 7
Global Const MQCMD_CHANGE_Q = 8
Global Const MQCMD_CLEAR_Q = 9
Global Const MQCMD_COPY_Q = 10
Global Const MQCMD_CREATE_Q = 11
Global Const MQCMD_DELETE_Q = 12
Global Const MQCMD_INQUIRE_Q = 13
Global Const MQCMD_REFRESH_Q_MGR = 16
Global Const MQCMD_RESET_Q_STATS = 17
Global Const MQCMD_INQUIRE_Q_NAMES = 18
Global Const MQCMD_INQUIRE_PROCESS_NAMES = 19
Global Const MQCMD_INQUIRE_CHANNEL_NAMES = 20
Global Const MQCMD_CHANGE_CHANNEL = 21
Global Const MQCMD_COPY_CHANNEL = 22
Global Const MQCMD_CREATE_CHANNEL = 23
Global Const MQCMD_DELETE_CHANNEL = 24
Global Const MQCMD_INQUIRE_CHANNEL = 25
Global Const MQCMD_PING_CHANNEL = 26
Global Const MQCMD_RESET_CHANNEL = 27
Global Const MQCMD_START_CHANNEL = 28
Global Const MQCMD_STOP_CHANNEL = 29
Global Const MQCMD_START_CHANNEL_INIT = 30
Global Const MQCMD_START_CHANNEL_LISTENER = 31
Global Const MQCMD_CHANGE_NAMELIST = 32
Global Const MQCMD_COPY_NAMELIST = 33
Global Const MQCMD_CREATE_NAMELIST = 34
Global Const MQCMD_DELETE_NAMELIST = 35
Global Const MQCMD_INQUIRE_NAMELIST = 36
Global Const MQCMD_INQUIRE_NAMELIST_NAMES = 37
Global Const MQCMD_ESCAPE = 38
Global Const MQCMD_RESOLVE_CHANNEL = 39
Global Const MQCMD_PING_Q_MGR = 40
Global Const MQCMD_INQUIRE_Q_STATUS = 41
Global Const MQCMD_INQUIRE_CHANNEL_STATUS = 42
Global Const MQCMD_CONFIG_EVENT = 43
Global Const MQCMD_Q_MGR_EVENT = 44
Global Const MQCMD_PERFM_EVENT = 45
Global Const MQCMD_CHANNEL_EVENT = 46
Global Const MQCMD_DELETE_PUBLICATION = 60
Global Const MQCMD_DEREGISTER_PUBLISHER = 61
Global Const MQCMD_DEREGISTER_SUBSCRIBER = 62
Global Const MQCMD_PUBLISH = 63
Global Const MQCMD_REGISTER_PUBLISHER = 64
Global Const MQCMD_REGISTER_SUBSCRIBER = 65
Global Const MQCMD_REQUEST_UPDATE = 66
Global Const MQCMD_BROKER_INTERNAL = 67
Global Const MQCMD_ACTIVITY_MSG = 69
Global Const MQCMD_INQUIRE_CLUSTER_Q_MGR = 70
Global Const MQCMD_RESUME_Q_MGR_CLUSTER = 71
Global Const MQCMD_SUSPEND_Q_MGR_CLUSTER = 72
Global Const MQCMD_REFRESH_CLUSTER = 73
Global Const MQCMD_RESET_CLUSTER = 74
Global Const MQCMD_TRACE_ROUTE = 75
Global Const MQCMD_REFRESH_SECURITY = 78
Global Const MQCMD_CHANGE_AUTH_INFO = 79
Global Const MQCMD_COPY_AUTH_INFO = 80
Global Const MQCMD_CREATE_AUTH_INFO = 81
Global Const MQCMD_DELETE_AUTH_INFO = 82
Global Const MQCMD_INQUIRE_AUTH_INFO = 83
Global Const MQCMD_INQUIRE_AUTH_INFO_NAMES = 84
Global Const MQCMD_INQUIRE_CONNECTION = 85
Global Const MQCMD_STOP_CONNECTION = 86
Global Const MQCMD_INQUIRE_AUTH_RECS = 87
Global Const MQCMD_INQUIRE_ENTITY_AUTH = 88
Global Const MQCMD_DELETE_AUTH_REC = 89
Global Const MQCMD_SET_AUTH_REC = 90
Global Const MQCMD_LOGGER_EVENT = 91
Global Const MQCMD_RESET_Q_MGR = 92
Global Const MQCMD_CHANGE_LISTENER = 93
Global Const MQCMD_COPY_LISTENER = 94
Global Const MQCMD_CREATE_LISTENER = 95
Global Const MQCMD_DELETE_LISTENER = 96
Global Const MQCMD_INQUIRE_LISTENER = 97
Global Const MQCMD_INQUIRE_LISTENER_STATUS = 98
Global Const MQCMD_COMMAND_EVENT = 99
Global Const MQCMD_CHANGE_SECURITY = 100
Global Const MQCMD_CHANGE_CF_STRUC = 101
Global Const MQCMD_CHANGE_STG_CLASS = 102
Global Const MQCMD_CHANGE_TRACE = 103
Global Const MQCMD_ARCHIVE_LOG = 104
Global Const MQCMD_BACKUP_CF_STRUC = 105
Global Const MQCMD_CREATE_BUFFER_POOL = 106
Global Const MQCMD_CREATE_PAGE_SET = 107
Global Const MQCMD_CREATE_CF_STRUC = 108
Global Const MQCMD_CREATE_STG_CLASS = 109
Global Const MQCMD_COPY_CF_STRUC = 110
Global Const MQCMD_COPY_STG_CLASS = 111
Global Const MQCMD_DELETE_CF_STRUC = 112
Global Const MQCMD_DELETE_STG_CLASS = 113
Global Const MQCMD_INQUIRE_ARCHIVE = 114
Global Const MQCMD_INQUIRE_CF_STRUC = 115
Global Const MQCMD_INQUIRE_CF_STRUC_STATUS = 116
Global Const MQCMD_INQUIRE_CMD_SERVER = 117
Global Const MQCMD_INQUIRE_CHANNEL_INIT = 118
Global Const MQCMD_INQUIRE_QSG = 119
Global Const MQCMD_INQUIRE_LOG = 120
Global Const MQCMD_INQUIRE_SECURITY = 121
Global Const MQCMD_INQUIRE_STG_CLASS = 122
Global Const MQCMD_INQUIRE_SYSTEM = 123
Global Const MQCMD_INQUIRE_THREAD = 124
Global Const MQCMD_INQUIRE_TRACE = 125
Global Const MQCMD_INQUIRE_USAGE = 126
Global Const MQCMD_MOVE_Q = 127
Global Const MQCMD_RECOVER_BSDS = 128
Global Const MQCMD_RECOVER_CF_STRUC = 129
Global Const MQCMD_RESET_TPIPE = 130
Global Const MQCMD_RESOLVE_INDOUBT = 131
Global Const MQCMD_RESUME_Q_MGR = 132
Global Const MQCMD_REVERIFY_SECURITY = 133
Global Const MQCMD_SET_ARCHIVE = 134
Global Const MQCMD_SET_LOG = 136
Global Const MQCMD_SET_SYSTEM = 137
Global Const MQCMD_START_CMD_SERVER = 138
Global Const MQCMD_START_Q_MGR = 139
Global Const MQCMD_START_TRACE = 140
Global Const MQCMD_STOP_CHANNEL_INIT = 141
Global Const MQCMD_STOP_CHANNEL_LISTENER = 142
Global Const MQCMD_STOP_CMD_SERVER = 143
Global Const MQCMD_STOP_Q_MGR = 144
Global Const MQCMD_STOP_TRACE = 145
Global Const MQCMD_SUSPEND_Q_MGR = 146
Global Const MQCMD_INQUIRE_CF_STRUC_NAMES = 147
Global Const MQCMD_INQUIRE_STG_CLASS_NAMES = 148
Global Const MQCMD_CHANGE_SERVICE = 149
Global Const MQCMD_COPY_SERVICE = 150
Global Const MQCMD_CREATE_SERVICE = 151
Global Const MQCMD_DELETE_SERVICE = 152
Global Const MQCMD_INQUIRE_SERVICE = 153
Global Const MQCMD_INQUIRE_SERVICE_STATUS = 154
Global Const MQCMD_START_SERVICE = 155
Global Const MQCMD_STOP_SERVICE = 156
Global Const MQCMD_DELETE_BUFFER_POOL = 157
Global Const MQCMD_DELETE_PAGE_SET = 158
Global Const MQCMD_CHANGE_BUFFER_POOL = 159
Global Const MQCMD_CHANGE_PAGE_SET = 160
Global Const MQCMD_INQUIRE_Q_MGR_STATUS = 161
Global Const MQCMD_CREATE_LOG = 162
Global Const MQCMD_STATISTICS_MQI = 164
Global Const MQCMD_STATISTICS_Q = 165
Global Const MQCMD_STATISTICS_CHANNEL = 166
Global Const MQCMD_ACCOUNTING_MQI = 167
Global Const MQCMD_ACCOUNTING_Q = 168
Global Const MQCMD_INQUIRE_AUTH_SERVICE = 169
Global Const MQCMD_NONE = 0

'Control Options'
Global Const MQCFC_LAST = 1
Global Const MQCFC_NOT_LAST = 0

'Reason Codes'
Global Const MQRCCF_CFH_TYPE_ERROR = 3001
Global Const MQRCCF_CFH_LENGTH_ERROR = 3002
Global Const MQRCCF_CFH_VERSION_ERROR = 3003
Global Const MQRCCF_CFH_MSG_SEQ_NUMBER_ERR = 3004
Global Const MQRCCF_CFH_CONTROL_ERROR = 3005
Global Const MQRCCF_CFH_PARM_COUNT_ERROR = 3006
Global Const MQRCCF_CFH_COMMAND_ERROR = 3007
Global Const MQRCCF_COMMAND_FAILED = 3008
Global Const MQRCCF_CFIN_LENGTH_ERROR = 3009
Global Const MQRCCF_CFST_LENGTH_ERROR = 3010
Global Const MQRCCF_CFST_STRING_LENGTH_ERR = 3011
Global Const MQRCCF_FORCE_VALUE_ERROR = 3012
Global Const MQRCCF_STRUCTURE_TYPE_ERROR = 3013
Global Const MQRCCF_CFIN_PARM_ID_ERROR = 3014
Global Const MQRCCF_CFST_PARM_ID_ERROR = 3015
Global Const MQRCCF_MSG_LENGTH_ERROR = 3016
Global Const MQRCCF_CFIN_DUPLICATE_PARM = 3017
Global Const MQRCCF_CFST_DUPLICATE_PARM = 3018
Global Const MQRCCF_PARM_COUNT_TOO_SMALL = 3019
Global Const MQRCCF_PARM_COUNT_TOO_BIG = 3020
Global Const MQRCCF_Q_ALREADY_IN_CELL = 3021
Global Const MQRCCF_Q_TYPE_ERROR = 3022
Global Const MQRCCF_MD_FORMAT_ERROR = 3023
Global Const MQRCCF_CFSL_LENGTH_ERROR = 3024
Global Const MQRCCF_REPLACE_VALUE_ERROR = 3025
Global Const MQRCCF_CFIL_DUPLICATE_VALUE = 3026
Global Const MQRCCF_CFIL_COUNT_ERROR = 3027
Global Const MQRCCF_CFIL_LENGTH_ERROR = 3028
Global Const MQRCCF_QUIESCE_VALUE_ERROR = 3029
Global Const MQRCCF_MODE_VALUE_ERROR = 3029
Global Const MQRCCF_MSG_SEQ_NUMBER_ERROR = 3030
Global Const MQRCCF_PING_DATA_COUNT_ERROR = 3031
Global Const MQRCCF_PING_DATA_COMPARE_ERROR = 3032
Global Const MQRCCF_CFSL_PARM_ID_ERROR = 3033
Global Const MQRCCF_CHANNEL_TYPE_ERROR = 3034
Global Const MQRCCF_PARM_SEQUENCE_ERROR = 3035
Global Const MQRCCF_XMIT_PROTOCOL_TYPE_ERR = 3036
Global Const MQRCCF_BATCH_SIZE_ERROR = 3037
Global Const MQRCCF_DISC_INT_ERROR = 3038
Global Const MQRCCF_SHORT_RETRY_ERROR = 3039
Global Const MQRCCF_SHORT_TIMER_ERROR = 3040
Global Const MQRCCF_LONG_RETRY_ERROR = 3041
Global Const MQRCCF_LONG_TIMER_ERROR = 3042
Global Const MQRCCF_SEQ_NUMBER_WRAP_ERROR = 3043
Global Const MQRCCF_MAX_MSG_LENGTH_ERROR = 3044
Global Const MQRCCF_PUT_AUTH_ERROR = 3045
Global Const MQRCCF_PURGE_VALUE_ERROR = 3046
Global Const MQRCCF_CFIL_PARM_ID_ERROR = 3047
Global Const MQRCCF_MSG_TRUNCATED = 3048
Global Const MQRCCF_CCSID_ERROR = 3049
Global Const MQRCCF_ENCODING_ERROR = 3050
Global Const MQRCCF_QUEUES_VALUE_ERROR = 3051
Global Const MQRCCF_DATA_CONV_VALUE_ERROR = 3052
Global Const MQRCCF_INDOUBT_VALUE_ERROR = 3053
Global Const MQRCCF_ESCAPE_TYPE_ERROR = 3054
Global Const MQRCCF_REPOS_VALUE_ERROR = 3055
Global Const MQRCCF_CHANNEL_TABLE_ERROR = 3062
Global Const MQRCCF_MCA_TYPE_ERROR = 3063
Global Const MQRCCF_CHL_INST_TYPE_ERROR = 3064
Global Const MQRCCF_CHL_STATUS_NOT_FOUND = 3065
Global Const MQRCCF_CFSL_DUPLICATE_PARM = 3066
Global Const MQRCCF_CFSL_TOTAL_LENGTH_ERROR = 3067
Global Const MQRCCF_CFSL_COUNT_ERROR = 3068
Global Const MQRCCF_CFSL_STRING_LENGTH_ERR = 3069
Global Const MQRCCF_BROKER_DELETED = 3070
Global Const MQRCCF_STREAM_ERROR = 3071
Global Const MQRCCF_TOPIC_ERROR = 3072
Global Const MQRCCF_NOT_REGISTERED = 3073
Global Const MQRCCF_Q_MGR_NAME_ERROR = 3074
Global Const MQRCCF_INCORRECT_STREAM = 3075
Global Const MQRCCF_Q_NAME_ERROR = 3076
Global Const MQRCCF_NO_RETAINED_MSG = 3077
Global Const MQRCCF_DUPLICATE_IDENTITY = 3078
Global Const MQRCCF_INCORRECT_Q = 3079
Global Const MQRCCF_CORREL_ID_ERROR = 3080
Global Const MQRCCF_NOT_AUTHORIZED = 3081
Global Const MQRCCF_UNKNOWN_STREAM = 3082
Global Const MQRCCF_REG_OPTIONS_ERROR = 3083
Global Const MQRCCF_PUB_OPTIONS_ERROR = 3084
Global Const MQRCCF_UNKNOWN_BROKER = 3085
Global Const MQRCCF_Q_MGR_CCSID_ERROR = 3086
Global Const MQRCCF_DEL_OPTIONS_ERROR = 3087
Global Const MQRCCF_CLUSTER_NAME_CONFLICT = 3088
Global Const MQRCCF_REPOS_NAME_CONFLICT = 3089
Global Const MQRCCF_CLUSTER_Q_USAGE_ERROR = 3090
Global Const MQRCCF_ACTION_VALUE_ERROR = 3091
Global Const MQRCCF_COMMS_LIBRARY_ERROR = 3092
Global Const MQRCCF_NETBIOS_NAME_ERROR = 3093
Global Const MQRCCF_BROKER_COMMAND_FAILED = 3094
Global Const MQRCCF_CFST_CONFLICTING_PARM = 3095
Global Const MQRCCF_PATH_NOT_VALID = 3096
Global Const MQRCCF_PARM_SYNTAX_ERROR = 3097
Global Const MQRCCF_PWD_LENGTH_ERROR = 3098
Global Const MQRCCF_FILTER_ERROR = 3150
Global Const MQRCCF_WRONG_USER = 3151
Global Const MQRCCF_DUPLICATE_SUBSCRIPTION = 3152
Global Const MQRCCF_SUB_NAME_ERROR = 3153
Global Const MQRCCF_SUB_IDENTITY_ERROR = 3154
Global Const MQRCCF_SUBSCRIPTION_IN_USE = 3155
Global Const MQRCCF_SUBSCRIPTION_LOCKED = 3156
Global Const MQRCCF_ALREADY_JOINED = 3157
Global Const MQRCCF_OBJECT_IN_USE = 3160
Global Const MQRCCF_UNKNOWN_FILE_NAME = 3161
Global Const MQRCCF_FILE_NOT_AVAILABLE = 3162
Global Const MQRCCF_DISC_RETRY_ERROR = 3163
Global Const MQRCCF_ALLOC_RETRY_ERROR = 3164
Global Const MQRCCF_ALLOC_SLOW_TIMER_ERROR = 3165
Global Const MQRCCF_ALLOC_FAST_TIMER_ERROR = 3166
Global Const MQRCCF_PORT_NUMBER_ERROR = 3167
Global Const MQRCCF_CHL_SYSTEM_NOT_ACTIVE = 3168
Global Const MQRCCF_ENTITY_NAME_MISSING = 3169
Global Const MQRCCF_PROFILE_NAME_ERROR = 3170
Global Const MQRCCF_AUTH_VALUE_ERROR = 3171
Global Const MQRCCF_AUTH_VALUE_MISSING = 3172
Global Const MQRCCF_OBJECT_TYPE_MISSING = 3173
Global Const MQRCCF_CONNECTION_ID_ERROR = 3174
Global Const MQRCCF_LOG_TYPE_ERROR = 3175
Global Const MQRCCF_PROGRAM_NOT_AVAILABLE = 3176
Global Const MQRCCF_PROGRAM_AUTH_FAILED = 3177
Global Const MQRCCF_NONE_FOUND = 3200
Global Const MQRCCF_SECURITY_SWITCH_OFF = 3201
Global Const MQRCCF_SECURITY_REFRESH_FAILED = 3202
Global Const MQRCCF_PARM_CONFLICT = 3203
Global Const MQRCCF_COMMAND_INHIBITED = 3204
Global Const MQRCCF_OBJECT_BEING_DELETED = 3205
Global Const MQRCCF_STORAGE_CLASS_IN_USE = 3207
Global Const MQRCCF_OBJECT_NAME_RESTRICTED = 3208
Global Const MQRCCF_OBJECT_LIMIT_EXCEEDED = 3209
Global Const MQRCCF_OBJECT_OPEN_FORCE = 3210
Global Const MQRCCF_DISPOSITION_CONFLICT = 3211
Global Const MQRCCF_Q_MGR_NOT_IN_QSG = 3212
Global Const MQRCCF_ATTR_VALUE_FIXED = 3213
Global Const MQRCCF_NAMELIST_ERROR = 3215
Global Const MQRCCF_NO_CHANNEL_INITIATOR = 3217
Global Const MQRCCF_CHANNEL_INITIATOR_ERROR = 3218
Global Const MQRCCF_COMMAND_LEVEL_CONFLICT = 3222
Global Const MQRCCF_Q_ATTR_CONFLICT = 3223
Global Const MQRCCF_EVENTS_DISABLED = 3224
Global Const MQRCCF_COMMAND_SCOPE_ERROR = 3225
Global Const MQRCCF_COMMAND_REPLY_ERROR = 3226
Global Const MQRCCF_FUNCTION_RESTRICTED = 3227
Global Const MQRCCF_PARM_MISSING = 3228
Global Const MQRCCF_PARM_VALUE_ERROR = 3229
Global Const MQRCCF_COMMAND_LENGTH_ERROR = 3230
Global Const MQRCCF_COMMAND_ORIGIN_ERROR = 3231
Global Const MQRCCF_LISTENER_CONFLICT = 3232
Global Const MQRCCF_LISTENER_STARTED = 3233
Global Const MQRCCF_LISTENER_STOPPED = 3234
Global Const MQRCCF_CHANNEL_ERROR = 3235
Global Const MQRCCF_CF_STRUC_ERROR = 3236
Global Const MQRCCF_UNKNOWN_USER_ID = 3237
Global Const MQRCCF_UNEXPECTED_ERROR = 3238
Global Const MQRCCF_NO_XCF_PARTNER = 3239
Global Const MQRCCF_CFGR_PARM_ID_ERROR = 3240
Global Const MQRCCF_CFIF_LENGTH_ERROR = 3241
Global Const MQRCCF_CFIF_OPERATOR_ERROR = 3242
Global Const MQRCCF_CFIF_PARM_ID_ERROR = 3243
Global Const MQRCCF_CFSF_FILTER_VAL_LEN_ERR = 3244
Global Const MQRCCF_CFSF_LENGTH_ERROR = 3245
Global Const MQRCCF_CFSF_OPERATOR_ERROR = 3246
Global Const MQRCCF_CFSF_PARM_ID_ERROR = 3247
Global Const MQRCCF_TOO_MANY_FILTERS = 3248
Global Const MQRCCF_LISTENER_RUNNING = 3249
Global Const MQRCCF_LSTR_STATUS_NOT_FOUND = 3250
Global Const MQRCCF_SERVICE_RUNNING = 3251
Global Const MQRCCF_SERV_STATUS_NOT_FOUND = 3252
Global Const MQRCCF_SERVICE_STOPPED = 3253
Global Const MQRCCF_CFBS_DUPLICATE_PARM = 3254
Global Const MQRCCF_CFBS_LENGTH_ERROR = 3255
Global Const MQRCCF_CFBS_PARM_ID_ERROR = 3256
Global Const MQRCCF_CFBS_STRING_LENGTH_ERR = 3257
Global Const MQRCCF_CFGR_LENGTH_ERROR = 3258
Global Const MQRCCF_CFGR_PARM_COUNT_ERROR = 3259
Global Const MQRCCF_CONN_NOT_STOPPED = 3260
Global Const MQRCCF_SERVICE_REQUEST_PENDING = 3261
Global Const MQRCCF_NO_START_CMD = 3262
Global Const MQRCCF_NO_STOP_CMD = 3263
Global Const MQRCCF_CFBF_LENGTH_ERROR = 3264
Global Const MQRCCF_CFBF_PARM_ID_ERROR = 3265
Global Const MQRCCF_CFBF_OPERATOR_ERROR = 3266
Global Const MQRCCF_CFBF_FILTER_VAL_LEN_ERR = 3267
Global Const MQRCCF_LISTENER_STILL_ACTIVE = 3268
Global Const MQRCCF_OBJECT_ALREADY_EXISTS = 4001
Global Const MQRCCF_OBJECT_WRONG_TYPE = 4002
Global Const MQRCCF_LIKE_OBJECT_WRONG_TYPE = 4003
Global Const MQRCCF_OBJECT_OPEN = 4004
Global Const MQRCCF_ATTR_VALUE_ERROR = 4005
Global Const MQRCCF_UNKNOWN_Q_MGR = 4006
Global Const MQRCCF_Q_WRONG_TYPE = 4007
Global Const MQRCCF_OBJECT_NAME_ERROR = 4008
Global Const MQRCCF_ALLOCATE_FAILED = 4009
Global Const MQRCCF_HOST_NOT_AVAILABLE = 4010
Global Const MQRCCF_CONFIGURATION_ERROR = 4011
Global Const MQRCCF_CONNECTION_REFUSED = 4012
Global Const MQRCCF_ENTRY_ERROR = 4013
Global Const MQRCCF_SEND_FAILED = 4014
Global Const MQRCCF_RECEIVED_DATA_ERROR = 4015
Global Const MQRCCF_RECEIVE_FAILED = 4016
Global Const MQRCCF_CONNECTION_CLOSED = 4017
Global Const MQRCCF_NO_STORAGE = 4018
Global Const MQRCCF_NO_COMMS_MANAGER = 4019
Global Const MQRCCF_LISTENER_NOT_STARTED = 4020
Global Const MQRCCF_BIND_FAILED = 4024
Global Const MQRCCF_CHANNEL_INDOUBT = 4025
Global Const MQRCCF_MQCONN_FAILED = 4026
Global Const MQRCCF_MQOPEN_FAILED = 4027
Global Const MQRCCF_MQGET_FAILED = 4028
Global Const MQRCCF_MQPUT_FAILED = 4029
Global Const MQRCCF_PING_ERROR = 4030
Global Const MQRCCF_CHANNEL_IN_USE = 4031
Global Const MQRCCF_CHANNEL_NOT_FOUND = 4032
Global Const MQRCCF_UNKNOWN_REMOTE_CHANNEL = 4033
Global Const MQRCCF_REMOTE_QM_UNAVAILABLE = 4034
Global Const MQRCCF_REMOTE_QM_TERMINATING = 4035
Global Const MQRCCF_MQINQ_FAILED = 4036
Global Const MQRCCF_NOT_XMIT_Q = 4037
Global Const MQRCCF_CHANNEL_DISABLED = 4038
Global Const MQRCCF_USER_EXIT_NOT_AVAILABLE = 4039
Global Const MQRCCF_COMMIT_FAILED = 4040
Global Const MQRCCF_WRONG_CHANNEL_TYPE = 4041
Global Const MQRCCF_CHANNEL_ALREADY_EXISTS = 4042
Global Const MQRCCF_DATA_TOO_LARGE = 4043
Global Const MQRCCF_CHANNEL_NAME_ERROR = 4044
Global Const MQRCCF_XMIT_Q_NAME_ERROR = 4045
Global Const MQRCCF_MCA_NAME_ERROR = 4047
Global Const MQRCCF_SEND_EXIT_NAME_ERROR = 4048
Global Const MQRCCF_SEC_EXIT_NAME_ERROR = 4049
Global Const MQRCCF_MSG_EXIT_NAME_ERROR = 4050
Global Const MQRCCF_RCV_EXIT_NAME_ERROR = 4051
Global Const MQRCCF_XMIT_Q_NAME_WRONG_TYPE = 4052
Global Const MQRCCF_MCA_NAME_WRONG_TYPE = 4053
Global Const MQRCCF_DISC_INT_WRONG_TYPE = 4054
Global Const MQRCCF_SHORT_RETRY_WRONG_TYPE = 4055
Global Const MQRCCF_SHORT_TIMER_WRONG_TYPE = 4056
Global Const MQRCCF_LONG_RETRY_WRONG_TYPE = 4057
Global Const MQRCCF_LONG_TIMER_WRONG_TYPE = 4058
Global Const MQRCCF_PUT_AUTH_WRONG_TYPE = 4059
Global Const MQRCCF_KEEP_ALIVE_INT_ERROR = 4060
Global Const MQRCCF_MISSING_CONN_NAME = 4061
Global Const MQRCCF_CONN_NAME_ERROR = 4062
Global Const MQRCCF_MQSET_FAILED = 4063
Global Const MQRCCF_CHANNEL_NOT_ACTIVE = 4064
Global Const MQRCCF_TERMINATED_BY_SEC_EXIT = 4065
Global Const MQRCCF_DYNAMIC_Q_SCOPE_ERROR = 4067
Global Const MQRCCF_CELL_DIR_NOT_AVAILABLE = 4068
Global Const MQRCCF_MR_COUNT_ERROR = 4069
Global Const MQRCCF_MR_COUNT_WRONG_TYPE = 4070
Global Const MQRCCF_MR_EXIT_NAME_ERROR = 4071
Global Const MQRCCF_MR_EXIT_NAME_WRONG_TYPE = 4072
Global Const MQRCCF_MR_INTERVAL_ERROR = 4073
Global Const MQRCCF_MR_INTERVAL_WRONG_TYPE = 4074
Global Const MQRCCF_NPM_SPEED_ERROR = 4075
Global Const MQRCCF_NPM_SPEED_WRONG_TYPE = 4076
Global Const MQRCCF_HB_INTERVAL_ERROR = 4077
Global Const MQRCCF_HB_INTERVAL_WRONG_TYPE = 4078
Global Const MQRCCF_CHAD_ERROR = 4079
Global Const MQRCCF_CHAD_WRONG_TYPE = 4080
Global Const MQRCCF_CHAD_EVENT_ERROR = 4081
Global Const MQRCCF_CHAD_EVENT_WRONG_TYPE = 4082
Global Const MQRCCF_CHAD_EXIT_ERROR = 4083
Global Const MQRCCF_CHAD_EXIT_WRONG_TYPE = 4084
Global Const MQRCCF_SUPPRESSED_BY_EXIT = 4085
Global Const MQRCCF_BATCH_INT_ERROR = 4086
Global Const MQRCCF_BATCH_INT_WRONG_TYPE = 4087
Global Const MQRCCF_NET_PRIORITY_ERROR = 4088
Global Const MQRCCF_NET_PRIORITY_WRONG_TYPE = 4089
Global Const MQRCCF_CHANNEL_CLOSED = 4090
Global Const MQRCCF_Q_STATUS_NOT_FOUND = 4091
Global Const MQRCCF_SSL_CIPHER_SPEC_ERROR = 4092
Global Const MQRCCF_SSL_PEER_NAME_ERROR = 4093
Global Const MQRCCF_SSL_CLIENT_AUTH_ERROR = 4094
Global Const MQRCCF_RETAINED_NOT_SUPPORTED = 4095

'****************************************************************'
'*  Values Related to MQCFBF Structure                          *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFBF_STRUC_LENGTH_FIXED = 20

'****************************************************************'
'*  Values Related to MQCFBS Structure                          *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFBS_STRUC_LENGTH_FIXED = 16

'****************************************************************'
'*  Values Related to MQCFGR Structure                          *'
'****************************************************************'
'Structure Length'
Global Const MQCFGR_STRUC_LENGTH = 16

'****************************************************************'
'*  Values Related to MQCFIF Structure                          *'
'****************************************************************'
'Structure Length'
Global Const MQCFIF_STRUC_LENGTH = 20

'****************************************************************'
'*  Values Related to MQCFIL Structure                          *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFIL_STRUC_LENGTH_FIXED = 16

'****************************************************************'
'*  Values Related to MQCFIL64 Structure                        *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFIL64_STRUC_LENGTH_FIXED = 16

'****************************************************************'
'*  Values Related to MQCFIN Structure                          *'
'****************************************************************'
'Structure Length'
Global Const MQCFIN_STRUC_LENGTH = 16

'****************************************************************'
'*  Values Related to MQCFIN64 Structure                        *'
'****************************************************************'
'Structure Length'
Global Const MQCFIN64_STRUC_LENGTH = 24

'****************************************************************'
'*  Values Related to MQCFSF Structure                          *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFSF_STRUC_LENGTH_FIXED = 24

'****************************************************************'
'*  Values Related to MQCFSL Structure                          *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFSL_STRUC_LENGTH_FIXED = 24

'****************************************************************'
'*  Values Related to MQCFST Structure                          *'
'****************************************************************'
'Structure Length (Fixed Part)'
Global Const MQCFST_STRUC_LENGTH_FIXED = 20

'****************************************************************'
'*  Values Related to MQEPH Structure                           *'
'****************************************************************'
'Structure Identifier'
Global Const MQEPH_STRUC_ID = "EPH "

'Structure Length (Fixed Part)'
Global Const MQEPH_STRUC_LENGTH_FIXED = 68

'Structure Version Number'
Global Const MQEPH_VERSION_1 = 1
Global Const MQEPH_CURRENT_VERSION = 1

'Flags'
Global Const MQEPH_NONE = &H0
Global Const MQEPH_CCSID_EMBEDDED = &H1

'****************************************************************'
'*  Values Related to All Structures                            *'
'****************************************************************'
'String Lengths'
Global Const MQ_ARCHIVE_PFX_LENGTH = 36
Global Const MQ_ARCHIVE_UNIT_LENGTH = 8
Global Const MQ_ASID_LENGTH = 4
Global Const MQ_AUTH_PROFILE_NAME_LENGTH = 48
Global Const MQ_CF_LEID_LENGTH = 12
Global Const MQ_COMMAND_MQSC_LENGTH = 32768
Global Const MQ_DATA_SET_NAME_LENGTH = 44
Global Const MQ_DB2_NAME_LENGTH = 4
Global Const MQ_DSG_NAME_LENGTH = 8
Global Const MQ_ENTITY_NAME_LENGTH = 64
Global Const MQ_ENV_INFO_LENGTH = 96
Global Const MQ_IP_ADDRESS_LENGTH = 48
Global Const MQ_LOG_CORREL_ID_LENGTH = 8
Global Const MQ_LOG_EXTENT_NAME_LENGTH = 24
Global Const MQ_LOG_PATH_LENGTH = 1024
Global Const MQ_LRSN_LENGTH = 12
Global Const MQ_ORIGIN_NAME_LENGTH = 8
Global Const MQ_PSB_NAME_LENGTH = 8
Global Const MQ_PST_ID_LENGTH = 8
Global Const MQ_Q_MGR_CPF_LENGTH = 4
Global Const MQ_RESPONSE_ID_LENGTH = 24
Global Const MQ_RBA_LENGTH = 12
Global Const MQ_SECURITY_PROFILE_LENGTH = 40
Global Const MQ_SERVICE_COMPONENT_LENGTH = 48
Global Const MQ_SYSP_SERVICE_LENGTH = 32
Global Const MQ_SYSTEM_NAME_LENGTH = 8
Global Const MQ_TASK_NUMBER_LENGTH = 8
Global Const MQ_TPIPE_PFX_LENGTH = 4
Global Const MQ_UOW_ID_LENGTH = 256
Global Const MQ_VOLSER_LENGTH = 6

'Filter Operators'
Global Const MQCFOP_LESS = 1
Global Const MQCFOP_EQUAL = 2
Global Const MQCFOP_GREATER = 4
Global Const MQCFOP_NOT_LESS = 6
Global Const MQCFOP_NOT_EQUAL = 5
Global Const MQCFOP_NOT_GREATER = 3
Global Const MQCFOP_LIKE = 18
Global Const MQCFOP_NOT_LIKE = 21
Global Const MQCFOP_CONTAINS = 10
Global Const MQCFOP_EXCLUDES = 13
Global Const MQCFOP_CONTAINS_GEN = 26
Global Const MQCFOP_EXCLUDES_GEN = 29

'Structure Type'
Global Const MQCFT_NONE = 0
Global Const MQCFT_COMMAND = 1
Global Const MQCFT_RESPONSE = 2
Global Const MQCFT_INTEGER = 3
Global Const MQCFT_STRING = 4
Global Const MQCFT_INTEGER_LIST = 5
Global Const MQCFT_STRING_LIST = 6
Global Const MQCFT_EVENT = 7
Global Const MQCFT_USER = 8
Global Const MQCFT_BYTE_STRING = 9
Global Const MQCFT_TRACE_ROUTE = 10
Global Const MQCFT_REPORT = 12
Global Const MQCFT_INTEGER_FILTER = 13
Global Const MQCFT_STRING_FILTER = 14
Global Const MQCFT_BYTE_STRING_FILTER = 15
Global Const MQCFT_COMMAND_XR = 16
Global Const MQCFT_XR_MSG = 17
Global Const MQCFT_XR_ITEM = 18
Global Const MQCFT_XR_SUMMARY = 19
Global Const MQCFT_GROUP = 20
Global Const MQCFT_STATISTICS = 21
Global Const MQCFT_ACCOUNTING = 22
Global Const MQCFT_INTEGER64 = 23
Global Const MQCFT_INTEGER64_LIST = 25

'****************************************************************'
'*  Values Related to Byte Parameter Structures                 *'
'****************************************************************'

'Byte Parameter Types'
Global Const MQBACF_FIRST = 7001
Global Const MQBACF_EVENT_ACCOUNTING_TOKEN = 7001
Global Const MQBACF_EVENT_SECURITY_ID = 7002
Global Const MQBACF_RESPONSE_SET = 7003
Global Const MQBACF_RESPONSE_ID = 7004
Global Const MQBACF_EXTERNAL_UOW_ID = 7005
Global Const MQBACF_CONNECTION_ID = 7006
Global Const MQBACF_GENERIC_CONNECTION_ID = 7007
Global Const MQBACF_ORIGIN_UOW_ID = 7008
Global Const MQBACF_Q_MGR_UOW_ID = 7009
Global Const MQBACF_ACCOUNTING_TOKEN = 7010
Global Const MQBACF_CORREL_ID = 7011
Global Const MQBACF_GROUP_ID = 7012
Global Const MQBACF_MSG_ID = 7013
Global Const MQBACF_CF_LEID = 7014
Global Const MQBACF_LAST_USED = 7014

'****************************************************************'
'*  Values Related to Integer Parameter Structures              *'
'****************************************************************'

'Integer Monitoring Parameter Types'
Global Const MQIAMO_FIRST = 701
Global Const MQIAMO_AVG_BATCH_SIZE = 702
Global Const MQIAMO_AVG_Q_TIME = 703
Global Const MQIAMO_BACKOUTS = 704
Global Const MQIAMO_BROWSES = 705
Global Const MQIAMO_BROWSE_MAX_BYTES = 706
Global Const MQIAMO_BROWSE_MIN_BYTES = 707
Global Const MQIAMO_BROWSES_FAILED = 708
Global Const MQIAMO_CLOSES = 709
Global Const MQIAMO_COMMITS = 710
Global Const MQIAMO_COMMITS_FAILED = 711
Global Const MQIAMO_CONNS = 712
Global Const MQIAMO_CONNS_MAX = 713
Global Const MQIAMO_DISCS = 714
Global Const MQIAMO_DISCS_IMPLICIT = 715
Global Const MQIAMO_DISC_TYPE = 716
Global Const MQIAMO_EXIT_TIME_AVG = 717
Global Const MQIAMO_EXIT_TIME_MAX = 718
Global Const MQIAMO_EXIT_TIME_MIN = 719
Global Const MQIAMO_FULL_BATCHES = 720
Global Const MQIAMO_GENERATED_MSGS = 721
Global Const MQIAMO_GETS = 722
Global Const MQIAMO_GET_MAX_BYTES = 723
Global Const MQIAMO_GET_MIN_BYTES = 724
Global Const MQIAMO_GETS_FAILED = 725
Global Const MQIAMO_INCOMPLETE_BATCHES = 726
Global Const MQIAMO_INQS = 727
Global Const MQIAMO_MSGS = 728
Global Const MQIAMO_NET_TIME_AVG = 729
Global Const MQIAMO_NET_TIME_MAX = 730
Global Const MQIAMO_NET_TIME_MIN = 731
Global Const MQIAMO_OBJECT_COUNT = 732
Global Const MQIAMO_OPENS = 733
Global Const MQIAMO_PUT1S = 734
Global Const MQIAMO_PUTS = 735
Global Const MQIAMO_PUT_MAX_BYTES = 736
Global Const MQIAMO_PUT_MIN_BYTES = 737
Global Const MQIAMO_PUT_RETRIES = 738
Global Const MQIAMO_Q_MAX_DEPTH = 739
Global Const MQIAMO_Q_MIN_DEPTH = 740
Global Const MQIAMO_Q_TIME_AVG = 741
Global Const MQIAMO_Q_TIME_MAX = 742
Global Const MQIAMO_Q_TIME_MIN = 743
Global Const MQIAMO_SETS = 744
Global Const MQIAMO_CONNS_FAILED = 749
Global Const MQIAMO_OPENS_FAILED = 751
Global Const MQIAMO_INQS_FAILED = 752
Global Const MQIAMO_SETS_FAILED = 753
Global Const MQIAMO_PUTS_FAILED = 754
Global Const MQIAMO_PUT1S_FAILED = 755
Global Const MQIAMO_CLOSES_FAILED = 757
Global Const MQIAMO_MSGS_EXPIRED = 758
Global Const MQIAMO_MSGS_NOT_QUEUED = 759
Global Const MQIAMO_MSGS_PURGED = 760
Global Const MQIAMO_LAST_USED = 760

'64-bit Integer Monitoring Parameter Types'
Global Const MQIAMO64_BROWSE_BYTES = 745
Global Const MQIAMO64_BYTES = 746
Global Const MQIAMO64_GET_BYTES = 747
Global Const MQIAMO64_PUT_BYTES = 748

'Integer Parameter Types'
Global Const MQIACF_FIRST = 1001
Global Const MQIACF_Q_MGR_ATTRS = 1001
Global Const MQIACF_Q_ATTRS = 1002
Global Const MQIACF_PROCESS_ATTRS = 1003
Global Const MQIACF_NAMELIST_ATTRS = 1004
Global Const MQIACF_FORCE = 1005
Global Const MQIACF_REPLACE = 1006
Global Const MQIACF_PURGE = 1007
Global Const MQIACF_QUIESCE = 1008
Global Const MQIACF_MODE = 1008
Global Const MQIACF_ALL = 1009
Global Const MQIACF_EVENT_APPL_TYPE = 1010
Global Const MQIACF_EVENT_ORIGIN = 1011
Global Const MQIACF_PARAMETER_ID = 1012
Global Const MQIACF_ERROR_ID = 1013
Global Const MQIACF_ERROR_IDENTIFIER = 1013
Global Const MQIACF_SELECTOR = 1014
Global Const MQIACF_CHANNEL_ATTRS = 1015
Global Const MQIACF_OBJECT_TYPE = 1016
Global Const MQIACF_ESCAPE_TYPE = 1017
Global Const MQIACF_ERROR_OFFSET = 1018
Global Const MQIACF_AUTH_INFO_ATTRS = 1019
Global Const MQIACF_REASON_QUALIFIER = 1020
Global Const MQIACF_COMMAND = 1021
Global Const MQIACF_OPEN_OPTIONS = 1022
Global Const MQIACF_OPEN_TYPE = 1023
Global Const MQIACF_PROCESS_ID = 1024
Global Const MQIACF_THREAD_ID = 1025
Global Const MQIACF_Q_STATUS_ATTRS = 1026
Global Const MQIACF_UNCOMMITTED_MSGS = 1027
Global Const MQIACF_HANDLE_STATE = 1028
Global Const MQIACF_AUX_ERROR_DATA_INT_1 = 1070
Global Const MQIACF_AUX_ERROR_DATA_INT_2 = 1071
Global Const MQIACF_CONV_REASON_CODE = 1072
Global Const MQIACF_BRIDGE_TYPE = 1073
Global Const MQIACF_INQUIRY = 1074
Global Const MQIACF_WAIT_INTERVAL = 1075
Global Const MQIACF_OPTIONS = 1076
Global Const MQIACF_BROKER_OPTIONS = 1077
Global Const MQIACF_REFRESH_TYPE = 1078
Global Const MQIACF_SEQUENCE_NUMBER = 1079
Global Const MQIACF_INTEGER_DATA = 1080
Global Const MQIACF_REGISTRATION_OPTIONS = 1081
Global Const MQIACF_PUBLICATION_OPTIONS = 1082
Global Const MQIACF_CLUSTER_INFO = 1083
Global Const MQIACF_Q_MGR_DEFINITION_TYPE = 1084
Global Const MQIACF_Q_MGR_TYPE = 1085
Global Const MQIACF_ACTION = 1086
Global Const MQIACF_SUSPEND = 1087
Global Const MQIACF_BROKER_COUNT = 1088
Global Const MQIACF_APPL_COUNT = 1089
Global Const MQIACF_ANONYMOUS_COUNT = 1090
Global Const MQIACF_REG_REG_OPTIONS = 1091
Global Const MQIACF_DELETE_OPTIONS = 1092
Global Const MQIACF_CLUSTER_Q_MGR_ATTRS = 1093
Global Const MQIACF_REFRESH_INTERVAL = 1094
Global Const MQIACF_REFRESH_REPOSITORY = 1095
Global Const MQIACF_REMOVE_QUEUES = 1096
Global Const MQIACF_OPEN_INPUT_TYPE = 1098
Global Const MQIACF_OPEN_OUTPUT = 1099
Global Const MQIACF_OPEN_SET = 1100
Global Const MQIACF_OPEN_INQUIRE = 1101
Global Const MQIACF_OPEN_BROWSE = 1102
Global Const MQIACF_Q_STATUS_TYPE = 1103
Global Const MQIACF_Q_HANDLE = 1104
Global Const MQIACF_Q_STATUS = 1105
Global Const MQIACF_SECURITY_TYPE = 1106
Global Const MQIACF_CONNECTION_ATTRS = 1107
Global Const MQIACF_CONNECT_OPTIONS = 1108
Global Const MQIACF_CONN_INFO_TYPE = 1110
Global Const MQIACF_CONN_INFO_CONN = 1111
Global Const MQIACF_CONN_INFO_HANDLE = 1112
Global Const MQIACF_CONN_INFO_ALL = 1113
Global Const MQIACF_AUTH_PROFILE_ATTRS = 1114
Global Const MQIACF_AUTHORIZATION_LIST = 1115
Global Const MQIACF_AUTH_ADD_AUTHS = 1116
Global Const MQIACF_AUTH_REMOVE_AUTHS = 1117
Global Const MQIACF_ENTITY_TYPE = 1118
Global Const MQIACF_COMMAND_INFO = 1120
Global Const MQIACF_CMDSCOPE_Q_MGR_COUNT = 1121
Global Const MQIACF_Q_MGR_SYSTEM = 1122
Global Const MQIACF_Q_MGR_EVENT = 1123
Global Const MQIACF_Q_MGR_DQM = 1124
Global Const MQIACF_Q_MGR_CLUSTER = 1125
Global Const MQIACF_QSG_DISPS = 1126
Global Const MQIACF_UOW_STATE = 1128
Global Const MQIACF_SECURITY_ITEM = 1129
Global Const MQIACF_CF_STRUC_STATUS = 1130
Global Const MQIACF_UOW_TYPE = 1132
Global Const MQIACF_CF_STRUC_ATTRS = 1133
Global Const MQIACF_EXCLUDE_INTERVAL = 1134
Global Const MQIACF_CF_STATUS_TYPE = 1135
Global Const MQIACF_CF_STATUS_SUMMARY = 1136
Global Const MQIACF_CF_STATUS_CONNECT = 1137
Global Const MQIACF_CF_STATUS_BACKUP = 1138
Global Const MQIACF_CF_STRUC_TYPE = 1139
Global Const MQIACF_CF_STRUC_SIZE_MAX = 1140
Global Const MQIACF_CF_STRUC_SIZE_USED = 1141
Global Const MQIACF_CF_STRUC_ENTRIES_MAX = 1142
Global Const MQIACF_CF_STRUC_ENTRIES_USED = 1143
Global Const MQIACF_CF_STRUC_BACKUP_SIZE = 1144
Global Const MQIACF_MOVE_TYPE = 1145
Global Const MQIACF_MOVE_TYPE_MOVE = 1146
Global Const MQIACF_MOVE_TYPE_ADD = 1147
Global Const MQIACF_Q_MGR_NUMBER = 1148
Global Const MQIACF_Q_MGR_STATUS = 1149
Global Const MQIACF_DB2_CONN_STATUS = 1150
Global Const MQIACF_SECURITY_ATTRS = 1151
Global Const MQIACF_SECURITY_TIMEOUT = 1152
Global Const MQIACF_SECURITY_INTERVAL = 1153
Global Const MQIACF_SECURITY_SWITCH = 1154
Global Const MQIACF_SECURITY_SETTING = 1155
Global Const MQIACF_STORAGE_CLASS_ATTRS = 1156
Global Const MQIACF_USAGE_TYPE = 1157
Global Const MQIACF_BUFFER_POOL_ID = 1158
Global Const MQIACF_USAGE_TOTAL_PAGES = 1159
Global Const MQIACF_USAGE_UNUSED_PAGES = 1160
Global Const MQIACF_USAGE_PERSIST_PAGES = 1161
Global Const MQIACF_USAGE_NONPERSIST_PAGES = 1162
Global Const MQIACF_USAGE_RESTART_EXTENTS = 1163
Global Const MQIACF_USAGE_EXPAND_COUNT = 1164
Global Const MQIACF_PAGESET_STATUS = 1165
Global Const MQIACF_USAGE_TOTAL_BUFFERS = 1166
Global Const MQIACF_USAGE_DATA_SET_TYPE = 1167
Global Const MQIACF_USAGE_PAGESET = 1168
Global Const MQIACF_USAGE_DATA_SET = 1169
Global Const MQIACF_USAGE_BUFFER_POOL = 1170
Global Const MQIACF_MOVE_COUNT = 1171
Global Const MQIACF_EXPIRY_Q_COUNT = 1172
Global Const MQIACF_CONFIGURATION_OBJECTS = 1173
Global Const MQIACF_CONFIGURATION_EVENTS = 1174
Global Const MQIACF_SYSP_TYPE = 1175
Global Const MQIACF_SYSP_DEALLOC_INTERVAL = 1176
Global Const MQIACF_SYSP_MAX_ARCHIVE = 1177
Global Const MQIACF_SYSP_MAX_READ_TAPES = 1178
Global Const MQIACF_SYSP_IN_BUFFER_SIZE = 1179
Global Const MQIACF_SYSP_OUT_BUFFER_SIZE = 1180
Global Const MQIACF_SYSP_OUT_BUFFER_COUNT = 1181
Global Const MQIACF_SYSP_ARCHIVE = 1182
Global Const MQIACF_SYSP_DUAL_ACTIVE = 1183
Global Const MQIACF_SYSP_DUAL_ARCHIVE = 1184
Global Const MQIACF_SYSP_DUAL_BSDS = 1185
Global Const MQIACF_SYSP_MAX_CONNS = 1186
Global Const MQIACF_SYSP_MAX_CONNS_FORE = 1187
Global Const MQIACF_SYSP_MAX_CONNS_BACK = 1188
Global Const MQIACF_SYSP_EXIT_INTERVAL = 1189
Global Const MQIACF_SYSP_EXIT_TASKS = 1190
Global Const MQIACF_SYSP_CHKPOINT_COUNT = 1191
Global Const MQIACF_SYSP_OTMA_INTERVAL = 1192
Global Const MQIACF_SYSP_Q_INDEX_DEFER = 1193
Global Const MQIACF_SYSP_DB2_TASKS = 1194
Global Const MQIACF_SYSP_RESLEVEL_AUDIT = 1195
Global Const MQIACF_SYSP_ROUTING_CODE = 1196
Global Const MQIACF_SYSP_SMF_ACCOUNTING = 1197
Global Const MQIACF_SYSP_SMF_STATS = 1198
Global Const MQIACF_SYSP_SMF_INTERVAL = 1199
Global Const MQIACF_SYSP_TRACE_CLASS = 1200
Global Const MQIACF_SYSP_TRACE_SIZE = 1201
Global Const MQIACF_SYSP_WLM_INTERVAL = 1202
Global Const MQIACF_SYSP_ALLOC_UNIT = 1203
Global Const MQIACF_SYSP_ARCHIVE_RETAIN = 1204
Global Const MQIACF_SYSP_ARCHIVE_WTOR = 1205
Global Const MQIACF_SYSP_BLOCK_SIZE = 1206
Global Const MQIACF_SYSP_CATALOG = 1207
Global Const MQIACF_SYSP_COMPACT = 1208
Global Const MQIACF_SYSP_ALLOC_PRIMARY = 1209
Global Const MQIACF_SYSP_ALLOC_SECONDARY = 1210
Global Const MQIACF_SYSP_PROTECT = 1211
Global Const MQIACF_SYSP_QUIESCE_INTERVAL = 1212
Global Const MQIACF_SYSP_TIMESTAMP = 1213
Global Const MQIACF_SYSP_UNIT_ADDRESS = 1214
Global Const MQIACF_SYSP_UNIT_STATUS = 1215
Global Const MQIACF_SYSP_LOG_COPY = 1216
Global Const MQIACF_SYSP_LOG_USED = 1217
Global Const MQIACF_SYSP_LOG_SUSPEND = 1218
Global Const MQIACF_SYSP_OFFLOAD_STATUS = 1219
Global Const MQIACF_SYSP_TOTAL_LOGS = 1220
Global Const MQIACF_SYSP_FULL_LOGS = 1221
Global Const MQIACF_LISTENER_ATTRS = 1222
Global Const MQIACF_LISTENER_STATUS_ATTRS = 1223
Global Const MQIACF_SERVICE_ATTRS = 1224
Global Const MQIACF_SERVICE_STATUS_ATTRS = 1225
Global Const MQIACF_Q_TIME_INDICATOR = 1226
Global Const MQIACF_OLDEST_MSG_AGE = 1227
Global Const MQIACF_AUTH_OPTIONS = 1228
Global Const MQIACF_Q_MGR_STATUS_ATTRS = 1229
Global Const MQIACF_CONNECTION_COUNT = 1230
Global Const MQIACF_Q_MGR_FACILITY = 1231
Global Const MQIACF_CHINIT_STATUS = 1232
Global Const MQIACF_CMD_SERVER_STATUS = 1233
Global Const MQIACF_ROUTE_DETAIL = 1234
Global Const MQIACF_RECORDED_ACTIVITIES = 1235
Global Const MQIACF_MAX_ACTIVITIES = 1236
Global Const MQIACF_DISCONTINUITY_COUNT = 1237
Global Const MQIACF_ROUTE_ACCUMULATION = 1238
Global Const MQIACF_ROUTE_DELIVERY = 1239
Global Const MQIACF_OPERATION_TYPE = 1240
Global Const MQIACF_BACKOUT_COUNT = 1241
Global Const MQIACF_COMP_CODE = 1242
Global Const MQIACF_ENCODING = 1243
Global Const MQIACF_EXPIRY = 1244
Global Const MQIACF_FEEDBACK = 1245
Global Const MQIACF_MSG_FLAGS = 1247
Global Const MQIACF_MSG_LENGTH = 1248
Global Const MQIACF_MSG_TYPE = 1249
Global Const MQIACF_OFFSET = 1250
Global Const MQIACF_ORIGINAL_LENGTH = 1251
Global Const MQIACF_PERSISTENCE = 1252
Global Const MQIACF_PRIORITY = 1253
Global Const MQIACF_REASON_CODE = 1254
Global Const MQIACF_REPORT = 1255
Global Const MQIACF_VERSION = 1256
Global Const MQIACF_UNRECORDED_ACTIVITIES = 1257
Global Const MQIACF_MONITORING = 1258
Global Const MQIACF_ROUTE_FORWARDING = 1259
Global Const MQIACF_SERVICE_STATUS = 1260
Global Const MQIACF_Q_TYPES = 1261
Global Const MQIACF_USER_ID_SUPPORT = 1262
Global Const MQIACF_INTERFACE_VERSION = 1263
Global Const MQIACF_AUTH_SERVICE_ATTRS = 1264
Global Const MQIACF_USAGE_EXPAND_TYPE = 1265
Global Const MQIACF_SYSP_CLUSTER_CACHE = 1266
Global Const MQIACF_SYSP_DB2_BLOB_TASKS = 1267
Global Const MQIACF_SYSP_WLM_INT_UNITS = 1268
Global Const MQIACF_LAST_USED = 1268

'Integer Channel Types'
Global Const MQIACH_FIRST = 1501
Global Const MQIACH_XMIT_PROTOCOL_TYPE = 1501
Global Const MQIACH_BATCH_SIZE = 1502
Global Const MQIACH_DISC_INTERVAL = 1503
Global Const MQIACH_SHORT_TIMER = 1504
Global Const MQIACH_SHORT_RETRY = 1505
Global Const MQIACH_LONG_TIMER = 1506
Global Const MQIACH_LONG_RETRY = 1507
Global Const MQIACH_PUT_AUTHORITY = 1508
Global Const MQIACH_SEQUENCE_NUMBER_WRAP = 1509
Global Const MQIACH_MAX_MSG_LENGTH = 1510
Global Const MQIACH_CHANNEL_TYPE = 1511
Global Const MQIACH_DATA_COUNT = 1512
Global Const MQIACH_NAME_COUNT = 1513
Global Const MQIACH_MSG_SEQUENCE_NUMBER = 1514
Global Const MQIACH_DATA_CONVERSION = 1515
Global Const MQIACH_IN_DOUBT = 1516
Global Const MQIACH_MCA_TYPE = 1517
Global Const MQIACH_SESSION_COUNT = 1518
Global Const MQIACH_ADAPTER = 1519
Global Const MQIACH_COMMAND_COUNT = 1520
Global Const MQIACH_SOCKET = 1521
Global Const MQIACH_PORT = 1522
Global Const MQIACH_CHANNEL_INSTANCE_TYPE = 1523
Global Const MQIACH_CHANNEL_INSTANCE_ATTRS = 1524
Global Const MQIACH_CHANNEL_ERROR_DATA = 1525
Global Const MQIACH_CHANNEL_TABLE = 1526
Global Const MQIACH_CHANNEL_STATUS = 1527
Global Const MQIACH_INDOUBT_STATUS = 1528
Global Const MQIACH_LAST_SEQ_NUMBER = 1529
Global Const MQIACH_LAST_SEQUENCE_NUMBER = 1529
Global Const MQIACH_CURRENT_MSGS = 1531
Global Const MQIACH_CURRENT_SEQ_NUMBER = 1532
Global Const MQIACH_CURRENT_SEQUENCE_NUMBER = 1532
Global Const MQIACH_SSL_RETURN_CODE = 1533
Global Const MQIACH_MSGS = 1534
Global Const MQIACH_BYTES_SENT = 1535
Global Const MQIACH_BYTES_RCVD = 1536
Global Const MQIACH_BYTES_RECEIVED = 1536
Global Const MQIACH_BATCHES = 1537
Global Const MQIACH_BUFFERS_SENT = 1538
Global Const MQIACH_BUFFERS_RCVD = 1539
Global Const MQIACH_BUFFERS_RECEIVED = 1539
Global Const MQIACH_LONG_RETRIES_LEFT = 1540
Global Const MQIACH_SHORT_RETRIES_LEFT = 1541
Global Const MQIACH_MCA_STATUS = 1542
Global Const MQIACH_STOP_REQUESTED = 1543
Global Const MQIACH_MR_COUNT = 1544
Global Const MQIACH_MR_INTERVAL = 1545
Global Const MQIACH_NPM_SPEED = 1562
Global Const MQIACH_HB_INTERVAL = 1563
Global Const MQIACH_BATCH_INTERVAL = 1564
Global Const MQIACH_NETWORK_PRIORITY = 1565
Global Const MQIACH_KEEP_ALIVE_INTERVAL = 1566
Global Const MQIACH_BATCH_HB = 1567
Global Const MQIACH_SSL_CLIENT_AUTH = 1568
Global Const MQIACH_ALLOC_RETRY = 1570
Global Const MQIACH_ALLOC_FAST_TIMER = 1571
Global Const MQIACH_ALLOC_SLOW_TIMER = 1572
Global Const MQIACH_DISC_RETRY = 1573
Global Const MQIACH_PORT_NUMBER = 1574
Global Const MQIACH_HDR_COMPRESSION = 1575
Global Const MQIACH_MSG_COMPRESSION = 1576
Global Const MQIACH_CLWL_CHANNEL_RANK = 1577
Global Const MQIACH_CLWL_CHANNEL_PRIORITY = 1578
Global Const MQIACH_CLWL_CHANNEL_WEIGHT = 1579
Global Const MQIACH_CHANNEL_DISP = 1580
Global Const MQIACH_INBOUND_DISP = 1581
Global Const MQIACH_CHANNEL_TYPES = 1582
Global Const MQIACH_ADAPS_STARTED = 1583
Global Const MQIACH_ADAPS_MAX = 1584
Global Const MQIACH_DISPS_STARTED = 1585
Global Const MQIACH_DISPS_MAX = 1586
Global Const MQIACH_SSLTASKS_STARTED = 1587
Global Const MQIACH_SSLTASKS_MAX = 1588
Global Const MQIACH_CURRENT_CHL = 1589
Global Const MQIACH_CURRENT_CHL_MAX = 1590
Global Const MQIACH_CURRENT_CHL_TCP = 1591
Global Const MQIACH_CURRENT_CHL_LU62 = 1592
Global Const MQIACH_ACTIVE_CHL = 1593
Global Const MQIACH_ACTIVE_CHL_MAX = 1594
Global Const MQIACH_ACTIVE_CHL_PAUSED = 1595
Global Const MQIACH_ACTIVE_CHL_STARTED = 1596
Global Const MQIACH_ACTIVE_CHL_STOPPED = 1597
Global Const MQIACH_ACTIVE_CHL_RETRY = 1598
Global Const MQIACH_LISTENER_STATUS = 1599
Global Const MQIACH_SHARED_CHL_RESTART = 1600
Global Const MQIACH_LISTENER_CONTROL = 1601
Global Const MQIACH_BACKLOG = 1602
Global Const MQIACH_XMITQ_TIME_INDICATOR = 1604
Global Const MQIACH_NETWORK_TIME_INDICATOR = 1605
Global Const MQIACH_EXIT_TIME_INDICATOR = 1606
Global Const MQIACH_BATCH_SIZE_INDICATOR = 1607
Global Const MQIACH_XMITQ_MSGS_AVAILABLE = 1608
Global Const MQIACH_CHANNEL_SUBSTATE = 1609
Global Const MQIACH_SSL_KEY_RESETS = 1610
Global Const MQIACH_COMPRESSION_RATE = 1611
Global Const MQIACH_COMPRESSION_TIME = 1612
Global Const MQIACH_MAX_XMIT_SIZE = 1613
Global Const MQIACH_LAST_USED = 1613

'****************************************************************'
'*  Values Related to Character Parameter Structures            *'
'****************************************************************'

'Character Monitoring Parameter Types'
Global Const MQCAMO_FIRST = 2701
Global Const MQCAMO_CLOSE_DATE = 2701
Global Const MQCAMO_CLOSE_TIME = 2702
Global Const MQCAMO_CONN_DATE = 2703
Global Const MQCAMO_CONN_TIME = 2704
Global Const MQCAMO_DISC_DATE = 2705
Global Const MQCAMO_DISC_TIME = 2706
Global Const MQCAMO_END_DATE = 2707
Global Const MQCAMO_END_TIME = 2708
Global Const MQCAMO_OPEN_DATE = 2709
Global Const MQCAMO_OPEN_TIME = 2710
Global Const MQCAMO_START_DATE = 2711
Global Const MQCAMO_START_TIME = 2712
Global Const MQCAMO_LAST_USED = 2712

'Character Parameter Types'
Global Const MQCACF_FIRST = 3001
Global Const MQCACF_FROM_Q_NAME = 3001
Global Const MQCACF_TO_Q_NAME = 3002
Global Const MQCACF_FROM_PROCESS_NAME = 3003
Global Const MQCACF_TO_PROCESS_NAME = 3004
Global Const MQCACF_FROM_NAMELIST_NAME = 3005
Global Const MQCACF_TO_NAMELIST_NAME = 3006
Global Const MQCACF_FROM_CHANNEL_NAME = 3007
Global Const MQCACF_TO_CHANNEL_NAME = 3008
Global Const MQCACF_FROM_AUTH_INFO_NAME = 3009
Global Const MQCACF_TO_AUTH_INFO_NAME = 3010
Global Const MQCACF_Q_NAMES = 3011
Global Const MQCACF_PROCESS_NAMES = 3012
Global Const MQCACF_NAMELIST_NAMES = 3013
Global Const MQCACF_ESCAPE_TEXT = 3014
Global Const MQCACF_LOCAL_Q_NAMES = 3015
Global Const MQCACF_MODEL_Q_NAMES = 3016
Global Const MQCACF_ALIAS_Q_NAMES = 3017
Global Const MQCACF_REMOTE_Q_NAMES = 3018
Global Const MQCACF_SENDER_CHANNEL_NAMES = 3019
Global Const MQCACF_SERVER_CHANNEL_NAMES = 3020
Global Const MQCACF_REQUESTER_CHANNEL_NAMES = 3021
Global Const MQCACF_RECEIVER_CHANNEL_NAMES = 3022
Global Const MQCACF_OBJECT_Q_MGR_NAME = 3023
Global Const MQCACF_APPL_NAME = 3024
Global Const MQCACF_USER_IDENTIFIER = 3025
Global Const MQCACF_AUX_ERROR_DATA_STR_1 = 3026
Global Const MQCACF_AUX_ERROR_DATA_STR_2 = 3027
Global Const MQCACF_AUX_ERROR_DATA_STR_3 = 3028
Global Const MQCACF_BRIDGE_NAME = 3029
Global Const MQCACF_STREAM_NAME = 3030
Global Const MQCACF_TOPIC = 3031
Global Const MQCACF_PARENT_Q_MGR_NAME = 3032
Global Const MQCACF_CORREL_ID = 3033
Global Const MQCACF_PUBLISH_TIMESTAMP = 3034
Global Const MQCACF_STRING_DATA = 3035
Global Const MQCACF_SUPPORTED_STREAM_NAME = 3036
Global Const MQCACF_REG_TOPIC = 3037
Global Const MQCACF_REG_TIME = 3038
Global Const MQCACF_REG_USER_ID = 3039
Global Const MQCACF_CHILD_Q_MGR_NAME = 3040
Global Const MQCACF_REG_STREAM_NAME = 3041
Global Const MQCACF_REG_Q_MGR_NAME = 3042
Global Const MQCACF_REG_Q_NAME = 3043
Global Const MQCACF_REG_CORREL_ID = 3044
Global Const MQCACF_EVENT_USER_ID = 3045
Global Const MQCACF_OBJECT_NAME = 3046
Global Const MQCACF_EVENT_Q_MGR = 3047
Global Const MQCACF_AUTH_INFO_NAMES = 3048
Global Const MQCACF_EVENT_APPL_IDENTITY = 3049
Global Const MQCACF_EVENT_APPL_NAME = 3050
Global Const MQCACF_EVENT_APPL_ORIGIN = 3051
Global Const MQCACF_SUBSCRIPTION_NAME = 3052
Global Const MQCACF_REG_SUB_NAME = 3053
Global Const MQCACF_SUBSCRIPTION_IDENTITY = 3054
Global Const MQCACF_REG_SUB_IDENTITY = 3055
Global Const MQCACF_SUBSCRIPTION_USER_DATA = 3056
Global Const MQCACF_REG_SUB_USER_DATA = 3057
Global Const MQCACF_APPL_TAG = 3058
Global Const MQCACF_DATA_SET_NAME = 3059
Global Const MQCACF_UOW_START_DATE = 3060
Global Const MQCACF_UOW_START_TIME = 3061
Global Const MQCACF_UOW_LOG_START_DATE = 3062
Global Const MQCACF_UOW_LOG_START_TIME = 3063
Global Const MQCACF_UOW_LOG_EXTENT_NAME = 3064
Global Const MQCACF_PRINCIPAL_ENTITY_NAMES = 3065
Global Const MQCACF_GROUP_ENTITY_NAMES = 3066
Global Const MQCACF_AUTH_PROFILE_NAME = 3067
Global Const MQCACF_ENTITY_NAME = 3068
Global Const MQCACF_SERVICE_COMPONENT = 3069
Global Const MQCACF_RESPONSE_Q_MGR_NAME = 3070
Global Const MQCACF_CURRENT_LOG_EXTENT_NAME = 3071
Global Const MQCACF_RESTART_LOG_EXTENT_NAME = 3072
Global Const MQCACF_MEDIA_LOG_EXTENT_NAME = 3073
Global Const MQCACF_LOG_PATH = 3074
Global Const MQCACF_COMMAND_MQSC = 3075
Global Const MQCACF_Q_MGR_CPF = 3076
Global Const MQCACF_USAGE_LOG_RBA = 3078
Global Const MQCACF_USAGE_LOG_LRSN = 3079
Global Const MQCACF_COMMAND_SCOPE = 3080
Global Const MQCACF_ASID = 3081
Global Const MQCACF_PSB_NAME = 3082
Global Const MQCACF_PST_ID = 3083
Global Const MQCACF_TASK_NUMBER = 3084
Global Const MQCACF_TRANSACTION_ID = 3085
Global Const MQCACF_Q_MGR_UOW_ID = 3086
Global Const MQCACF_ORIGIN_NAME = 3088
Global Const MQCACF_ENV_INFO = 3089
Global Const MQCACF_SECURITY_PROFILE = 3090
Global Const MQCACF_CONFIGURATION_DATE = 3091
Global Const MQCACF_CONFIGURATION_TIME = 3092
Global Const MQCACF_FROM_CF_STRUC_NAME = 3093
Global Const MQCACF_TO_CF_STRUC_NAME = 3094
Global Const MQCACF_CF_STRUC_NAMES = 3095
Global Const MQCACF_FAIL_DATE = 3096
Global Const MQCACF_FAIL_TIME = 3097
Global Const MQCACF_BACKUP_DATE = 3098
Global Const MQCACF_BACKUP_TIME = 3099
Global Const MQCACF_SYSTEM_NAME = 3100
Global Const MQCACF_CF_STRUC_BACKUP_START = 3101
Global Const MQCACF_CF_STRUC_BACKUP_END = 3102
Global Const MQCACF_CF_STRUC_LOG_Q_MGRS = 3103
Global Const MQCACF_FROM_STORAGE_CLASS = 3104
Global Const MQCACF_TO_STORAGE_CLASS = 3105
Global Const MQCACF_STORAGE_CLASS_NAMES = 3106
Global Const MQCACF_DSG_NAME = 3108
Global Const MQCACF_DB2_NAME = 3109
Global Const MQCACF_SYSP_CMD_USER_ID = 3110
Global Const MQCACF_SYSP_OTMA_GROUP = 3111
Global Const MQCACF_SYSP_OTMA_MEMBER = 3112
Global Const MQCACF_SYSP_OTMA_DRU_EXIT = 3113
Global Const MQCACF_SYSP_OTMA_TPIPE_PFX = 3114
Global Const MQCACF_SYSP_ARCHIVE_PFX1 = 3115
Global Const MQCACF_SYSP_ARCHIVE_UNIT1 = 3116
Global Const MQCACF_SYSP_LOG_CORREL_ID = 3117
Global Const MQCACF_SYSP_UNIT_VOLSER = 3118
Global Const MQCACF_SYSP_Q_MGR_TIME = 3119
Global Const MQCACF_SYSP_Q_MGR_DATE = 3120
Global Const MQCACF_SYSP_Q_MGR_RBA = 3121
Global Const MQCACF_SYSP_LOG_RBA = 3122
Global Const MQCACF_SYSP_SERVICE = 3123
Global Const MQCACF_FROM_LISTENER_NAME = 3124
Global Const MQCACF_TO_LISTENER_NAME = 3125
Global Const MQCACF_FROM_SERVICE_NAME = 3126
Global Const MQCACF_TO_SERVICE_NAME = 3127
Global Const MQCACF_LAST_PUT_DATE = 3128
Global Const MQCACF_LAST_PUT_TIME = 3129
Global Const MQCACF_LAST_GET_DATE = 3130
Global Const MQCACF_LAST_GET_TIME = 3131
Global Const MQCACF_OPERATION_DATE = 3132
Global Const MQCACF_OPERATION_TIME = 3133
Global Const MQCACF_ACTIVITY_DESC = 3134
Global Const MQCACF_APPL_IDENTITY_DATA = 3135
Global Const MQCACF_APPL_ORIGIN_DATA = 3136
Global Const MQCACF_PUT_DATE = 3137
Global Const MQCACF_PUT_TIME = 3138
Global Const MQCACF_REPLY_TO_Q = 3139
Global Const MQCACF_REPLY_TO_Q_MGR = 3140
Global Const MQCACF_RESOLVED_Q_NAME = 3141
Global Const MQCACF_STRUC_ID = 3142
Global Const MQCACF_VALUE_NAME = 3143
Global Const MQCACF_SERVICE_START_DATE = 3144
Global Const MQCACF_SERVICE_START_TIME = 3145
Global Const MQCACF_SYSP_OFFLINE_RBA = 3146
Global Const MQCACF_SYSP_ARCHIVE_PFX2 = 3147
Global Const MQCACF_SYSP_ARCHIVE_UNIT2 = 3148
Global Const MQCACF_LAST_USED = 3148

'Character Channel Parameter Types'
Global Const MQCACH_FIRST = 3501
Global Const MQCACH_CHANNEL_NAME = 3501
Global Const MQCACH_DESC = 3502
Global Const MQCACH_MODE_NAME = 3503
Global Const MQCACH_TP_NAME = 3504
Global Const MQCACH_XMIT_Q_NAME = 3505
Global Const MQCACH_CONNECTION_NAME = 3506
Global Const MQCACH_MCA_NAME = 3507
Global Const MQCACH_SEC_EXIT_NAME = 3508
Global Const MQCACH_MSG_EXIT_NAME = 3509
Global Const MQCACH_SEND_EXIT_NAME = 3510
Global Const MQCACH_RCV_EXIT_NAME = 3511
Global Const MQCACH_CHANNEL_NAMES = 3512
Global Const MQCACH_SEC_EXIT_USER_DATA = 3513
Global Const MQCACH_MSG_EXIT_USER_DATA = 3514
Global Const MQCACH_SEND_EXIT_USER_DATA = 3515
Global Const MQCACH_RCV_EXIT_USER_DATA = 3516
Global Const MQCACH_USER_ID = 3517
Global Const MQCACH_PASSWORD = 3518
Global Const MQCACH_LOCAL_ADDRESS = 3520
Global Const MQCACH_LOCAL_NAME = 3521
Global Const MQCACH_LAST_MSG_TIME = 3524
Global Const MQCACH_LAST_MSG_DATE = 3525
Global Const MQCACH_MCA_USER_ID = 3527
Global Const MQCACH_CHANNEL_START_TIME = 3528
Global Const MQCACH_CHANNEL_START_DATE = 3529
Global Const MQCACH_MCA_JOB_NAME = 3530
Global Const MQCACH_LAST_LUWID = 3531
Global Const MQCACH_CURRENT_LUWID = 3532
Global Const MQCACH_FORMAT_NAME = 3533
Global Const MQCACH_MR_EXIT_NAME = 3534
Global Const MQCACH_MR_EXIT_USER_DATA = 3535
Global Const MQCACH_SSL_CIPHER_SPEC = 3544
Global Const MQCACH_SSL_PEER_NAME = 3545
Global Const MQCACH_SSL_HANDSHAKE_STAGE = 3546
Global Const MQCACH_SSL_SHORT_PEER_NAME = 3547
Global Const MQCACH_REMOTE_APPL_TAG = 3548
Global Const MQCACH_SSL_CERT_USER_ID = 3549
Global Const MQCACH_SSL_CERT_ISSUER_NAME = 3550
Global Const MQCACH_LU_NAME = 3551
Global Const MQCACH_IP_ADDRESS = 3552
Global Const MQCACH_TCP_NAME = 3553
Global Const MQCACH_LISTENER_NAME = 3554
Global Const MQCACH_LISTENER_DESC = 3555
Global Const MQCACH_LISTENER_START_DATE = 3556
Global Const MQCACH_LISTENER_START_TIME = 3557
Global Const MQCACH_SSL_KEY_RESET_DATE = 3558
Global Const MQCACH_SSL_KEY_RESET_TIME = 3559
Global Const MQCACH_LAST_USED = 3559

'****************************************************************'
'*  Values Related to Group Parameter Structures                *'
'****************************************************************'

'Group Parameter Types'
Global Const MQGACF_FIRST = 8001
Global Const MQGACF_COMMAND_CONTEXT = 8001
Global Const MQGACF_COMMAND_DATA = 8002
Global Const MQGACF_TRACE_ROUTE = 8003
Global Const MQGACF_OPERATION = 8004
Global Const MQGACF_ACTIVITY = 8005
Global Const MQGACF_EMBEDDED_MQMD = 8006
Global Const MQGACF_MESSAGE = 8007
Global Const MQGACF_MQMD = 8008
Global Const MQGACF_VALUE_NAMING = 8009
Global Const MQGACF_Q_ACCOUNTING_DATA = 8010
Global Const MQGACF_Q_STATISTICS_DATA = 8011
Global Const MQGACF_CHL_STATISTICS_DATA = 8012
Global Const MQGACF_LAST_USED = 8012

'****************************************************************'
'*  Parameter Values                                            *'
'****************************************************************'

'Action Options'
Global Const MQACT_FORCE_REMOVE = 1
Global Const MQACT_ADVANCE_LOG = 2
Global Const MQACT_COLLECT_STATISTICS = 3

'Authority Values'
Global Const MQAUTH_NONE = 0
Global Const MQAUTH_ALT_USER_AUTHORITY = 1
Global Const MQAUTH_BROWSE = 2
Global Const MQAUTH_CHANGE = 3
Global Const MQAUTH_CLEAR = 4
Global Const MQAUTH_CONNECT = 5
Global Const MQAUTH_CREATE = 6
Global Const MQAUTH_DELETE = 7
Global Const MQAUTH_DISPLAY = 8
Global Const MQAUTH_INPUT = 9
Global Const MQAUTH_INQUIRE = 10
Global Const MQAUTH_OUTPUT = 11
Global Const MQAUTH_PASS_ALL_CONTEXT = 12
Global Const MQAUTH_PASS_IDENTITY_CONTEXT = 13
Global Const MQAUTH_SET = 14
Global Const MQAUTH_SET_ALL_CONTEXT = 15
Global Const MQAUTH_SET_IDENTITY_CONTEXT = 16
Global Const MQAUTH_CONTROL = 17
Global Const MQAUTH_CONTROL_EXTENDED = 18

'Authority Options'
Global Const MQAUTHOPT_CUMULATIVE = &H100
Global Const MQAUTHOPT_ENTITY_EXPLICIT = &H1
Global Const MQAUTHOPT_ENTITY_SET = &H2
Global Const MQAUTHOPT_NAME_ALL_MATCHING = &H20
Global Const MQAUTHOPT_NAME_AS_WILDCARD = &H40
Global Const MQAUTHOPT_NAME_EXPLICIT = &H10

'Bridge Types'
Global Const MQBT_OTMA = 1

'Refresh Repository Options'
Global Const MQCFO_REFRESH_REPOSITORY_YES = 1
Global Const MQCFO_REFRESH_REPOSITORY_NO = 0

'Remove Queues Options'
Global Const MQCFO_REMOVE_QUEUES_YES = 1
Global Const MQCFO_REMOVE_QUEUES_NO = 0

'CF Status'
Global Const MQCFSTATUS_NOT_FOUND = 0
Global Const MQCFSTATUS_ACTIVE = 1
Global Const MQCFSTATUS_IN_RECOVER = 2
Global Const MQCFSTATUS_IN_BACKUP = 3
Global Const MQCFSTATUS_FAILED = 4
Global Const MQCFSTATUS_NONE = 5
Global Const MQCFSTATUS_UNKNOWN = 6
Global Const MQCFSTATUS_ADMIN_INCOMPLETE = 20
Global Const MQCFSTATUS_NEVER_USED = 21
Global Const MQCFSTATUS_NO_BACKUP = 22
Global Const MQCFSTATUS_NOT_FAILED = 23
Global Const MQCFSTATUS_NOT_RECOVERABLE = 24
Global Const MQCFSTATUS_XES_ERROR = 25

'CF Types'
Global Const MQCFTYPE_APPL = 0
Global Const MQCFTYPE_ADMIN = 1

'Indoubt Status'
Global Const MQCHIDS_NOT_INDOUBT = 0
Global Const MQCHIDS_INDOUBT = 1

'Channel Dispositions'
Global Const MQCHLD_ALL = -1
Global Const MQCHLD_PRIVATE = 4
Global Const MQCHLD_SHARED = 2
Global Const MQCHLD_FIXSHARED = 5

'Channel Status'
Global Const MQCHS_INACTIVE = 0
Global Const MQCHS_BINDING = 1
Global Const MQCHS_STARTING = 2
Global Const MQCHS_RUNNING = 3
Global Const MQCHS_STOPPING = 4
Global Const MQCHS_RETRYING = 5
Global Const MQCHS_STOPPED = 6
Global Const MQCHS_REQUESTING = 7
Global Const MQCHS_PAUSED = 8
Global Const MQCHS_INITIALIZING = 13

'Channel Substates'
Global Const MQCHSSTATE_OTHER = 0
Global Const MQCHSSTATE_END_OF_BATCH = 100
Global Const MQCHSSTATE_SENDING = 200
Global Const MQCHSSTATE_RECEIVING = 300
Global Const MQCHSSTATE_SERIALIZING = 400
Global Const MQCHSSTATE_RESYNCHING = 500
Global Const MQCHSSTATE_HEARTBEATING = 600
Global Const MQCHSSTATE_IN_SCYEXIT = 700
Global Const MQCHSSTATE_IN_RCVEXIT = 800
Global Const MQCHSSTATE_IN_SENDEXIT = 900
Global Const MQCHSSTATE_IN_MSGEXIT = 1000
Global Const MQCHSSTATE_IN_MREXIT = 1100
Global Const MQCHSSTATE_IN_CHADEXIT = 1200
Global Const MQCHSSTATE_NET_CONNECTING = 1250
Global Const MQCHSSTATE_SSL_HANDSHAKING = 1300
Global Const MQCHSSTATE_NAME_SERVER = 1400
Global Const MQCHSSTATE_IN_MQPUT = 1500
Global Const MQCHSSTATE_IN_MQGET = 1600
Global Const MQCHSSTATE_IN_MQI_CALL = 1700
Global Const MQCHSSTATE_COMPRESSING = 1800

'Channel Shared Restart Options'
Global Const MQCHSH_RESTART_NO = 0
Global Const MQCHSH_RESTART_YES = 1

'Channel Stop Options'
Global Const MQCHSR_STOP_NOT_REQUESTED = 0
Global Const MQCHSR_STOP_REQUESTED = 1

'Channel Table Types'
Global Const MQCHTAB_Q_MGR = 1
Global Const MQCHTAB_CLNTCONN = 2

'Command Information Values'
Global Const MQCMDI_CMDSCOPE_ACCEPTED = 1
Global Const MQCMDI_CMDSCOPE_GENERATED = 2
Global Const MQCMDI_CMDSCOPE_COMPLETED = 3
Global Const MQCMDI_QSG_DISP_COMPLETED = 4
Global Const MQCMDI_COMMAND_ACCEPTED = 5
Global Const MQCMDI_CLUSTER_REQUEST_QUEUED = 6
Global Const MQCMDI_CHANNEL_INIT_STARTED = 7
Global Const MQCMDI_RECOVER_STARTED = 11
Global Const MQCMDI_BACKUP_STARTED = 12
Global Const MQCMDI_RECOVER_COMPLETED = 13
Global Const MQCMDI_SEC_TIMER_ZERO = 14
Global Const MQCMDI_REFRESH_CONFIGURATION = 16
Global Const MQCMDI_SEC_SIGNOFF_ERROR = 17
Global Const MQCMDI_IMS_BRIDGE_SUSPENDED = 18
Global Const MQCMDI_DB2_SUSPENDED = 19
Global Const MQCMDI_DB2_OBSOLETE_MSGS = 20

'Disconnect Types'
Global Const MQDISCONNECT_NORMAL = 0
Global Const MQDISCONNECT_IMPLICIT = 1
Global Const MQDISCONNECT_Q_MGR = 2

'Escape Types'
Global Const MQET_MQSC = 1

'Event Origins'
Global Const MQEVO_OTHER = 0
Global Const MQEVO_CONSOLE = 1
Global Const MQEVO_INIT = 2
Global Const MQEVO_MSG = 3
Global Const MQEVO_MQSET = 4
Global Const MQEVO_INTERNAL = 5

'Event Recording'
Global Const MQEVR_DISABLED = 0
Global Const MQEVR_ENABLED = 1
Global Const MQEVR_EXCEPTION = 2
Global Const MQEVR_NO_DISPLAY = 3

'Force Options'
Global Const MQFC_YES = 1
Global Const MQFC_NO = 0

'Handle States'
Global Const MQHSTATE_INACTIVE = 0
Global Const MQHSTATE_ACTIVE = 1

'Inbound Dispositions'
Global Const MQINBD_Q_MGR = 0
Global Const MQINBD_GROUP = 3

'Indoubt Options'
Global Const MQIDO_COMMIT = 1
Global Const MQIDO_BACKOUT = 2

'Message Channel Agent Status'
Global Const MQMCAS_STOPPED = 0
Global Const MQMCAS_RUNNING = 3

'Mode Options'
Global Const MQMODE_FORCE = 0
Global Const MQMODE_QUIESCE = 1
Global Const MQMODE_TERMINATE = 2

'Purge Options'
Global Const MQPO_YES = 1
Global Const MQPO_NO = 0

'Queue Manager Definition Types'
Global Const MQQMDT_EXPLICIT_CLUSTER_SENDER = 1
Global Const MQQMDT_AUTO_CLUSTER_SENDER = 2
Global Const MQQMDT_AUTO_EXP_CLUSTER_SENDER = 4
Global Const MQQMDT_CLUSTER_RECEIVER = 3

'Queue Manager Facility'
Global Const MQQMFAC_IMS_BRIDGE = 1
Global Const MQQMFAC_DB2 = 2

'Queue Manager Status'
Global Const MQQMSTA_STARTING = 1
Global Const MQQMSTA_RUNNING = 2
Global Const MQQMSTA_QUIESCING = 3

'Queue Manager Types'
Global Const MQQMT_NORMAL = 0
Global Const MQQMT_REPOSITORY = 1

'Quiesce Options'
Global Const MQQO_YES = 1
Global Const MQQO_NO = 0

'Queue Service-Interval Events'
Global Const MQQSIE_NONE = 0
Global Const MQQSIE_HIGH = 1
Global Const MQQSIE_OK = 2

'Queue Status Open Types'
Global Const MQQSOT_ALL = 1
Global Const MQQSOT_INPUT = 2
Global Const MQQSOT_OUTPUT = 3

'QSG Status'
Global Const MQQSGS_UNKNOWN = 0
Global Const MQQSGS_CREATED = 1
Global Const MQQSGS_ACTIVE = 2
Global Const MQQSGS_INACTIVE = 3
Global Const MQQSGS_FAILED = 4
Global Const MQQSGS_PENDING = 5

'Queue Status Open Options for SET, BROWSE, INPUT'
Global Const MQQSO_NO = 0
Global Const MQQSO_YES = 1
Global Const MQQSO_SHARED = 1
Global Const MQQSO_EXCLUSIVE = 2

'Queue Status Uncommitted Messages'
Global Const MQQSUM_YES = 1
Global Const MQQSUM_NO = 0

'Replace Options'
Global Const MQRP_YES = 1
Global Const MQRP_NO = 0

'Reason Qualifiers'
Global Const MQRQ_CONN_NOT_AUTHORIZED = 1
Global Const MQRQ_OPEN_NOT_AUTHORIZED = 2
Global Const MQRQ_CLOSE_NOT_AUTHORIZED = 3
Global Const MQRQ_CMD_NOT_AUTHORIZED = 4
Global Const MQRQ_Q_MGR_STOPPING = 5
Global Const MQRQ_Q_MGR_QUIESCING = 6
Global Const MQRQ_CHANNEL_STOPPED_OK = 7
Global Const MQRQ_CHANNEL_STOPPED_ERROR = 8
Global Const MQRQ_CHANNEL_STOPPED_RETRY = 9
Global Const MQRQ_CHANNEL_STOPPED_DISABLED = 10
Global Const MQRQ_BRIDGE_STOPPED_OK = 11
Global Const MQRQ_BRIDGE_STOPPED_ERROR = 12
Global Const MQRQ_SSL_HANDSHAKE_ERROR = 13
Global Const MQRQ_SSL_CIPHER_SPEC_ERROR = 14
Global Const MQRQ_SSL_CLIENT_AUTH_ERROR = 15
Global Const MQRQ_SSL_PEER_NAME_ERROR = 16

'Refresh Types'
Global Const MQRT_CONFIGURATION = 1
Global Const MQRT_EXPIRY = 2
Global Const MQRT_NSPROC = 3

'Queue Definition Scope'
Global Const MQSCO_Q_MGR = 1
Global Const MQSCO_CELL = 2

'Security Items'
Global Const MQSECITEM_ALL = 0
Global Const MQSECITEM_MQADMIN = 1
Global Const MQSECITEM_MQNLIST = 2
Global Const MQSECITEM_MQPROC = 3
Global Const MQSECITEM_MQQUEUE = 4
Global Const MQSECITEM_MQCONN = 5
Global Const MQSECITEM_MQCMDS = 6

'Security Switches'
Global Const MQSECSW_PROCESS = 1
Global Const MQSECSW_NAMELIST = 2
Global Const MQSECSW_Q = 3
Global Const MQSECSW_CONTEXT = 6
Global Const MQSECSW_ALTERNATE_USER = 7
Global Const MQSECSW_COMMAND = 8
Global Const MQSECSW_CONNECTION = 9
Global Const MQSECSW_SUBSYSTEM = 10
Global Const MQSECSW_COMMAND_RESOURCES = 11
Global Const MQSECSW_Q_MGR = 15
Global Const MQSECSW_QSG = 16

'Security Switch States'
Global Const MQSECSW_OFF_FOUND = 21
Global Const MQSECSW_ON_FOUND = 22
Global Const MQSECSW_OFF_NOT_FOUND = 23
Global Const MQSECSW_ON_NOT_FOUND = 24
Global Const MQSECSW_OFF_ERROR = 25
Global Const MQSECSW_ON_OVERRIDDEN = 26

'Security Types'
Global Const MQSECTYPE_AUTHSERV = 1
Global Const MQSECTYPE_SSL = 2
Global Const MQSECTYPE_CLASSES = 3

'Suspend Status'
Global Const MQSUS_YES = 1
Global Const MQSUS_NO = 0

'System Parameter Values'
Global Const MQSYSP_NO = 0
Global Const MQSYSP_YES = 1
Global Const MQSYSP_EXTENDED = 2
Global Const MQSYSP_TYPE_INITIAL = 10
Global Const MQSYSP_TYPE_SET = 11
Global Const MQSYSP_TYPE_LOG_COPY = 12
Global Const MQSYSP_TYPE_LOG_STATUS = 13
Global Const MQSYSP_TYPE_ARCHIVE_TAPE = 14
Global Const MQSYSP_ALLOC_BLK = 20
Global Const MQSYSP_ALLOC_TRK = 21
Global Const MQSYSP_ALLOC_CYL = 22
Global Const MQSYSP_STATUS_BUSY = 30
Global Const MQSYSP_STATUS_PREMOUNT = 31
Global Const MQSYSP_STATUS_AVAILABLE = 32
Global Const MQSYSP_STATUS_UNKNOWN = 33
Global Const MQSYSP_STATUS_ALLOC_ARCHIVE = 34
Global Const MQSYSP_STATUS_COPYING_BSDS = 35
Global Const MQSYSP_STATUS_COPYING_LOG = 36

'Time units'
Global Const MQTIME_UNIT_MINS = 0
Global Const MQTIME_UNIT_SECS = 1

'User ID Support'
Global Const MQUIDSUPP_NO = 0
Global Const MQUIDSUPP_YES = 1

'UOW States'
Global Const MQUOWST_NONE = 0
Global Const MQUOWST_ACTIVE = 1
Global Const MQUOWST_PREPARED = 2
Global Const MQUOWST_UNRESOLVED = 3

'UOW Types'
Global Const MQUOWT_Q_MGR = 0
Global Const MQUOWT_CICS = 1
Global Const MQUOWT_RRS = 2
Global Const MQUOWT_IMS = 3
Global Const MQUOWT_XA = 4

'Page Set Usage Values'
Global Const MQUSAGE_PS_AVAILABLE = 0
Global Const MQUSAGE_PS_DEFINED = 1
Global Const MQUSAGE_PS_OFFLINE = 2
Global Const MQUSAGE_PS_NOT_DEFINED = 3
Global Const MQUSAGE_EXPAND_USER = 1
Global Const MQUSAGE_EXPAND_SYSTEM = 2
Global Const MQUSAGE_EXPAND_NONE = 3

'Data Set Usage Values'
Global Const MQUSAGE_DS_OLDEST_ACTIVE_UOW = 10
Global Const MQUSAGE_DS_OLDEST_PS_RECOVERY = 11
Global Const MQUSAGE_DS_OLDEST_CF_RECOVERY = 12

'****************************************************************'
'*  Values Related to Route Tracing and Activity Operations     *'
'****************************************************************'

'Activity Operations'
Global Const MQOPER_SYSTEM_FIRST = 0
Global Const MQOPER_UNKNOWN = 0
Global Const MQOPER_BROWSE = 1
Global Const MQOPER_DISCARD = 2
Global Const MQOPER_GET = 3
Global Const MQOPER_PUT = 4
Global Const MQOPER_PUT_REPLY = 5
Global Const MQOPER_PUT_REPORT = 6
Global Const MQOPER_RECEIVE = 7
Global Const MQOPER_SEND = 8
Global Const MQOPER_TRANSFORM = 9
Global Const MQOPER_SYSTEM_LAST = 65535
Global Const MQOPER_APPL_FIRST = 65536
Global Const MQOPER_APPL_LAST = 999999999

'Route Tracing Max Activities (MQIACF_MAX_ACTIVITIES)'
Global Const MQROUTE_UNLIMITED_ACTIVITIES = 0

'Route Tracing Detail (MQIACF_ROUTE_DETAIL)'
Global Const MQROUTE_DETAIL_LOW = &H2
Global Const MQROUTE_DETAIL_MEDIUM = &H8
Global Const MQROUTE_DETAIL_HIGH = &H20

'Route Tracing Forwarding (MQIACF_ROUTE_FORWARDING)'
Global Const MQROUTE_FORWARD_ALL = &H100
Global Const MQROUTE_FORWARD_IF_SUPPORTED = &H200
Global Const MQROUTE_FORWARD_REJ_UNSUP_MASK = &HFFFF0000

'Route Tracing Delivery (MQIACF_ROUTE_DELIVERY)'
Global Const MQROUTE_DELIVER_YES = &H1000
Global Const MQROUTE_DELIVER_NO = &H2000
Global Const MQROUTE_DELIVER_REJ_UNSUP_MASK = &HFFFF0000

'Route Tracing Accumulation (MQIACF_ROUTE_ACCUMULATION)'
Global Const MQROUTE_ACCUMULATE_NONE = &H10003
Global Const MQROUTE_ACCUMULATE_IN_MSG = &H10004
Global Const MQROUTE_ACCUMULATE_AND_REPLY = &H10005

'****************************************************************'
'*  Values Related to Publish/Subscribe                         *'
'****************************************************************'

'Delete Options'
Global Const MQDELO_NONE = &H0
Global Const MQDELO_LOCAL = &H4

'Publication Options'
Global Const MQPUBO_NONE = &H0
Global Const MQPUBO_CORREL_ID_AS_IDENTITY = &H1
Global Const MQPUBO_RETAIN_PUBLICATION = &H2
Global Const MQPUBO_OTHER_SUBSCRIBERS_ONLY = &H4
Global Const MQPUBO_NO_REGISTRATION = &H8
Global Const MQPUBO_IS_RETAINED_PUBLICATION = &H10

'Registration Options'
Global Const MQREGO_NONE = &H0
Global Const MQREGO_CORREL_ID_AS_IDENTITY = &H1
Global Const MQREGO_ANONYMOUS = &H2
Global Const MQREGO_LOCAL = &H4
Global Const MQREGO_DIRECT_REQUESTS = &H8
Global Const MQREGO_NEW_PUBLICATIONS_ONLY = &H10
Global Const MQREGO_PUBLISH_ON_REQUEST_ONLY = &H20
Global Const MQREGO_DEREGISTER_ALL = &H40
Global Const MQREGO_INCLUDE_STREAM_NAME = &H80
Global Const MQREGO_INFORM_IF_RETAINED = &H100
Global Const MQREGO_DUPLICATES_OK = &H200
Global Const MQREGO_NON_PERSISTENT = &H400
Global Const MQREGO_PERSISTENT = &H800
Global Const MQREGO_PERSISTENT_AS_PUBLISH = &H1000
Global Const MQREGO_PERSISTENT_AS_Q = &H2000
Global Const MQREGO_ADD_NAME = &H4000
Global Const MQREGO_NO_ALTERATION = &H8000
Global Const MQREGO_FULL_RESPONSE = &H10000
Global Const MQREGO_JOIN_SHARED = &H20000
Global Const MQREGO_JOIN_EXCLUSIVE = &H40000
Global Const MQREGO_LEAVE_ONLY = &H80000
Global Const MQREGO_VARIABLE_USER_ID = &H100000
Global Const MQREGO_LOCKED = &H200000

'User Attribute Selectors'
Global Const MQUA_FIRST = 65536
Global Const MQUA_LAST = 999999999


'****************************************************************'
'*  MQCFH Structure -- PCF Header                               *'
'****************************************************************'

Type MQCFH
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Version As Long 'Structure version number'
  Command As Long 'Command identifier'
  MsgSeqNumber As Long 'Message sequence number'
  Control As Long 'Control options'
  CompCode As Long 'Completion code'
  Reason As Long 'Reason code qualifying completion code'
  ParameterCount As Long 'Count of parameter structures'
End Type

'Default Instance of MQCFH Structure'
Global MQCFH_DEFAULT As MQCFH


'****************************************************************'
'*  MQCFBF Structure -- PCF Byte String Filter Parameter        *'
'****************************************************************'

Type MQCFBF
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  Operator As Long 'Operator identifier'
  FilterValueLength As Long 'Filter value length'
End Type

'Default Instance of MQCFBF Structure'
Global MQCFBF_DEFAULT As MQCFBF


'****************************************************************'
'*  MQCFBS Structure -- PCF Byte String Parameter               *'
'****************************************************************'

Type MQCFBS
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  StringLength As Long 'Length of string'
End Type

'Default Instance of MQCFBS Structure'
Global MQCFBS_DEFAULT As MQCFBS


'****************************************************************'
'*  MQCFGR Structure -- PCF Group Parameter                     *'
'****************************************************************'

Type MQCFGR
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  ParameterCount As Long 'Count of group parameter structures'
End Type

'Default Instance of MQCFGR Structure'
Global MQCFGR_DEFAULT As MQCFGR


'****************************************************************'
'*  MQCFIF Structure -- PCF Integer Filter Parameter            *'
'****************************************************************'

Type MQCFIF
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  Operator As Long 'Operator identifier'
  FilterValue As Long 'Filter value'
End Type

'Default Instance of MQCFIF Structure'
Global MQCFIF_DEFAULT As MQCFIF


'****************************************************************'
'*  MQCFIL Structure -- PCF Integer-List Parameter              *'
'****************************************************************'

Type MQCFIL
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  Count As Long 'Count of parameter values'
End Type

'Default Instance of MQCFIL Structure'
Global MQCFIL_DEFAULT As MQCFIL


'****************************************************************'
'*  MQCFIN Structure -- PCF Integer Parameter                   *'
'****************************************************************'

Type MQCFIN
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  Value As Long 'Parameter value'
End Type

'Default Instance of MQCFIN Structure'
Global MQCFIN_DEFAULT As MQCFIN


'****************************************************************'
'*  MQCFSF Structure -- PCF String Filter Parameter             *'
'****************************************************************'

Type MQCFSF
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  Operator As Long 'Operator identifier'
  CodedCharSetId As Long 'Coded character set identifier'
  FilterValueLength As Long 'Filter value length'
End Type

'Default Instance of MQCFSF Structure'
Global MQCFSF_DEFAULT As MQCFSF


'****************************************************************'
'*  MQCFSL Structure -- PCF String-List Parameter               *'
'****************************************************************'

Type MQCFSL
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  CodedCharSetId As Long 'Coded character set identifier'
  Count As Long 'Count of parameter values'
  StringLength As Long 'Length of one string'
End Type

'Default Instance of MQCFSL Structure'
Global MQCFSL_DEFAULT As MQCFSL


'****************************************************************'
'*  MQCFST Structure -- PCF String Parameter                    *'
'****************************************************************'

Type MQCFST
  Type As Long 'Structure type'
  StrucLength As Long 'Structure length'
  Parameter As Long 'Parameter identifier'
  CodedCharSetId As Long 'Coded character set identifier'
  StringLength As Long 'Length of string'
End Type

'Default Instance of MQCFST Structure'
Global MQCFST_DEFAULT As MQCFST


'****************************************************************'
'*  MQEPH Structure -- Embedded PCF header                      *'
'****************************************************************'

Type MQEPH
  StrucId As String * 4 'Structure identifier'
  Version As Long 'Structure version number'
  StrucLength As Long 'Total length of MQEPH including MQCFH and parameter structures that follow'
  Encoding As Long 'Numeric encoding of data that follows last PCF parameter structure'
  CodedCharSetId As Long 'Character set identifier of data that follows last PCF parameter structure'
  Format As String * 8 'Format name of data that follows last PCF parameter structure'
  Flags As Long 'Flags'
  PCFHeader As MQCFH 'Programmable Command Format Header'
End Type

'Default Instance of MQEPH Structure'
Global MQEPH_DEFAULT As MQEPH


'*********************************************************************'
'*  MQ_SETDEFAULTS_CF Subroutine -- Set Defaults                     *'
'*********************************************************************'

'****************************************************************'
'*  End of CMQCFB                                               *'
'****************************************************************'

Sub MQCFH_DEFAULTS(Struc As MQCFH)
  Struc.Type = MQCFT_COMMAND
  Struc.StrucLength = MQCFH_STRUC_LENGTH
  Struc.Version = MQCFH_VERSION_1
  Struc.Command = MQCMD_NONE
  Struc.MsgSeqNumber = 1
  Struc.Control = MQCFC_LAST
  Struc.CompCode = MQCC_OK
  Struc.Reason = MQRC_NONE
  Struc.ParameterCount = 0
End Sub

Sub MQCFBF_DEFAULTS(Struc As MQCFBF)
  Struc.Type = MQCFT_BYTE_STRING_FILTER
  Struc.StrucLength = MQCFBF_STRUC_LENGTH_FIXED
  Struc.Parameter = 0
  Struc.Operator = 0
  Struc.FilterValueLength = 0
End Sub

Sub MQCFBS_DEFAULTS(Struc As MQCFBS)
  Struc.Type = MQCFT_BYTE_STRING
  Struc.StrucLength = MQCFBS_STRUC_LENGTH_FIXED
  Struc.Parameter = 0
  Struc.StringLength = 0
End Sub

Sub MQCFGR_DEFAULTS(Struc As MQCFGR)
  Struc.Type = MQCFT_GROUP
  Struc.StrucLength = MQCFGR_STRUC_LENGTH
  Struc.Parameter = 0
  Struc.ParameterCount = 0
End Sub

Sub MQCFIF_DEFAULTS(Struc As MQCFIF)
  Struc.Type = MQCFT_INTEGER_FILTER
  Struc.StrucLength = MQCFIF_STRUC_LENGTH
  Struc.Parameter = 0
  Struc.Operator = 0
  Struc.FilterValue = 0
End Sub

Sub MQCFIL_DEFAULTS(Struc As MQCFIL)
  Struc.Type = MQCFT_INTEGER_LIST
  Struc.StrucLength = MQCFIL_STRUC_LENGTH_FIXED
  Struc.Parameter = 0
  Struc.Count = 0
End Sub

Sub MQCFIN_DEFAULTS(Struc As MQCFIN)
  Struc.Type = MQCFT_INTEGER
  Struc.StrucLength = MQCFIN_STRUC_LENGTH
  Struc.Parameter = 0
  Struc.Value = 0
End Sub

Sub MQCFSF_DEFAULTS(Struc As MQCFSF)
  Struc.Type = MQCFT_STRING_FILTER
  Struc.StrucLength = MQCFSF_STRUC_LENGTH_FIXED
  Struc.Parameter = 0
  Struc.Operator = 0
  Struc.CodedCharSetId = MQCCSI_DEFAULT
  Struc.FilterValueLength = 0
End Sub

Sub MQCFSL_DEFAULTS(Struc As MQCFSL)
  Struc.Type = MQCFT_STRING_LIST
  Struc.StrucLength = MQCFSL_STRUC_LENGTH_FIXED
  Struc.Parameter = 0
  Struc.CodedCharSetId = MQCCSI_DEFAULT
  Struc.Count = 0
  Struc.StringLength = 0
End Sub

Sub MQCFST_DEFAULTS(Struc As MQCFST)
  Struc.Type = MQCFT_STRING
  Struc.StrucLength = MQCFST_STRUC_LENGTH_FIXED
  Struc.Parameter = 0
  Struc.CodedCharSetId = MQCCSI_DEFAULT
  Struc.StringLength = 0
End Sub

Sub MQEPH_DEFAULTS(Struc As MQEPH)
  Struc.StrucId = MQEPH_STRUC_ID
  Struc.Version = MQEPH_VERSION_1
  Struc.StrucLength = MQEPH_STRUC_LENGTH_FIXED
  Struc.Encoding = 0
  Struc.CodedCharSetId = MQCCSI_UNDEFINED
  Struc.Format = MQFMT_NONE
  Struc.Flags = MQEPH_NONE
  Dim TempPCFHeader As MQCFH
  MQCFH_DEFAULTS TempPCFHeader
  Struc.PCFHeader = TempPCFHeader
End Sub

Sub MQ_SETDEFAULTS_CF()

  'Set default structures'
  MQCFH_DEFAULTS MQCFH_DEFAULT
  MQCFBF_DEFAULTS MQCFBF_DEFAULT
  MQCFBS_DEFAULTS MQCFBS_DEFAULT
  MQCFGR_DEFAULTS MQCFGR_DEFAULT
  MQCFIF_DEFAULTS MQCFIF_DEFAULT
  MQCFIL_DEFAULTS MQCFIL_DEFAULT
  MQCFIN_DEFAULTS MQCFIN_DEFAULT
  MQCFSF_DEFAULTS MQCFSF_DEFAULT
  MQCFSL_DEFAULTS MQCFSL_DEFAULT
  MQCFST_DEFAULTS MQCFST_DEFAULT
  MQEPH_DEFAULTS MQEPH_DEFAULT

End Sub

'****************************************************************'
'*  End of CMQCFB                                               *'
'****************************************************************'

