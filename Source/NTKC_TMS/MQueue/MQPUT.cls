VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQPUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim gHcon As Long
Dim gHobj As Long


Public Function close_Conn() As Boolean
  Dim C_options As Long     ' MQCLOSE options
  Dim CompCode As Long      ' Completion code
  Dim Reason As Long        ' Reason code

  ' If queue was opened successfully - then close it
  If gHobj Then
    C_options = 0
    MQCLOSE gHcon, gHobj, C_options, CompCode, Reason

    If CompCode <> MQCC_OK Then
      Beep
      close_Conn = False
    Else
      close_Conn = True
    End If
    Else
    close_Conn = False
  End If
  
  ' If queue manager connected successfully - then disconnect
  If gHcon Then
    MQDISC gHcon, CompCode, Reason
    'close_Conn = False
  End If
  
  SaveErrorLog "MQPUT", "close_ConnMQ", C_options, "CompCode = " & CompCode & " Reason = " & Reason
End Function

Public Function open_Conn(ByVal qm As String, ByVal qn As String) As Boolean
  Dim CompCode As Long      ' Completion code
  Dim Reason As Long        ' Reason code
  Dim O_options As Long     ' MQOPEN options
  Dim od As MQOD            ' Object descriptor

  ' Setup constants
  MQ_SETDEFAULTS
  
  If gHcon = 0 Then
    ' Connect to the default queue manager: qm
    ' gHcon <> -1 okie
    MQCONN qm, gHcon, CompCode, Reason
  
    ' CompCode = MQCC_WARNING, then queue manager was already connected
    ' CompCode = MQCC_OK, then queue manager connected OK
    ' otherwise it's an error
    If CompCode = MQCC_FAILED Then
      Beep
      open_Conn = False
      'Exit Function
    End If
  End If

  ' Open the queue specified in the Queue textbox
  ' If the connect was successful
  If gHcon Then
    MQOD_DEFAULTS od
    od.ObjectName = qn 'Queue name
    
    ' We use the default Queue Manager
    '
    ' If it is a client, then this will be the
    ' Server's default queue manager
    '
    ' To specify a queue manager uncomment the
    ' Following line and enter a queue manager
    ' name (which is case sensitive)
    '
    ' od.ObjectQMgrName = "MyQueueManager"
    
    O_options = MQOO_OUTPUT
    MQOPEN gHcon, od, O_options, gHobj, CompCode, Reason
    ' Reason=0 -- khong loi
    ' CompCode <>2 --loi
    
    If CompCode <> MQCC_OK Then
      Beep
      open_Conn = False
      'Exit Function
      'COMMON PROBLEMS:
      '================================================
      '  CompCode = 2, Reason = 2085
      '================================================
      '
      'PROBLEM:
      '  MQRC_UNKNOWN_OBJECT_NQME
      '  Queue name is incorrect
      '
      'SOLUTION:
      '  Check spelling and case of queuename
      '
      
''      If Reason = 2085 Then
''        MsgBox "HINT: Check the spelling of the Queue name, which is Case-Sensitive.", vbOKOnly, "ReasonCode = (2085) - MQRC_UNKNOWN_OBJECT_NAME"
''      End If
           
      '================================================
      '  CompCode = 2, Reason = 2087
      '================================================
      '
      'PROBLEM:
      '  MQRC_UNKNOWN_REMOTE_Q_MGR
      '  Queue manager name is incorrect
      '
      'SOLUTION:
      '  Check spelling and case of QueueManager
      '  name - not applcable if default is being
      '  used
''
''      If Reason = 2087 Then
''        MsgBox "HINT: Check the spelling of the Queue Manager name, which is Case-Sensitive.", vbOKOnly, "ReasonCode = (2087) - MQRC_REMOTE_Q_MGR"
''      End If
    Else
      open_Conn = True
    End If
    Else
    open_Conn = False
  End If
  
  SaveErrorLog "MQPUT", "open_ConnMQ", O_options, "CompCode = " & CompCode & " Reason = " & Reason
  
End Function

Public Function put_Msg(ByVal xml As String) As Boolean
    Dim md       As MQMD            ' Message descriptor
    Dim pmo      As MQPMO          ' MQPUT message options
    Dim buflen   As Long        ' Length of message to be put
    Dim Buffer   As String      ' Message to be put
    Dim CompCode As Long      ' Completion code
    Dim Reason   As Long        ' Reason code

    ' Setup MQMD and MQPMO to their respective
    ' initial values
    MQMD_DEFAULTS md
    md.Format = MQFMT_STRING
    ' We must set the code page in the MD to the code page of our characters
    ' Assume Windows Latin 1 - this should be changed for other scripts
    'md.CodedCharSetId = 1252
    md.CodedCharSetId = 1208
    MQPMO_DEFAULTS pmo

    ' Get message to be "put" from the Put text box
    buflen = Len(xml)
    Buffer = xml
    MQPUT gHcon, gHobj, md, pmo, buflen, Buffer, CompCode, Reason

    If CompCode <> MQCC_OK Then
        put_Msg = False
    Else
        put_Msg = True
    End If
  
    SaveErrorLog "MQPUT", "close_ConnMQ", buflen, "CompCode = " & CompCode & " Reason = " & Reason
End Function

