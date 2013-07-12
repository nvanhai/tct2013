VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessRunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
    

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32" _
    Alias "Process32First" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32" _
    Alias "Process32Next" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Sub CloseHandle Lib "kernel32" _
   (ByVal hPass As Long)
   
'**************************************************
'Description:ProcessRunning function call API functions and list processes
'            are running, then check whether a process name is  in this list.
'Author:ThanhDX.
'Date:20/01/2006.
'Input:
'       strProcessName: process name is checked.
'Return:
'       True: process (by process name) is running.
'       False: otherwise.
'**************************************************
 
Public Function ProcessRunning(ByVal strProcessName As String) As Boolean

    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim success As Long
    Dim blnReturn As Boolean
        
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

    If hSnapShot = -1 Then Exit Function

    uProcess.dwSize = Len(uProcess)
    success = ProcessFirst(hSnapShot, uProcess)

    If success = 1 Then
        Do
            If InStr(1, uProcess.szExeFile, strProcessName) <> 0 Then
                blnReturn = True
                Exit Do
            End If
        Loop While ProcessNext(hSnapShot, uProcess)
    End If

    Call CloseHandle(hSnapShot)
    ProcessRunning = blnReturn

End Function

