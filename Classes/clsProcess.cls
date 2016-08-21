VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NMC! V1.5    Process Management Member Class                                     *
'*                                                                                     *
'*  Created:     April 14, 2006                                                        *
'*  Updated:     April 29, 2006                                                        *
'*  Purpose:     NT System Security Master Class                                       *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.5.0                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  Throughout Project                                                    *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

'/~ Process Management ~/
'/~ Process_Enumerate       - list running processes    [out -col]
'/~ Process_Exists          - test process existence    [in -string | out -bool]
'/~ Process_Terminate       - terminate a process       [in -string | out -bool]
'/~ Process_GetClass        - get process class         [in -long | out -long]
'/~ Process_SetClass        - change process priority   [in - string + enum | out -bool]
'/~ Process_EnumG2          - server compliant enum     [out -col]
'/~ Thread_Enumerate        - list a process threads    [in -string | out -bool]
'/~ Thread_Suspend          - suspend a thread          [in -long + string | out -bool]
'/~ Thread_Resume           - resume a thread           [in -long + string | out -bool]
'/~ Thread_GetPriority      - get thread priority       [in -long | out -long]
'/~ Thread_SetPriority      - set thread priority       [in -long + enum | -out bool]
'/~ Thread_Terminate        - kill a thread             [in -long | out -bool]
'/~ Module_Enumerate        - list a process modules    [in -string | out -col]
'/~ Module_EnumG2           - server compliant enum     [in -string | out -col]


'>  Start Process Dependencies
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Const PROCESS_VM_READ                           As Integer = 16
Private Const PROCESS_SET_INFORMATION                   As Long = &H200
Private Const PROCESS_QUERY_INFORMATION                 As Long = &H400
Private Const MAX_PATH                                  As Long = 260
Private Const TH32CS_SNAPTHREAD                         As Long = &H4
Private Const TH32CS_SNAPHEAPLIST                       As Long = &H1
Private Const TH32CS_SNAPPROCESS                        As Long = &H2
Private Const TH32CS_SNAPMODULE                         As Long = &H8
Private Const PROCESS_CTERMINATE                        As Long = &H1
Private Const THREAD_SUSPEND_RESUME                     As Long = &H2
Private Const THREAD_SET_INFORMATION                    As Long = &H20
Private Const THREAD_QUERY_INFORMATION                  As Long = &H40
Private Const TH32CS_SNAPALL                            As Double = _
        (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or _
        TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Private Type PROCESSENTRY32
    dwSize                                              As Long
    cntUsage                                            As Long
    th32ProcessID                                       As Long
    th32DefaultHeapID                                   As Long
    th32ModuleID                                        As Long
    cntThreads                                          As Long
    th32ParentProcessID                                 As Long
    pcPriClassBase                                      As Long
    dwFlags                                             As Long
    szexeFile                                           As String * MAX_PATH
End Type

Private Type MODULEENTRY32
    dwSize                                              As Long
    th32ModuleID                                        As Long
    th32ProcessID                                       As Long
    GlblcntUsage                                        As Long
    ProccntUsage                                        As Long
    modBaseAddr                                         As Long
    modBaseSize                                         As Long
    hModule                                             As Long
    szModule                                            As String * 256
    szExePath                                           As String * MAX_PATH
End Type

Private Type THREADENTRY32
    dwSize                                              As Long
    cntUsage                                            As Long
    th32ThreadID                                        As Long
    th32OwnerProcessID                                  As Long
    tpBasePri                                           As Long
    tpDeltaPri                                          As Long
    dwFlags                                             As Long
End Type

Public Enum eThreadPriority
    Thread_Idle = -15
    Thread_LowRT = 15
    Thread_Minimum = -2
    Thread_Normal = 0
    Thread_Maximum = 2
End Enum
    
Public Enum eProcessClass
    Process_Idle = &H40
    Process_Normal = &H20
    Process_High = &H80
    Process_RealTime = &H100
End Enum

'/* psapi g2
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, _
                                                        ByVal lBuffer As Long, _
                                                        ByRef lAllocate As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal lProcess As Long, _
                                                               ByVal hModule As Long, _
                                                               ByVal ModuleName As String, _
                                                               ByVal lSize As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, _
                                                                                           ByVal hModule As Long, _
                                                                                           ByVal ModuleName As String, _
                                                                                           ByVal nSize As Long) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal lProcess As Long, _
                                                             ByRef lphModule As Long, _
                                                             ByVal lBuffer As Long, _
                                                             ByRef lAllocate As Long) As Long

'/* process api
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, _
                                                         ByVal dwExitCode As Long) As Long

Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, _
                                                           ByVal nPriority As Long) As Long

Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long

Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, _
                                                          ByVal dwPriorityClass As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, _
                                                                  ByVal lProcessID As Long) As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                        uProcess As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       uProcess As PROCESSENTRY32) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, _
                                                          ByVal uExitCode As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                     ByVal blnheritHandle As Long, _
                                                     ByVal dwAppProcessId As Long) As Long

Private Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       uProcess As MODULEENTRY32) As Long

Private Declare Function Module32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, _
                                                          ByRef lpme As MODULEENTRY32) As Boolean
                                                          
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, _
                                                                ByRef lpExitCode As Long) As Boolean

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Event eNComplete(ByVal sTask As String)
Public Event eNErrorCond(ByVal sRoutine As String, ByVal sError As String)



'***************************************************************************************
'*                                  PROCESS MANAGEMENT
'***************************************************************************************

'>  Process Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : Process_Enumerate
'* Purpose        : list running processes
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function Process_Enumerate() As Collection

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim cTemp               As Collection
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    Set cTemp = New Collection
    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        .dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* name/id/thrdcnt/priority/parent
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1) + Chr$(31) + _
            CStr(.th32ProcessID)
            cTemp.Add sItem
        Loop
        CloseHandle lSnapShot
    End With
    '/* success
    If cTemp.Count > 0 Then
        Set Process_Enumerate = cTemp
        RaiseEvent eNComplete("Processes have been enumerated.. Found: " + CStr(cTemp.Count) + " items")
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_Enumerate", CStr(err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function

'* Name           : Process_EnumG2
'* Purpose        : list running processes (psapi)
'* Inputs         : none
'* Outputs        : collection
'*********************************************
Public Function Process_EnumG2() As Collection

Dim lBuffer            As Long
Dim lAllocate          As Long
Dim lItems             As Long
Dim aProcessID()       As Long
Dim lDimension         As Long
Dim aModule(1 To 1024) As Long
Dim lReturn            As Long
Dim lSize              As Long
Dim lHandle            As Long
Dim lCount             As Long
Dim sProcess           As String
Dim cTemp              As Collection

On Error GoTo Handler

    Set cTemp = New Collection
    '/* size our array
    lBuffer = 8
    lAllocate = 96
    Do While lBuffer <= lAllocate
        lBuffer = lBuffer * 2
        ReDim aProcessID(lBuffer / 4) As Long
        lReturn = EnumProcesses(aProcessID(1), lBuffer, lAllocate)
    Loop
    
    '/* loop through return
    lItems = lAllocate / 4
    For lCount = 1 To lItems
        '/* open process
        lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aProcessID(lCount))
        '/* invalid
        If lHandle = 0 Then GoTo skip
        
        '/* size module array
        lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lDimension)
        lReturn = EnumProcessModules(lHandle, aModule(1), lDimension, lDimension)
        '/* get the file name
        If Not lReturn = 0 Then
            sProcess = Space$(MAX_PATH)
            lSize = 500
            'Get the module file name
            lReturn = GetModuleFileNameExA(lHandle, aModule(1), sProcess, lSize)
            sProcess = Left$(sProcess, lReturn)
            '/* add to collection name/id/path
            cTemp.Add Mid$(sProcess, InStrRev(sProcess, Chr$(92)) + 1) + Chr$(31) + _
                CStr(aProcessID(lCount)) + Chr$(31) + _
                sProcess
        End If
skip:
        '/* close handle
        CloseHandle lHandle
    Next lCount
    
    '/* success
    If cTemp.Count > 0 Then
        Set Process_EnumG2 = cTemp
        RaiseEvent eNComplete("Processes have been enumerated.. Found: " + CStr(cTemp.Count) + " items")
    End If

On Error GoTo 0
Exit Function

Handler:
    If Not lHandle = 0 Then CloseHandle lHandle
    RaiseEvent eNErrorCond("Process_EnumG2", CStr(err.LastDllError))

End Function

'* Name           : Process_Exists
'* Purpose        : test process existence
'* Inputs         : process name, use loose naming
'* Outputs        : boolean
'*********************************************
Public Function Process_Exists(ByVal sProcess As String, _
                               Optional ByVal bLoose As Boolean) As Boolean

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        .dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* compare name against process list
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1)
            If Not bLoose Then
                If sProcess = sItem Then
                    '/* success
                    Process_Exists = True
                    RaiseEvent eNComplete("Process: " + sProcess + " Exists.")
                    Exit Do
                End If
            Else
                If LCase$(sProcess) = LCase$(sItem) Then
                    Process_Exists = True
                    RaiseEvent eNComplete("Process: " + sProcess + " Exists.")
                    Exit Do
                End If
            End If
        Loop
        CloseHandle lSnapShot
    End With
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_Exists", CStr(err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function


'* Name           : Process_ExistsG2
'* Purpose        : test process existence (psapi)
'* Inputs         : process name, use loose naming
'* Outputs        : boolean
'*********************************************
Public Function Process_ExistsG2(ByVal sProcess As String, _
                                 Optional ByVal bLoose As Boolean) As Boolean
'/* return a process id from a process name

Dim lBuffer             As Long
Dim lAllocate           As Long
Dim lItems              As Long
Dim aProcessID()        As Long
Dim lDimension          As Long
Dim aModule(1 To 1024)  As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim lHandle             As Long
Dim lCount              As Long
Dim sTemp               As String

On Error GoTo Handler

    '/* size our array
    lBuffer = 8
    lAllocate = 96
    Do While lBuffer <= lAllocate
        lBuffer = lBuffer * 2
        ReDim aProcessID(lBuffer / 4) As Long
        lReturn = EnumProcesses(aProcessID(1), lBuffer, lAllocate)
    Loop
    
    '/* loop through return
    lItems = lAllocate / 4
    For lCount = 1 To lItems
        '/* open process
        lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aProcessID(lCount))
        '/* invalid
        If lHandle = 0 Then GoTo skip
        
        '/* size module array
        lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lDimension)
        lReturn = EnumProcessModules(lHandle, aModule(1), lDimension, lDimension)
        '/* get the file name
        If Not lReturn = 0 Then
            sTemp = Space$(MAX_PATH)
            lSize = 500
            'Get the module file name
            lReturn = GetModuleFileNameExA(lHandle, aModule(1), sTemp, lSize)
            sTemp = Left$(sTemp, lReturn)
            '/* add to collection name/id/path
            sTemp = Mid$(sTemp, InStrRev(sTemp, Chr$(92)) + 1)
            If bLoose Then
                If LCase$(sTemp) = LCase$(sProcess) Then
                    Process_ExistsG2 = True
                    CloseHandle lHandle
                    Exit For
                End If
            Else
                If sTemp = sProcess Then
                    Process_ExistsG2 = True
                    CloseHandle lHandle
                    Exit For
                End If
            End If
        End If
skip:
        '/* close handle
        CloseHandle lHandle
    Next lCount

On Error GoTo 0
Exit Function

Handler:

End Function

'* Name           : Process_Terminate
'* Purpose        : terminate a process
'* Inputs         : process name
'* Outputs        : boolean
'*********************************************
Public Function Process_Terminate(ByVal lProcess As Long) As Boolean

Dim lExitCode           As Long
Dim lHandle             As Long
Dim lReturn             As Long

On Error GoTo Handler

    '/* get process handle
    lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_CTERMINATE, False, lProcess)
    If lHandle = 0 Then GoTo Handler
    '/* get exit code
    GetExitCodeProcess lHandle, lExitCode
    If lExitCode = 0 Then GoTo Handler
    '/* kill process
    lReturn = TerminateProcess(lHandle, lExitCode)
    If lReturn = 0 Then GoTo Handler
    
    '/* success
    CloseHandle lHandle
    Process_Terminate = True
    RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " has been terminated.")

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_Terminate", CStr(err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Process_GetClass
'* Purpose        : return the process class
'* Inputs         : process name
'* Outputs        : long
'*********************************************
Public Function Process_GetClass(ByVal lProcess As Long) As Long
'/* return the process class level

Dim lReturn             As Long
Dim sResult             As String
Dim lHandle             As Long

On Error GoTo Handler

    '/* get process handle
    lHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, lProcess)
    If lHandle = 0 Then
        RaiseEvent eNErrorCond("Process_GetClass", "Could not aquire process handle.")
        Exit Function
    End If
    '/* return class
    lReturn = GetPriorityClass(lHandle)
    
    If Not lReturn = -1 Then
        Process_GetClass = lReturn
        Select Case lReturn
        '/* Idle
        Case &H40
            sResult = "Idle"
        '/* Normal
        Case &H20
            sResult = "Normal"
        '/* High
        Case &H80
            sResult = "High"
        '/* Priority
        Case &H100
            sResult = "Realtime"
        '/* unknown
        Case Else
            sResult = "Unknown"
        End Select
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " priority class is " + sResult + Chr(46))
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseHandle lHandle
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Process_GetClass", CStr(err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle

End Function

'* Name           : Module_Enumerate
'* Purpose        : list all process modules
'* Inputs         : process name
'* Outputs        : collection
'*********************************************
Public Function Module_Enumerate(ByVal lProcess As Long) As Collection

Dim lReturn             As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim cTemp               As Collection
Dim tModule             As MODULEENTRY32

On Error GoTo Handler

    '/* set up variables
    Set cTemp = New Collection
    '/* create a snapshot
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lProcess)
    If lSnapShot = -1 Then
        RaiseEvent eNComplete("This Process has No Modules associated with it.")
        Exit Function
    End If
    '/* prepare struct
    tModule.dwSize = Len(tModule)
    '/* fill struct
    lReturn = Module32First(lSnapShot, tModule)
    
    '/* get first entry
    If Not lReturn = 0 Then
        With tModule
            sItem = .szModule + Chr$(31) + _
            .szExePath + Chr$(31) + _
            CStr(.GlblcntUsage) + Chr$(31) + _
            CStr(.hModule) + Chr$(31) + _
            CStr(.modBaseAddr) + Chr$(31) + _
            CStr(.ProccntUsage) + Chr$(31) + _
            CStr(.th32ProcessID) + Chr$(31)
        End With
        cTemp.Add sItem
    Else
        GoTo Handler
    End If
    
    '/* loop through entries
    '/* name/exe path/id32/usage/modid/address/procuse/prcid
    Do While lReturn
        lReturn = Module32Next(lSnapShot, tModule)
        If lReturn = 0 Then Exit Do
        With tModule
            sItem = CStr(.szModule) + Chr$(31) + _
            CStr(.szExePath) + Chr$(31) + _
            CStr(.GlblcntUsage) + Chr$(31) + _
            CStr(.hModule) + Chr$(31) + _
            CStr(.modBaseAddr) + Chr$(31) + _
            CStr(.ProccntUsage) + Chr$(31) + _
            CStr(.th32ProcessID) + Chr$(31)
        End With
        cTemp.Add sItem
    Loop
    
    '/* success
    If cTemp.Count > 0 Then
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " modules have been enumerated.")
        Set Module_Enumerate = cTemp
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Module_Enumerate", CStr(err.LastDllError))
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function

'* Name           : Module_EnumG2
'* Purpose        : list all process modules using psapi
'* Inputs         : process name
'* Outputs        : collection
'*********************************************
Public Function Module_EnumG2(ByVal lProcess As Long) As Collection

Dim aModule(1 To 1024)  As Long
Dim lBuffer             As Long
Dim lCount              As Long
Dim lHandle             As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim sModule             As String
Dim cTemp               As Collection

On Error GoTo Handler

    Set cTemp = New Collection
    lSize = 500
    '/* get the process handle
    lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcess)
    If lHandle = 0 Then GoTo Handler
    '/* size the buffer
    lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lBuffer)
    '/* get the modules
    lReturn = EnumProcessModules(lHandle, aModule(1), lBuffer, lBuffer)
    sModule = Space$(MAX_PATH)
    
    '/* get the first file name
    lReturn = GetModuleFileNameExA(lHandle, aModule(1), sModule, lSize)
    '/* no modules
    If lReturn = 0 Then
        RaiseEvent eNComplete("This Process has No Modules associated with it.")
        CloseHandle lHandle
        Exit Function
    End If
    
    Do
        lCount = lCount + 1
        sModule = Space$(MAX_PATH)
        lSize = 500
        lReturn = GetModuleFileNameExA(lHandle, aModule(lCount), sModule, lSize)
        sModule = Left$(sModule, lReturn)
        '/* add to collection name/path/handle/parent id
        If LenB(Trim$(sModule)) Then
            cTemp.Add Mid$(sModule, InStrRev(sModule, Chr$(92)) + 1) + Chr$(31) + _
                sModule + Chr$(31) + _
                CStr(aModule(lCount)) + Chr$(31) + _
                CStr(lProcess)
        End If
    Loop While aModule(lCount) > 0

    '/* cleanup
    CloseHandle lHandle
    
    '/* success
    If cTemp.Count > 0 Then
        Set Module_EnumG2 = cTemp
        RaiseEvent eNComplete("Process ID: " + CStr(lProcess) + " modules have been enumerated.")
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eNErrorCond("Module_EnumG2", CStr(err.LastDllError))
    If Not lHandle = 0 Then CloseHandle lHandle
    
End Function

'* Name           : Return_ProcessID
'* Purpose        : return process id from name (kernal32)
'* Inputs         : process name
'* Outputs        : long
'*********************************************
Public Function Return_ProcessID(ByVal sProcess As String) As Long
'/* return a process id from a process name

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        tProc.dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* compare name against process list
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1)
            If LCase$(sProcess) = LCase$(sItem) Then
                Return_ProcessID = .th32ProcessID
                Exit Do
            End If
        Loop
        CloseHandle lSnapShot
    End With
    
On Error GoTo 0
Exit Function

Handler:
    If Not lSnapShot = 0 Then CloseHandle lSnapShot
    Return_ProcessID = 0

End Function

'* Name           : Return_ProcessIDG2
'* Purpose        : return process id from name (psapi)
'* Inputs         : process name
'* Outputs        : long
'*********************************************
Public Function Return_ProcessIDG2(ByVal sProcess As String) As Long
'/* return a process id from a process name

Dim lBuffer             As Long
Dim lAllocate           As Long
Dim lItems              As Long
Dim aProcessID()        As Long
Dim lDimension          As Long
Dim aModule(1 To 1024)  As Long
Dim lReturn             As Long
Dim lSize               As Long
Dim lHandle             As Long
Dim lCount              As Long
Dim sTemp               As String

On Error GoTo Handler

    '/* size our array
    lBuffer = 8
    lAllocate = 96
    Do While lBuffer <= lAllocate
        lBuffer = lBuffer * 2
        ReDim aProcessID(lBuffer / 4) As Long
        lReturn = EnumProcesses(aProcessID(1), lBuffer, lAllocate)
    Loop
    
    '/* loop through return
    lItems = lAllocate / 4
    For lCount = 1 To lItems
        '/* open process
        lHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aProcessID(lCount))
        '/* invalid
        If lHandle = 0 Then GoTo skip
        
        '/* size module array
        lReturn = EnumProcessModules(lHandle, aModule(1), 1024, lDimension)
        lReturn = EnumProcessModules(lHandle, aModule(1), lDimension, lDimension)
        '/* get the file name
        If Not lReturn = 0 Then
            sTemp = Space$(MAX_PATH)
            lSize = 500
            'Get the module file name
            lReturn = GetModuleFileNameExA(lHandle, aModule(1), sTemp, lSize)
            sTemp = Left$(sTemp, lReturn)
            '/* add to collection name/id/path
            sTemp = Mid$(sTemp, InStrRev(sTemp, Chr$(92)) + 1)
            If LCase$(sTemp) = LCase$(sProcess) Then
                Return_ProcessIDG2 = aProcessID(lCount)
                CloseHandle lHandle
                Exit For
            End If
        End If
skip:
        '/* close handle
        CloseHandle lHandle
    Next lCount

On Error GoTo 0
Exit Function

Handler:
    Return_ProcessIDG2 = 0

End Function

Public Function ProcessExists(ByVal sProcess As String) As Boolean

Dim lEntry              As Long
Dim lSnapShot           As Long
Dim sItem               As String
Dim tProc               As PROCESSENTRY32

On Error GoTo Handler

    With tProc
        '/* create our snapshot
        lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        .dwSize = Len(tProc)
        '/* first entry
        lEntry = Process32First(lSnapShot, tProc)
        '/* loop through entries
        Do While lEntry
            lEntry = Process32Next(lSnapShot, tProc)
            If lEntry = 0 Then Exit Do
            '/* compare name against process list
            sItem = Left$(.szexeFile, InStr(1, .szexeFile, Chr$(0)) - 1)

                If LCase$(sProcess) = LCase$(sItem) Then
                    ProcessExists = True
                    Exit Do
                End If
        Loop
        CloseHandle lSnapShot
    End With
    
On Error GoTo 0
Exit Function

Handler:
    If Not lSnapShot = 0 Then CloseHandle lSnapShot

End Function

