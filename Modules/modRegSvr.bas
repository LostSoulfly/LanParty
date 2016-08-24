Attribute VB_Name = "modRegSvr"
Option Explicit

Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Public Function RegServer(ByVal FileName As String) As Boolean
'USAGE: PASS FULL PATH OF ACTIVE .DLL OR
'OCX YOU WANT TO REGISTER
If FileExists(FileName) Then RegServer = RegSvr32(FileName, False)

If RegServer = False Then
    If IsHost64Bit Then
        ExecFile Environ("SYSTEMROOT") & "\SYSWOW64\regsvr32.exe", "/s " & Chr(34) & FileName & Chr(34)
    Else
        ExecFile "regsvr32", "/s " & Chr(34) & FileName & Chr(34)
    End If
End If

'adddebug "RegServer: " & RegServer
End Function

Public Function UnRegServer(ByVal FileName As String) As Boolean
'USAGE: PASS FULL PATH OF ACTIVE .DLL OR
'OCX YOU WANT TO UNREGISTER
If FileExists(FileName) Then UnRegServer = RegSvr32(FileName, True)
End Function
    
Private Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean

Dim lLib As Long
Dim lProcAddress As Long
Dim lThreadID As Long
Dim lSuccess As Long
Dim lExitCode As Long
Dim lThread As Long
Dim bAns As Boolean
Dim sPurpose As String

sPurpose = IIf(bUnReg, "DllUnregisterServer", _
  "DllRegisterServer")

If Not FileExists(FileName) Then Exit Function

lLib = LoadLibraryRegister(FileName)
'could load file
If lLib = 0 Then Exit Function

lProcAddress = GetProcAddressRegister(lLib, sPurpose)

'adddebug "DLL Reg: " & sPurpose & " - " & CStr(lProcAddress)

If lProcAddress = 0 Then
  'Not an ActiveX Component
   FreeLibraryRegister lLib
   Exit Function
Else
   lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
   If lThread Then
        lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
        If Not lSuccess Then
           Call GetExitCodeThread(lThread, lExitCode)
           Call ExitThread(lExitCode)
           bAns = False
           Exit Function
        Else
           bAns = True
        End If
        CloseHandle lThread
        FreeLibraryRegister lLib
   End If
End If
    RegSvr32 = bAns
End Function





