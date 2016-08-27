Attribute VB_Name = "modVersion"
Option Explicit

Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Declare Function RtlGetVersion Lib "NTDLL" (ByRef lpVersionInformation As Long) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Public Enum OSEnum
    WindowsXP = 1
    WindowsVista = 2
    Windows7 = 3
    Windows8 = 4
    Windows10 = 5
    Server2003 = 6
    WindowsDontKnow = 7
End Enum

Public OSVer As Integer

Public Function NativeGetVersion() As String
Dim tOSVw(&H54) As Long
    tOSVw(0) = &H54 * &H4
    Call RtlGetVersion(tOSVw(0))
    'NativeGetVersion = Join(Array(tOSVw(1), tOSVw(2), tOSVw(3)), ".")
    NativeGetVersion = VersionToName(Join(Array(tOSVw(1), tOSVw(2)), "."))
End Function

Public Function GetAppVersion() As String
    GetAppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function GetOSVer() As Integer
    NativeGetVersion
    GetOSVer = OSVer
End Function

Public Function VersionToName(ByRef sVersion As String) As String
    Select Case sVersion
        Case "5.1": VersionToName = "Windows XP": OSVer = OSEnum.WindowsXP
        Case "5.3": VersionToName = "Windows 2003 (SERVER)": OSVer = OSEnum.Server2003
        Case "6.0": VersionToName = "Windows Vista": OSVer = OSEnum.WindowsVista
        Case "6.1": VersionToName = "Windows 7": OSVer = OSEnum.Windows7
        Case "6.2": VersionToName = "Windows 8": OSVer = OSEnum.Windows8
        Case "6.3": VersionToName = "Windows 8.1": OSVer = OSEnum.Windows8
        Case "10.0": VersionToName = "Windows 10": OSVer = OSEnum.Windows10
        Case Else: VersionToName = "Unknown": OSVer = OSEnum.WindowsDontKnow
    End Select
End Function

Public Function IsHost64Bit() As Boolean
    Dim Handle As Long
    Dim Is64Bit As Boolean

    ' Assume initially that this is not a WOW64 process
    Is64Bit = False

    ' Then try to prove that wrong by attempting to load the
    ' IsWow64Process function dynamically
    Handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    ' The function exists, so call it
    If Handle <> 0 Then
        IsWow64Process GetCurrentProcess(), Is64Bit
    End If

    ' Return the value
    IsHost64Bit = Is64Bit
End Function

Public Sub CheckUpdate()
On Error Resume Next
'Download version file
If FileExists(App.Path & "\Version.txt") Then Kill App.Path & "\version.txt"
If DownloadFile("http://trollparty.org/LAN/version.txt", App.Path & "\Version.txt") Then
    'version file was downloaded
    Dim strFile() As String
    Dim strDownload As String
    Dim strVersion As String
    Dim strChangeLog As String
    strFile = Split(LoadFile(App.Path & "\Version.txt"), vbNewLine)
    'MsgBox LoadFile(App.Path & "\Version.txt")
    If UBound(strFile) >= 1 Then
        If Not InStr(1, strFile(0), ".") >= 1 Or Len(strFile(0)) > 10 Then AddUserChat "Update file is corrupt or incorrect format!", "System", False: Exit Sub
        strVersion = Trim$(Replace(strFile(0), ".", ""))
        strDownload = Trim$(strFile(1))
        
        Dim i As Long
        For i = 2 To UBound(strFile)
            strChangeLog = strChangeLog & strFile(i) & vbNewLine
        Next i
        
        strChangeLog = strChangeLog & "Download URL: " & strDownload
           
        If FileExists(App.Path & "\Version.txt") Then Kill App.Path & "\Version.txt"
            If CLng(strVersion) = CLng(App.Major & App.Minor & App.Revision) Then
                AddUserChat "Your version is up to date!", "System", False
                Exit Sub
            ElseIf CLng(strVersion) < CLng(App.Major & App.Minor & App.Revision) Then
                AddUserChat "You have a newer version than the server!", "System", False
                Exit Sub
            ElseIf CLng(strVersion) > CLng(App.Major & App.Minor & App.Revision) Then
                AddUserChat "There is an update available: " & App.Major & App.Minor & App.Revision & " -> " & strVersion, "System", False
                If MsgBox(strChangeLog, vbYesNo, "Update Available") = vbYes Then
                        AddUserChat "Downloading update, please wait..", "System", False
                        DoEvents
                    If DownloadFile(strDownload, App.Path & "\LanParty.New.exe") = True Then
                        AddUserChat "Update downloaded! Updating..", "System", False
                        Sleep 1500
                        If Not isAdmin Then
                            MsgBox "You aren't running this program as an Administrator." & vbNewLine & vbNewLine & _
                            "The update may fail to initialize due to this. If it doesn't work, please run it as an Administrator!", vbOKOnly, "Elevation Required!"
                        End If
                        Shell App.Path & "\LanParty.New.exe"
                        DoEvents
                        frmMain.DoExit
                        
                    Else
                        AddUserChat "There was an issue downloading the update..", "System", False
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                
            End If
    Else
        If FileExists(App.Path & "\Version.txt") Then Kill App.Path & "\Version.txt"
        AddUserChat "Version file is corrupt or incorrect format!", "System", False: Exit Sub
    End If
End If

End Sub
