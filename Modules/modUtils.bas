Attribute VB_Name = "modUtils"
Option Explicit

Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000
 
Const SW_HIDE = 0
Const SW_NORMAL = 1
 
Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10
Public Const BIF_VALIDATE = &H20
Public Const BIF_NEWDIALOGSTYLE = &H40
Public Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Public Const BIF_BROWSEINCLUDEURLS = &H80
Public Const BIF_UAHINT = &H100
Public Const BIF_NONEWFOLDERBUTTON = &H200
Public Const BIF_NOTRANSLATETARGETS = &H400
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const BIF_BROWSEINCLUDEFILES = &H4000
Public Const BIF_SHAREABLE = &H8000
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Public Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Private mstrSTARTFOLDER As String

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias _
    "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
    
Private UniqueKeyCharacters(93) As Byte

Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean

   DeleteUrlCacheEntry sSourceUrl

   DoEvents
   
   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS

End Function

Public Function GetFolder(ByVal hWndModal As Long, Optional StartFolder As String = "", Optional Title As String = "Please select a folder:", _
   Optional IncludeFiles As Boolean = False, Optional IncludeNewFolderButton As Boolean = False) As String
    Dim bInf As BrowseInfo
    Dim RetVal As Long
    Dim PathID As Long
    Dim RetPath As String
    Dim Offset As Integer
    'Set the properties of the folder dialog
    bInf.hwndOwner = hWndModal
    bInf.pIDLRoot = 0
    bInf.lpszTitle = Title
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
    If IncludeFiles Then bInf.ulFlags = bInf.ulFlags Or BIF_BROWSEINCLUDEFILES
    If IncludeNewFolderButton Then bInf.ulFlags = bInf.ulFlags Or BIF_NEWDIALOGSTYLE
    If StartFolder <> "" Then
       mstrSTARTFOLDER = StartFolder & vbNullChar
       bInf.lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc) 'get address of function.
   End If
    'Show the Browse For Folder dialog
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
         'Trim off the null chars ending the path
         'and display the returned folder
         Offset = InStr(RetPath, Chr$(0))
         GetFolder = Left$(RetPath, Offset - 1)
         'Free memory allocated for PIDL
         CoTaskMemFree PathID
    Else
         GetFolder = ""
    End If
End Function
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
   On Error Resume Next
   Dim lpIDList As Long
   Dim Ret As Long
   Dim sBuffer As String
   Select Case uMsg
       Case BFFM_INITIALIZED
           Call SendMessage(hwnd, BFFM_SETSELECTION, 1, mstrSTARTFOLDER)
       Case BFFM_SELCHANGED
           sBuffer = Space(MAX_PATH)
           Ret = SHGetPathFromIDList(lp, sBuffer)
           If Ret = 1 Then
               Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
           End If
   End Select
   BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
 GetAddressofFunction = add
End Function
 
 
Public Sub ShowInTheTaskbar(hwnd As Long, bShow As Boolean)
    Dim lStyle As Long
    
    ShowWindow hwnd, SW_HIDE
    
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle
    
    App.TaskVisible = bShow
    
    ShowWindow hwnd, SW_NORMAL
End Sub
 
Public Function IsVisibleInTheTaskbar(hwnd As Long) As Boolean
    Dim lStyle As Long
    
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    If lStyle And WS_EX_APPWINDOW Then
        IsVisibleInTheTaskbar = True
    End If
End Function

Public Function VolumeSerialNumber() As String

Dim VolLabel As String
Dim VolSize As Long
Dim Serial As Long
Dim MaxLen As Long
Dim flags As Long
Dim Name As String
Dim NameSize As Long
Dim s As String
Dim Ret As Boolean
Dim RootPath(4) As String
Dim i As Integer

RootPath(0) = "C:\"
RootPath(1) = Environ$("SYSTEMDRIVE") & "\"
RootPath(2) = "E:\"
RootPath(3) = "F:\"
RootPath(4) = "G:\"


For i = 0 To UBound(RootPath)

    Ret = GetVolumeSerialNumber(RootPath(i), VolLabel, VolSize, Serial, MaxLen, flags, Name, NameSize)

    If Ret Then
        'Create an 8 character string
        s = Format(Hex(Serial), "00000000")
        VolumeSerialNumber = s
    End If

Next i

If LenB(VolumeSerialNumber) < 16 Then
    VolumeSerialNumber = CalculateAdler(VolumeSerialNumber & Environ$("COMPUTERNAME"))
End If

If LenB(VolumeSerialNumber) < 16 Then
    VolumeSerialNumber = CalculateAdler(VolumeSerialNumber & Environ$("USERNAME"))
End If


If LenB(VolumeSerialNumber) < 16 Then
    VolumeSerialNumber = VolumeSerialNumber & CalculateAdler(VolumeSerialNumber & App.Path)
End If

If LenB(VolumeSerialNumber) < 8 Then
    VolumeSerialNumber = Format(Hex(1 + Rnd * 999999999), "00000000")
End If

If Len(VolumeSerialNumber) > 8 Then
    Dim intRemove As Integer
    intRemove = VolumeSerialNumber - 8
    VolumeSerialNumber = Mid(VolumeSerialNumber, 1, Len(VolumeSerialNumber) - intRemove)
End If

Do While Len(VolumeSerialNumber) < 8
    VolumeSerialNumber = VolumeSerialNumber & "0"
Loop

If Settings.blDebug Then AddChat "UniqueID Init: " & VolumeSerialNumber

End Function

Public Function CalculateAdler(Data As String) As String
Dim objAdler32 As New clsAdler32
Dim lngAdler32 As Long, myByte() As Byte
myByte = Data
CalculateAdler = Hex$(objAdler32.Adler32(lngAdler32, myByte, UBound(myByte)))

Set objAdler32 = Nothing

End Function

Public Function isAdmin() As Boolean
If IsUserAnAdmin = 1 Then isAdmin = True
End Function

Public Function LoadFile(dFile As String) As String

    Dim ff As Integer

    On Error Resume Next

    ff = FreeFile
    Open dFile For Binary As #ff
        LoadFile = Space(LOF(ff))
        Get #ff, , LoadFile
    Close #ff

End Function

Public Function WriteFile(strOutput As String, Optional strFile As String, Optional Overwrite As Boolean = False) As String

    Dim ff As Integer

    On Error Resume Next

    If LenB(strFile$) = 0 Then strFile = App.Path & "\debug.txt"

    ff = FreeFile
    If Not Overwrite Then
        Open strFile For Binary Access Read Write As #ff
        Seek #ff, LOF(ff)
        Put #ff, , strOutput & vbCrLf
    Else
        Open strFile For Output As #ff
        Print #ff, strOutput
    End If
        
        
    Close #ff

End Function

Public Function GetMsgTypeName(MsgType As Long) As String
GetMsgTypeName = ""
Select Case MsgType
Case Is = 1
    GetMsgTypeName = "LAuth"
Case Is = 2
    GetMsgTypeName = "LBeacon"
Case Is = 3
    GetMsgTypeName = "LGoodbye"
Case Is = 4
    GetMsgTypeName = "LDebug"
Case Is = 5
    GetMsgTypeName = "LPing"
Case Is = 6
    GetMsgTypeName = "LPong"
Case Is = 7
    GetMsgTypeName = "LSuggest"
Case Is = 8
    GetMsgTypeName = "LChat"
Case Is = 9
   GetMsgTypeName = "LPrivateChat"
Case Is = 10
    GetMsgTypeName = "LVote"
Case Is = 11
    GetMsgTypeName = "LChangeName"
Case Is = 12
    GetMsgTypeName = "LSyncAdmin"
Case Is = 13
    GetMsgTypeName = "LLanAdmin"
Case Is = 14
    GetMsgTypeName = "LNowPlaying"
Case Is = 15
    GetMsgTypeName = "LReqList"
Case Is = 16
    GetMsgTypeName = "LDrew"
Case Is = 17
    GetMsgTypeName = "LCrypted"
Case Else
    GetMsgTypeName = "Unknown"
End Select

End Function

Public Sub WriteLog(Data As String, File As String)
On Error Resume Next
    Dim ff As Integer
    Dim eDate As String
        ff = FreeFile

        ' Update database
        Open (File) For Append As #ff
            Print #ff, (Now & ": " & Data)
        Close #ff
    
End Sub

Public Sub AddDebug(Text As String, Optional blToFile As Boolean = False)
If blToFile Then WriteFile "[DEBUG] " & Text & vbCrLf
If Settings.blDebug = False Then Exit Sub
    If frmChat.Visible = True Then
        frmChat.txtChat.Text = frmChat.txtChat.Text & Time & " [DEBUG] " & Text & vbCrLf
    End If
End Sub

Public Sub AddUserChat(Text As String, Name As String, Optional EnhSec As Boolean)
    'If frmChat.Visible = True Then
    
    If EnhSec = True Then
        frmChat.txtChat.Text = frmChat.txtChat.Text & Time & " <" & Name & "> " & Text & vbCrLf
    Else
        frmChat.txtChat.Text = frmChat.txtChat.Text & Time & " [" & Name & "] " & Text & vbCrLf
    End If
End Sub

Public Sub AddChat(Text As String)
    'If frmChat.Visible = True Then
        frmChat.txtChat.Text = frmChat.txtChat.Text & Time & " " & Text & vbCrLf
    'End If
End Sub

Public Sub InitializeUniqueID()
Dim MyCompID As String

    MyCompID = VolumeSerialNumber
    
    'if the serial doesn't match the last saved one, we call FirstRun sub
    If Not Settings.UniqueID = MyCompID Then Call FirstRun
    
    Settings.UniqueID = MyCompID

End Sub

Public Function FileExists(ByRef sFileName As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(sFileName) And vbDirectory) <> vbDirectory
End Function

Public Function ExecFile(FilePath As String, FileArgs As String, Optional Operation As String = "open", Optional Directory As String = vbNullString, Optional Visible As Long = 1) As Long
Dim Ret As Long
    Ret = ShellExecute(frmMain.hwnd, Operation, FilePath, FileArgs, Directory, Visible)
    'ret = ShellExecute(0, Operation, FilePath, FileArgs, Directory, Visible)
    AddDebug "execFile (shellex): (" & FilePath & ") = " & CStr(Ret)
    DoEvents
    If Visible > 6 Then Visible = 0
    If Ret <= 32 Then
        Ret = Shell(FilePath & FileArgs, Visible)
        AddDebug "execFile (shell): (" & FilePath & ") = " & CStr(Ret)
    End If
    
    ExecFile = Ret

End Function

Public Sub SetCaption(Text As String)
    frmMain.Caption = "LanParty - " & Text
End Sub

Public Function AddToString(Data As String, ByRef ToString As String, Optional blNewLine As Boolean = True)

    ToString = ToString & Data & IIf(blNewLine, vbNewLine, "")

End Function

Public Function DirExists(ByVal Path As String) As Boolean
On Error Resume Next
    If Dir(Path) <> "" Then
        DirExists = True
    Else
        DirExists = False
    End If

End Function

Public Function FormatToLocalPath(ByVal Path As String) As String
    If InStr(1, LCase$(Path), LCase$(App.Path & "\")) > 0 Then
        FormatToLocalPath = Right$(Path, Len(Path) - Len(App.Path))
    Else
        FormatToLocalPath = Path
    End If
End Function

Public Function FullPathFromLocal(ByVal Path As String) As String

If InStr(1, Path, ":") > 0 Then
    FullPathFromLocal = Path
    Exit Function
End If

'make sure the first character is a slash
    If Not Left$(Path, 1) = "\" Then
        FullPathFromLocal = App.Path & "\" & Path
    Else
        FullPathFromLocal = App.Path & Path
    End If

End Function

Public Function FixFilePath(ByVal Path As String) As String
'AddDebug "FixFilePath: " & Path

If Len(Path$) = 0 Then Exit Function

    Path = Replace(Path, "\\", "\")
    Path = Replace(Path, "\\", "\")
    Path = Replace(Path, "//", "/")
    If FileExists(FullPathFromLocal(Path)) Then FixFilePath = FullPathFromLocal(Path): Exit Function
    If FileExists(FormatToLocalPath(Path)) Then FixFilePath = FormatToLocalPath(Path): Exit Function
    If FileExists(Path) Then FixFilePath = Path: Exit Function
    If Left$(Path, 1) = "\" Then
        If FileExists(Environ("WINDIR") & "\System32" & Path) Then FixFilePath = Environ("WINDIR") & "\System32" & Path
    Else
        If FileExists(Environ("WINDIR") & "\System32\" & Path) Then FixFilePath = Environ("WINDIR") & "\System32\" & Path
    End If
End Function

Public Function FixCmdArgs(ByVal Args As String) As String

If Len(Args$) = 0 Then Exit Function

FixCmdArgs = Replace(Args, "%SCREENX%", GetScreenX)
FixCmdArgs = Replace(FixCmdArgs, "%SCREENY%", GetScreenY)
FixCmdArgs = Replace(FixCmdArgs, "%DATE%", Format(Now, "mm-dd-yyyy"))
FixCmdArgs = Replace(FixCmdArgs, "%TIME%", Format(Time, "hh-mm-ss"))
FixCmdArgs = Replace(FixCmdArgs, "%USERNAME%", Settings.UserName)

End Function


Public Function VerifyKeyAsString(ByVal Key As String) As Boolean

    VerifyKeyAsString = VerifyKey(StrConv(Key, vbFromUnicode))

End Function

Public Function VerifyKey(ByRef Key() As Byte) As Boolean

    Dim i As Long
    
    VerifyKey = True
    
    Do
        If Key(i) < 32 Or Key(i) > 126 Then VerifyKey = False: Exit Function
        i = i + 1
    Loop While i < UBound(Key)

End Function

Public Function VerifyKeyOld(ByRef Key() As Byte) As Boolean

    Dim i As Long
    
    VerifyKeyOld = True
    
    For i = 0 To UBound(Key)
        If Key(i) < 32 Or Key(i) > 126 Then VerifyKeyOld = False: Exit Function
    Next i

End Function

Public Sub InitUniqueKeyChars()
        
    'UniqueKeyCharacters = StrConv("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789~?!@#$%^&*()_+=-`,./<>';"":}{][\|", vbFromUnicode)
    
UniqueKeyCharacters(0) = 97
UniqueKeyCharacters(1) = 98
UniqueKeyCharacters(2) = 99
UniqueKeyCharacters(3) = 100
UniqueKeyCharacters(4) = 101
UniqueKeyCharacters(5) = 102
UniqueKeyCharacters(6) = 103
UniqueKeyCharacters(7) = 104
UniqueKeyCharacters(8) = 105
UniqueKeyCharacters(9) = 106
UniqueKeyCharacters(10) = 107
UniqueKeyCharacters(11) = 108
UniqueKeyCharacters(12) = 109
UniqueKeyCharacters(13) = 110
UniqueKeyCharacters(14) = 111
UniqueKeyCharacters(15) = 112
UniqueKeyCharacters(16) = 113
UniqueKeyCharacters(17) = 114
UniqueKeyCharacters(18) = 115
UniqueKeyCharacters(19) = 116
UniqueKeyCharacters(20) = 117
UniqueKeyCharacters(21) = 118
UniqueKeyCharacters(22) = 119
UniqueKeyCharacters(23) = 120
UniqueKeyCharacters(24) = 121
UniqueKeyCharacters(25) = 122
UniqueKeyCharacters(26) = 65
UniqueKeyCharacters(27) = 66
UniqueKeyCharacters(28) = 67
UniqueKeyCharacters(29) = 68
UniqueKeyCharacters(30) = 69
UniqueKeyCharacters(31) = 70
UniqueKeyCharacters(32) = 71
UniqueKeyCharacters(33) = 72
UniqueKeyCharacters(34) = 73
UniqueKeyCharacters(35) = 74
UniqueKeyCharacters(36) = 75
UniqueKeyCharacters(37) = 76
UniqueKeyCharacters(38) = 77
UniqueKeyCharacters(39) = 78
UniqueKeyCharacters(40) = 79
UniqueKeyCharacters(41) = 80
UniqueKeyCharacters(42) = 81
UniqueKeyCharacters(43) = 82
UniqueKeyCharacters(44) = 83
UniqueKeyCharacters(45) = 84
UniqueKeyCharacters(46) = 85
UniqueKeyCharacters(47) = 86
UniqueKeyCharacters(48) = 87
UniqueKeyCharacters(49) = 88
UniqueKeyCharacters(50) = 89
UniqueKeyCharacters(51) = 90
UniqueKeyCharacters(52) = 48
UniqueKeyCharacters(53) = 49
UniqueKeyCharacters(54) = 50
UniqueKeyCharacters(55) = 51
UniqueKeyCharacters(56) = 52
UniqueKeyCharacters(57) = 53
UniqueKeyCharacters(58) = 54
UniqueKeyCharacters(59) = 55
UniqueKeyCharacters(60) = 56
UniqueKeyCharacters(61) = 57
UniqueKeyCharacters(62) = 126
UniqueKeyCharacters(63) = 63
UniqueKeyCharacters(64) = 33
UniqueKeyCharacters(65) = 64
UniqueKeyCharacters(66) = 35
UniqueKeyCharacters(67) = 36
UniqueKeyCharacters(68) = 37
UniqueKeyCharacters(69) = 94
UniqueKeyCharacters(70) = 38
UniqueKeyCharacters(71) = 42
UniqueKeyCharacters(72) = 40
UniqueKeyCharacters(73) = 41
UniqueKeyCharacters(74) = 95
UniqueKeyCharacters(75) = 43
UniqueKeyCharacters(76) = 61
UniqueKeyCharacters(77) = 45
UniqueKeyCharacters(78) = 96
UniqueKeyCharacters(79) = 44
UniqueKeyCharacters(80) = 46
UniqueKeyCharacters(81) = 47
UniqueKeyCharacters(82) = 60
UniqueKeyCharacters(83) = 62
UniqueKeyCharacters(84) = 39
UniqueKeyCharacters(85) = 59
UniqueKeyCharacters(86) = 34
UniqueKeyCharacters(87) = 58
UniqueKeyCharacters(88) = 125
UniqueKeyCharacters(89) = 123
UniqueKeyCharacters(90) = 93
UniqueKeyCharacters(91) = 91
UniqueKeyCharacters(92) = 92
UniqueKeyCharacters(93) = 124
    
    'DoEvents
    
    'Dim i As Long
    'Dim strTest As String
    
    'For i = 0 To UBound(UniqueKeyCharacters)
    '     strTest = strTest & "UniqueKeyCharacters(" & i & ") = " & UniqueKeyCharacters(i) & vbNewLine
    'Next i
    
    'TestKeyGen 150
    
End Sub

Public Sub TestKeyGen(numTimes As Integer)

    Dim startTick As Long
    Dim stopTick As Long
    Dim l1 As Long
    Dim l2 As Long
    Dim l3 As Long
    Dim i As Integer
    Dim times As Integer
    Dim average1 As Long
    Dim average2 As Long
    Dim average3 As Long
    
For times = 1 To numTimes
    
    startTick = GetTickCount
    For i = 0 To 10
        GenUniqueKey 20
    Next i
    stopTick = GetTickCount
    l1 = stopTick - startTick
    average1 = average1 + l1
    startTick = GetTickCount
    For i = 0 To 10
        GenUniqueKey2 20
    Next i
    stopTick = GetTickCount
    l2 = stopTick - startTick
    average2 = average2 + l2
    
    startTick = GetTickCount
    For i = 0 To 10
        GenUniqueKey3 20
    Next i
    stopTick = GetTickCount
    l3 = stopTick - startTick
    average3 = average3 + l3
    
    'MsgBox "GenUniqueKey New. " & i & " iterations in " & l1 & "ms" & vbNewLine & _
    '"GenUniqueKey2 old. " & i & " iterations in " & l2 & "ms" & vbNewLine & vbNewLine & _
    '"Difference (2 - 1): " & (l2 - l1) & " ms" & vbNewLine & "GenKey3: " & l3

    WriteLog vbNewLine & "GenUniqueKey " & i & " iterations in " & l1 & "ms" & vbNewLine & _
    "GenUniqueKey2 " & i & " iterations in " & l2 & "ms" & vbNewLine & _
    "Difference (2 - 1): " & (l2 - l1) & " ms" & vbNewLine & _
    "GenUniqueKey3: " & i & " iterations in " & l3 & "ms" & vbNewLine, "test.txt"

DoEvents

Next times

    WriteLog "1: " & (average1 / numTimes), "test.txt"
    WriteLog "2: " & (average2 / numTimes), "test.txt"
    WriteLog "3: " & (average3 / numTimes), "test.txt"
    
End Sub

Public Function GenUniqueKey(Optional KeyLen As Integer = 0, Optional KeyGen As Integer = -1) As String
Dim i As Long
    If Settings.blDebug Then AddUserChat "Using KeyGen: " & KeyGen, "System", True

    If KeyGen = -1 Then KeyGen = Settings.KeyGen

    Select Case KeyGen
    
        Case Is = 0
            GenUniqueKey = GenUniqueKey1(KeyLen)
            
        Case Is = 1
            GenUniqueKey = GenUniqueKey1(KeyLen)
            
        Case Is = 2
            GenUniqueKey = GenUniqueKey2(KeyLen)
            
        Case Is = 3
            GenUniqueKey = GenUniqueKey3(KeyLen)
            
        Case Is = 4
            For i = 0 To KeyLen
                GenUniqueKey = GenUniqueKey & Int(Rnd * 9): Randomize
            Next i
            
        Case Is = 5
            Select Case Int(1 + Rnd * 3)
                Case Is = 1
                    GenUniqueKey = GenUniqueKey1(KeyLen)
                Case Is = 2
                    GenUniqueKey = GenUniqueKey2(KeyLen)
                Case Is = 3
                    GenUniqueKey = GenUniqueKey3(KeyLen)
            End Select
            
        Case Is = 6
        If KeyLen = 0 Then KeyLen = Int(1 + Rnd * 10)
            For i = 0 To KeyLen
                GenUniqueKey = GenUniqueKey & Int(1 + Rnd * 2)
            Next i

        Case Is = 7
        If KeyLen = 0 Then KeyLen = Int(1 + Rnd * 10)
            For i = 0 To KeyLen
                GenUniqueKey = GenUniqueKey & 1
            Next i
            
        Case Is = 8
            GenUniqueKey = GenUniqueKeySimple(KeyLen)
            
        Case Else
            GenUniqueKey = GenUniqueKey1(KeyLen)
            
    End Select

If Settings.Jason Then AddUserChat "New Key Generated (Len: " & KeyLen & "): " & GenUniqueKey, "System", True

End Function

Public Function GenUniqueKey1(Optional KeyLen As Integer = 0) As String
'Dim i As Long
Dim DS2 As clsDS2
Set DS2 = New clsDS2
'Dim Key As String
'Dim dblLen As Double
'Dim lnglen As Double
If KeyLen = 0 Then KeyLen = DS2.PRNG(29, 40)
'KeyLen = (KeyLen - 1)

    'dblLen = CDbl(UBound(UniqueKeyCharacters))
    
    Do
        GenUniqueKey1 = GenUniqueKey1 & Chr$(UniqueKeyCharacters(DS2.PRNG(0, 93)))
    Loop While Len(GenUniqueKey1) < (KeyLen)
        
    'If Not VerifyKey(StrConv(GenUniqueKey, vbFromUnicode)) Then MsgBox "Fail!"
    
'If Settings.Jason Then AddChat "[System] New Key1 Generated (Len: " & KeyLen & "): " & GenUniqueKey
End Function

Public Function GenUniqueKeySimple(Optional KeyLen As Integer = 0) As String
Dim i As Long
Dim DS2 As clsDS2
Set DS2 = New clsDS2
If KeyLen = 0 Then KeyLen = DS2.PRNG(4, 12)

    Dim strChars As String
'    strChars = strChars & UCase(strChars) & "0123456789" & "~!@#$%^&*()_+=-`,./<>';"":}{][\|"
    strChars = "abcdefghijklmnopqrstuvwxyz0123456789"
    
    For i = 0 To KeyLen
        GenUniqueKeySimple = GenUniqueKeySimple & Mid$(strChars, DS2.PRNG(1, 36), 1)
    Next
    
End Function

Public Function GenUniqueKey2(Optional KeyLen As Integer = 0) As String
Dim i As Long
'Dim Key As String
Dim DS2 As clsDS2
'Dim intLen As Double
Set DS2 = New clsDS2
If KeyLen = 0 Then KeyLen = DS2.PRNG(29, 40)
'KeyLen = (KeyLen - 1)

    Dim strChars As String
'    strChars = "abcdefghijklmnopqrstuvwxyz"
'    strChars = strChars & UCase(strChars) & "0123456789" & "~!@#$%^&*()_+=-`,./<>';"":}{][\|"
    strChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789~?!@#$%^&*()_+=-`,./<>';"":}{][\|"
'    Debug.Print "[" & strChars & "]"
    
    'intLen = 94
    
    For i = 0 To KeyLen
        GenUniqueKey2 = GenUniqueKey2 & Mid$(strChars, DS2.PRNG(1, 94), 1)
    Next
    
    'If Not VerifyKey(StrConv(GenUniqueKey2, vbFromUnicode)) Then MsgBox "Fail!"
    
'If Settings.Jason Then AddChat "[System] New Key2 Generated (Len: " & KeyLen & "): " & GenUniqueKey2
End Function

Public Function GetScreenX() As Long
Dim Width As Long
GetScreenX = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function GetScreenY() As Long
Dim Width As Long
GetScreenY = GetSystemMetrics(SM_CYSCREEN)
End Function

'below function seems to cause issues on servers, likely because they put the app in a Unicode environment.
'while earlier/non server OSes are not unicode by nature. I've implemented a slower version
'above, but it should be just as effective.

Public Function GenUniqueKey3(Optional KeyLen As Integer = 0) As String

'Dim i As Integer
'Dim Key As String
Dim DS2 As clsDS2
Set DS2 = New clsDS2
If KeyLen = 0 Then KeyLen = DS2.PRNG(29, 40)
'KeyLen = KeyLen - 1

'For i = 0 To KeyLen
'   Key = Key & ChrW$(DS2.PRNG(33, 126))
'Next i

'Debug.Print Key

Do
    GenUniqueKey3 = GenUniqueKey3 & ChrW$(DS2.PRNG(33, 126))
Loop While Len(GenUniqueKey3) < (KeyLen)

'Debug.Print Key

'GenUniqueKey3 = Key$

If Not VerifyKey(StrConv(GenUniqueKey3, vbFromUnicode)) Then MsgBox "Fail!"

'If Settings.Jason Then AddChat "[System] New Key3 Generated (Len: " & KeyLen & "): " & GenUniqueKey3
End Function
