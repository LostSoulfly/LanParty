Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

Private Sub Main()
    Dim iccex As InitCommonControlsExStruct, hMod As Long
    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
    ' feel free to remove any that don't apply to this project
    Const ICC_ANIMATE_CLASS As Long = &H80&
    Const ICC_BAR_CLASSES As Long = &H4&
    Const ICC_COOL_CLASSES As Long = &H400&
    Const ICC_DATE_CLASSES As Long = &H100&
    Const ICC_HOTKEY_CLASS As Long = &H40&
    Const ICC_INTERNET_CLASSES As Long = &H800&
    Const ICC_LINK_CLASS As Long = &H8000&
    Const ICC_LISTVIEW_CLASSES As Long = &H1&
    Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
    Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
    Const ICC_PROGRESS_CLASS As Long = &H20&
    Const ICC_TAB_CLASSES As Long = &H8&
    Const ICC_TREEVIEW_CLASSES As Long = &H2&
    Const ICC_UPDOWN_CLASS As Long = &H10&
    Const ICC_USEREX_CLASSES As Long = &H200&
    Const ICC_STANDARD_CLASSES As Long = &H4000&
    Const ICC_WIN95_CLASSES As Long = &HFF&
    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
       ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
       ' example if using CommonControls v5.0 Progress bar:
       ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
    End With
    On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
    hMod = LoadLibraryA("shell32.dll") ' patch to prevent XP crashes when VB usercontrols present
    InitCommonControlsEx iccex
    If err Then
        InitCommonControls ' try Win9x version
        err.Clear
    End If
    On Error GoTo 0
    '... show your main form next (i.e., frmDebug.Show)
    ' frmDebug.Show
    If hMod Then FreeLibrary hMod

Call ActualMain
'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'          (note). 'bug' may no longer apply with Win7+
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.

End Sub

Sub ActualMain()
Dim i As Integer
    'here we need to make sure that we are running as admin
    'that the winsock dll exists
    'that it is registered
    'and then we can move on to loading the main form.

'todo: remove for final builds
'Settings.blDebug = False

On Error Resume Next

If LCase$(App.EXEName) = "lanparty.new" Then
'MsgBox "lanparty.update"


    
    For i = 0 To 20
        Sleep 500 'pause for a second
            If FileExists(App.Path & "\LanParty.exe") Then
                Kill App.Path & "\LanParty.exe"
                DoEvents
            Else
                Exit For
            End If
    Next

    If i = 20 Then
        MsgBox "Unable to update; the original file cannot be replaced!", vbOKOnly, "Update Failed.."
        End
    End If

    FileCopy App.Path & "\" & App.EXEName & ".exe", App.Path & "\LanParty.exe"
    DoEvents
    Sleep 1000
    Shell App.Path & "\LanParty.exe", vbNormalFocus
    DoEvents
    End
End If

If FileExists(App.Path & "\LanParty.Update.exe") Then
    Sleep 500
    For i = 0 To 20
        Sleep 500 'pause for a second
            If FileExists(App.Path & "\LanParty.New.exe") Then
                Kill App.Path & "\LanParty.New.exe"
                DoEvents
            Else
                Exit For
            End If
    Next
End If

If FileExists(App.Path & "\LanParty.New.exe") Then
    Sleep 500
    For i = 0 To 20
        Sleep 500 'pause for a second
            If FileExists(App.Path & "\LanParty.New.exe") Then
                Kill App.Path & "\LanParty.New.exe"
                DoEvents
            Else
                Exit For
            End If
    Next
End If

On Error GoTo oops

InitializeSettings

If Not isAdmin And Settings.blDebug = False Then
    MsgBox "This program should be run with Admin rights to ensure proper launching/installing of games." & _
    vbNewLine & vbNewLine & "If you have issues, please run this program elevated.", vbCritical, "Stuff won't work right. I guarantee it."
    'End
End If

'build the Winsock resource file and register it

If Not FileExists(App.Path & "\MSWINSCK.OCX") And isAdmin Then
    If MsgBox("I was unable to find some components required to run this application." & vbNewLine & _
    "Lucky for you, I happen to have this component. Shall I extract and register it?", vbYesNo, "MSWINSCK.OCX") = vbYes Then modWinsock.BuildMyResourceFile
End If

If Not FileExists(App.Path & "\comdlg32.OCX") And isAdmin Then
    If MsgBox("I was unable to find some components required to run this application." & vbNewLine & _
    "Lucky for you, I happen to have this component. Shall I extract and register it?", vbYesNo, "comdlg32.OCX") = vbYes Then modComDlg.BuildMyResourceFile
End If

'If Not Settings.blDebug Then modWinsock.BuildMyResourceFile
'If Not Settings.blDebug Then modRichTx32.BuildMyResourceFile
'If Not Settings.blDebug Then modComDlg.BuildMyResourceFile

'init the simple crypto stuff
Set DS2 = New clsDS2
If Not Settings.SameVersion Then
    CryptKey = "e)_o2%$b.Bz+5xjVCcgJ'n-Zw*MIAL7>(|lYscES,&WXv4/NA?{<1tHd3UKfa}=_"
Else
    CryptKey = "e)_o2%$b.Bz+5xjVCcgJ'n-Zw*MIAL7>(|lYscES,&WXv4/NA?{<1tHd3UKfa}=_" & CalculateAdler(LoadFile(App.Path & "\" & App.EXEName & ".exe"))
End If

If Settings.Jason Then AddChat "[System] CryptKey: " & CStr(CryptKey)
'If Settings.Jason Then AddChat "[System] This is the main cryptography key that all packets are encrypted with. The last few characters are generated at runtime " & _
'    "and change with every version of the program, as it is essentially a hash of this EXE itself. This means that each version can only talk to the same version of itself."
    
InitUniqueKeyChars  'initialize the byte array of characters for the improved UniqueKeyGenerator
modVote.InitializeVotes
InitializeUniqueID  'init and set my UID, and display the message upon first open on a new machine
InitMessages        'init the subs that handle the packets received from other clients
Load frmMain        'load the main form and begin running it's form_load
If Not Settings.DisableLan Then LoadUDP     'if we don't have DisableLan enabled, then load up UDP
frmMain.Visible = True
frmMain.Show
If frmChat.Visible = True Then frmChat.Show
If Not Settings.DisableLan Then frmMain.tmrBeacon.Enabled = True
IsSyncingAdmins = True  'since we're just starting, let's sync admins once
HasSyncedAdmins = False 'obviously we haven't connected to anyone yet, so how could we sync admins already?

If Settings.DisableLan Then DisableNetwork  'one final check to make sure everything net related is disabled

If Settings.AutoUpdate Then CheckUpdate
DoEvents

'NewUser "Dragoon", VolumeSerialNumber, "127.0.0.1", "BRADLEY_SURFACE"
'NewUser "Dragoon", VolumeSerialNumber, "127.0.0.1", "BRADLEY_SURFACE"
'NewUser "Dragoon1", VolumeSerialNumber, "127.0.0.1", "BRADLEY_SURFACE"
'NewUser "Dragoon2", VolumeSerialNumber & "1", "127.0.0.1", "BRADLEY_SURFACE"
'NewUser "Dragoon3", VolumeSerialNumber & "2", "127.0.0.1", "BRADLEY_SURFACE"

Exit Sub
oops:

If Settings.blDebug Then AddDebug "Sub Main: " & err.Number & " " & err.Description, True
Resume Next

End Sub

Public Sub FirstRun()
'message boxes etc
MsgBox "The LanParty client facilitate easy LAN gameplay and coordination." & vbNewLine & vbNewLine & _
"An individual that is elected by his peers can become a LanParty Admin, allowing him to suggest commands or games (or run them directly if the option is enabled in settings)." & vbNewLine & vbNewLine, vbInformation, "Welcome!."

End Sub
