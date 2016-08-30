VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "LanParty"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar vsScroll 
      Height          =   6495
      Left            =   7920
      Max             =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   6700
      Left            =   0
      ScaleHeight     =   6705
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      Begin VB.Timer tmrVotesSync 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   6960
         Top             =   1440
      End
      Begin VB.Timer tmrMonitorGame 
         Interval        =   3000
         Left            =   6960
         Top             =   1920
      End
      Begin VB.Timer tmrAdmins 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   6960
         Top             =   960
      End
      Begin VB.Timer tmrBeacon 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   6960
         Top             =   480
      End
      Begin VB.Timer tmrPing 
         Interval        =   60000
         Left            =   6960
         Top             =   0
      End
      Begin MSWinsockLib.Winsock sckListen 
         Left            =   7440
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock sckBroadcast 
         Left            =   7440
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.Timer tmrClose 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7440
         Top             =   0
      End
      Begin VB.Timer tmrIconSize 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7440
         Top             =   480
      End
      Begin VB.Image imgIcon 
         Height          =   1695
         Index           =   0
         Left            =   240
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblIcon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<Game Name>"
         Height          =   555
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Check For Updates"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuLP 
      Caption         =   "Lan Party"
      Begin VB.Menu mnuManual 
         Caption         =   "Manual Connection (By IP/Hostname)"
      End
      Begin VB.Menu mnuNewVote 
         Caption         =   "Open Vote Window"
      End
      Begin VB.Menu mnuMonitorGame 
         Caption         =   "Monitor Current Game"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLanChat 
         Caption         =   "Show LanChat"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDockChat 
         Caption         =   "Dock LanChat"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Icon Size"
         Begin VB.Menu mnuHuge 
            Caption         =   "HUGE ICONS"
         End
         Begin VB.Menu mnuLarge 
            Caption         =   "Large Icons"
         End
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal Icons"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSmallIcons 
            Caption         =   "Small Icons"
         End
         Begin VB.Menu mnuList 
            Caption         =   "List Icons"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuAdv 
      Caption         =   "Advanced"
      Begin VB.Menu mnuAllowCommands 
         Caption         =   "Allow LanAdmin To Run Commands"
      End
      Begin VB.Menu mnuSuggestAll 
         Caption         =   "Suggest a Command To All"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuScanGame 
         Caption         =   "Scan For Running Game"
      End
      Begin VB.Menu mnuSaveGames 
         Caption         =   "Save Games File"
      End
      Begin VB.Menu mnuEditGameFile 
         Caption         =   "Edit GameInfo File"
      End
      Begin VB.Menu mnuApplyAdmin 
         Caption         =   "Apply For Admin Status"
      End
      Begin VB.Menu mnuViewAdmins 
         Caption         =   "View Admin Status"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGameEditor 
         Caption         =   "Game Editor"
      End
      Begin VB.Menu mnuFirewall 
         Caption         =   "Windows Firewall"
         Begin VB.Menu mnuEnableFirewall 
            Caption         =   "Enable Firewall"
         End
         Begin VB.Menu mnuDisableFirewall 
            Caption         =   "Disable Firewall"
         End
      End
      Begin VB.Menu mnuCryptView 
         Caption         =   "View Debug Info"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "[Selected User]"
      Begin VB.Menu mnuStartPChat 
         Caption         =   "Start New Private Chat"
      End
      Begin VB.Menu mnuMsg 
         Caption         =   "Send MessageBox"
         Begin VB.Menu mnuMsgThisUser 
            Caption         =   "This User"
         End
         Begin VB.Menu mnuMsgAllUsers 
            Caption         =   "All Users"
         End
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "Suggest A Command/Execute CMD"
         Begin VB.Menu mnuExecThisUser 
            Caption         =   "This User"
         End
         Begin VB.Menu mnuExecAll 
            Caption         =   "All Users"
         End
      End
      Begin VB.Menu mnuSuggest 
         Caption         =   "Suggest a Command"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute User Locally"
      End
      Begin VB.Menu mnuGlobalMuteMnu 
         Caption         =   "Global Mute"
         Begin VB.Menu mnuGlobalMute 
            Caption         =   "Mute"
         End
         Begin VB.Menu mnuGlobalUnMute 
            Caption         =   "Unmute"
         End
      End
      Begin VB.Menu mnuChangeUserName 
         Caption         =   "Change User Name"
      End
      Begin VB.Menu mnuFreeze 
         Caption         =   "Freeze User"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick User"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game Menu"
      Begin VB.Menu mnuLaunch 
         Caption         =   "Launch Game"
      End
      Begin VB.Menu mnuLaunchNoArgs 
         Caption         =   "Launch (No CMD Args)"
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Open Game Folder"
      End
      Begin VB.Menu mnuLocateGame 
         Caption         =   "Locate Game"
      End
      Begin VB.Menu mnuInstallGame 
         Caption         =   "Install Game"
      End
      Begin VB.Menu mnuLaunchAll 
         Caption         =   "Suggest Launch To All"
      End
      Begin VB.Menu mnuEditGame 
         Caption         =   "Edit Game"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Game"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddGame 
         Caption         =   "Add Game"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuResources 
      Caption         =   "Network Resources"
      Begin VB.Menu mnuResource 
         Caption         =   "<Resource1>"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private arrResources() As String

Private Sub Form_Load()

'load and apply settings from the udt
RefreshSettings
Me.Height = Settings.MainWindowHeight
Me.Width = Settings.MainWindowWidth
Me.Visible = True
Me.Caption = "Loading, please wait.."

'Load all games from file
InitializeGameArray
InitializeNetworkMenu

'check that games that are required to be installed are found
CheckInstalled

DoEvents
mnuUser.Visible = False
'todo: enable for builds
'CRASHES on breakpoints often
WheelHook Me.hwnd

InitializeUsers
Me.Caption = "LanParty Launcher v" & App.Major & "." & App.Minor & "." & App.Revision
'mnuApplyAdmin_Click
RefreshBackground
SelectGame 1
AddUserChat "Please wait while I locate other users..", "System", False
UpdateAdminMenus False
'Start broadcasting our existence.

End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
On Error Resume Next

If Rotation > 0 Then

    If vsScroll.Value = vsScroll.Min Then Exit Sub
    vsScroll.Value = vsScroll.Value - 1

Else

    If vsScroll.Value = vsScroll.Max Then Exit Sub
    vsScroll.Value = vsScroll.Value + 1

End If
  
  
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
If Me.Visible = False Then Exit Sub
vsScroll.Left = frmMain.Width - vsScroll.Width * 1.4
vsScroll.Height = frmMain.Height - 1300
'picContainer.Width = frmMain.Width - vsScroll.Width
'picContainer.Height = frmMain.Height + 400
picContainer.Left = 0
picContainer.Top = 0
DoEvents
tmrIconSize.Enabled = True
Settings.MainWindowHeight = Me.Height
Settings.MainWindowWidth = Me.Width

If frmChat.Visible = True Then
    Settings.ChatWindowHeight = frmChat.Height
    If Settings.DockChat Then Settings.ChatWindowWidth = Me.Width
    frmChat.Resize
End If
'picContainer.Picture = Resize(picContainer.Picture.Handle, picContainer.Picture.Type, frmMain.Width, frmMain.Height, , True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
DoEvents
tmrClose.Enabled = True
End Sub

Private Sub mnuAbout_Click()
MsgBox "This has been in development since early December, and is comprised of over 8000 lines of code, the majority of it written by me." & vbNewLine & _
"I wrote this initially for the TrollParty 2016 LAN, in an attempt to make it as easy as possible for people to just come over and play games." & vbNewLine & vbNewLine & _
"I hope this program will make things easier for you personally and enable you to spend more time playing games and having fun." & vbNewLine & "-Dragoon (Bradley)", vbInformation, "About"
End Sub

Private Sub mnuAllowCommands_Click()
    mnuAllowCommands.Checked = Not mnuAllowCommands.Checked
    Settings.AllowCommands = mnuAllowCommands.Checked
End Sub

Private Sub mnuApplyAdmin_Click()
Dim VoteID As String

VoteID = modVote.NewAdminVote(GenUniqueKey(21), Settings.UniqueID)

If Len(VoteID$) < 19 Then Exit Sub

CryptToAll NewAdminVotePacket(VoteID, Settings.UniqueID)

ShowVoteWindow VoteExists(VoteID)

AddUserChat "Starting new LanAdmin vote..", "System", False

End Sub

Private Sub mnuChangeUserName_Click()
Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub
Dim UserName As String

UserName = InputBox("What is the new UserName you'd like to give this person?", "New UserName", GetUserNameByIndex(UserIndex))

If Len(UserName$) < 3 Then Exit Sub

    CryptToAll LanAdminPacket(3, User(UserIndex).UniqueID, UserName)
    
End Sub

Private Sub mnuCryptView_Click()
Load frmDebug
frmDebug.Visible = True
End Sub

Private Sub mnuDisableFirewall_Click()
If Not isAdmin Then MsgBox "You must be running LanParty as an elevated user (As an Admin) in order to do this.", vbCritical, "Baka.": Exit Sub

If GetOSVer = 1 Or GetOSVer = 6 Then
    ExecFile "netsh", "firewall set opmode disable", , , 0
Else
    ExecFile "netsh", "advfirewall set allprofiles state off", , , 0
End If

End Sub

Private Sub mnuDockChat_Click()
    
If Not mnuLanChat.Checked Then mnuLanChat_Click
    
    mnuDockChat.Checked = Not mnuDockChat.Checked
    
    If mnuDockChat.Checked Then If Not mnuLanChat.Checked Then mnuLanChat_Click
    
    Settings.DockChat = mnuDockChat.Checked
    If mnuDockChat.Checked Then frmChat.DockChat Else frmChat.UnDockChat

End Sub

Private Sub mnuEditGame_Click()
frmGameEdit.Visible = True
frmGameEdit.RefreshGame CurrentGameIndex
End Sub

Private Sub mnuEditGameFile_Click()
ExecFile GameFile, ""
End Sub

Private Sub mnuEnableFirewall_Click()
If Not isAdmin Then MsgBox "You must be running LanParty as an elevated user (As an Admin) in order to do this.", vbCritical, "Baka.": Exit Sub


If GetOSVer = 1 Or GetOSVer = 6 Then
    ExecFile "netsh", "firewall set opmode enable", , , 0
Else
    ExecFile "netsh", "advfirewall set allprofiles state on", , , 0
End If

End Sub

Private Sub mnuExecAll_Click()
Dim Command As String
Dim Args As String
Dim blShell As Boolean

Command = InputBox("What command would you execute?", "Suggest A Command")
If Len(Command$) = 0 Then Exit Sub
Args = InputBox("What are the command arguments for this?", "Exe Args")
If Len(Args$) = 0 Then Exit Sub
If MsgBox("Should this be Shelled? (As opposed to ShellExecute)", vbYesNo, "Shell the command?") = vbYes Then
    blShell = True
End If
    
    If MsgBox("Here's the rundown." & vbNewLine & "Full command: " & Command & " " & Args & vbNewLine & " Shell: " & blShell & vbNewLine & "Are you sure you want to send this command?", vbYesNo, "Confirmation") = vbYes Then
        CryptToAll LanAdminExecPacket(4, Command, Args, blShell)
        'SendCryptTo GetUserIndexFromChat, LanAdminExecPacket(4, Command, Args, blShell)
    End If

End Sub

Private Sub mnuExecThisUser_Click()
Dim Command As String
Dim Args As String
Dim blShell As Boolean

Command = InputBox("What command would you execute?", "Suggest A Command")
If Len(Command$) = 0 Then Exit Sub
Args = InputBox("What are the command arguments for this?", "Exe Args")
'If Len(Args$) = 0 Then Exit Sub
If MsgBox("Should this be Shelled? (As opposed to ShellExecute)", vbYesNo, "Shell the command?") = vbYes Then
    blShell = True
End If
    
    If MsgBox("Here's the rundown." & vbNewLine & "Full command: " & Command & " " & Args & vbNewLine & " Shell: " & blShell & vbNewLine & "Are you sure you want to send this command?", vbYesNo, "Confirmation") = vbYes Then
        'CryptToAll LanAdminExecPacket(4, Command, Args, blShell)
        SendCryptTo GetUserIndexFromChat, LanAdminExecPacket(4, Command, Args, blShell)
    End If

End Sub

Private Sub mnuExit_Click()
    'this crashes if you do it in the same sub as End for some reason..
    'todo: reenable for builds.
    DoExit
End Sub

Public Sub DoExit()

    If Settings.LogChat Then If frmChat.Visible Then WriteLog frmChat.txtChat.Text, App.Path & "\Chat.log"
    DoEvents

    WheelUnHook Me.hwnd
    tmrClose.Enabled = True
End Sub

Private Sub mnuFreeze_Click()

Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub

    CryptToAll LanAdminPacket(6, User(UserIndex).UniqueID)
End Sub

Private Sub mnuGameEditor_Click()
    frmGameEdit.Visible = True
End Sub

Private Sub mnuGlobalMute_Click()
Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub
    MuteUser UserIndex, True
    CryptToAll LanAdminPacket(0, User(UserIndex).UniqueID)
End Sub

Private Sub mnuGlobalUnMute_Click()
Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub
    
    MuteUser UserIndex, False
    CryptToAll LanAdminPacket(1, User(UserIndex).UniqueID)
End Sub

Private Sub mnuInstallGame_Click()
    InstallGame CurrentGameIndex
End Sub

Private Sub mnuKick_Click()
Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub

    CryptToAll LanAdminPacket(2, User(UserIndex).UniqueID)
End Sub

Private Sub mnuLanChat_Click()

    If Settings.DockChat Then mnuDockChat_Click: DoEvents

    mnuLanChat.Checked = Not mnuLanChat.Checked
    Settings.ShowChat = mnuLanChat.Checked
    frmChat.Visible = Settings.ShowChat
    
End Sub

Private Sub mnuHuge_Click()
SetIconsTo 4
End Sub

Public Sub SetIconsTo(size As Integer)
UncheckAllSizes

    Select Case size
    
    Case Is = 1
        mnuSmallIcons.Checked = True
        IconWidth = 600
        IconHeight = 600
    Case Is = 2
        mnuNormal.Checked = True
        IconWidth = 900
        IconHeight = 900
    Case Is = 3
        mnuLarge.Checked = True
        IconHeight = 1200
        IconWidth = 1200
    Case Is = 4
        mnuHuge.Checked = True
        IconHeight = 1600
        IconWidth = 1600
    Case Is = 5
        IconHeight = 2000
        IconWidth = 2000
    Case Is = 6
        IconHeight = 2400
        IconWidth = 2400
    Case Is = 7
        IconHeight = 2800
        IconWidth = 2800
    Case Is = 8
        IconHeight = 3200
        IconWidth = 3200
    Case Is = 9
        IconHeight = 3600
        IconWidth = 3600
    End Select
    
    Settings.IconSize = size
    
    UpdateIconList True
    Call tmrIconSize_Timer
End Sub

Private Sub mnuLarge_Click()
SetIconsTo 3
End Sub

Private Sub mnuLaunch_Click()
    LaunchGame CurrentGameIndex
End Sub

Private Sub mnuLaunchAll_Click()

If CurrentGameIndex = -1 Then Exit Sub
If Len(Game(CurrentGameIndex).GameUID$) > 0 Then CryptToAll LanAdminPacket(5, Game(CurrentGameIndex).GameUID)
End Sub

Private Sub mnuLaunchNoArgs_Click()
    LaunchGame CurrentGameIndex, False
End Sub

Private Sub mnuList_Click()
UncheckAllSizes

mnuList.Checked = True
UpdateIconList True
End Sub

Private Sub mnuLocateGame_Click()
    LocateGame CurrentGameIndex
End Sub

Private Sub mnuManual_Click()
Dim IP As String

IP = InputBox("What IP/Hostname would you like to attempt to connect to?", "Manual Connection")

If Len(IP$) = 0 Then Exit Sub

SendDataToUDP IP, AuthPacket(0)

End Sub

Private Sub mnuMonitorGame_Click()
    mnuMonitorGame.Checked = Not mnuMonitorGame.Checked
    Settings.MonitorGame = mnuMonitorGame.Checked
End Sub

Private Sub mnuMsgAllUsers_Click()
Dim Message As String
Dim Title As String

Message = InputBox("What message would you send?", "Send MsgBox")
If Len(Message) = 0 Then Exit Sub
Title = InputBox("What is the title for the message box?", "Message Title")
If Len(Title) = 0 Then Exit Sub
    
    If MsgBox(Message & vbNewLine & "Are you sure you want to send this command?", vbYesNo, Title) = vbYes Then
        'CryptToAll LanAdminExecPacket(4, Command, Args, blShell)
        CryptToAll LanAdminExecPacket(7, Message, Title, False)
    End If

End Sub

Private Sub mnuMsgThisUser_Click()
Dim Message As String
Dim Title As String

Message = InputBox("What message would you send?", "Send MsgBox")
If Len(Message) = 0 Then Exit Sub
Title = InputBox("What is the title for the message box?", "Message Title")
If Len(Title) = 0 Then Exit Sub
    
    If MsgBox(Message & vbNewLine & "Are you sure you want to send this command?", vbYesNo, Title) = vbYes Then
        'CryptToAll LanAdminExecPacket(4, Command, Args, blShell)
        SendCryptTo GetUserIndexFromChat, LanAdminExecPacket(7, Message, Title, False)
    End If

End Sub

Private Sub mnuMute_Click()

Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub

    User(UserIndex).Muted = Not User(UserIndex).Muted
    Call UpdateMenu(UserIndex)
        
End Sub

Private Sub mnuNewVote_Click()
frmVoteList.Visible = True
frmVoteList.RefreshList
End Sub

Private Sub mnuNormal_Click()
SetIconsTo 2
End Sub

Private Sub mnuOpenFolder_Click()
Dim dirPath As String

If Len(Game(CurrentGameIndex).EXEPath) = 0 Then
    If Game(CurrentGameIndex).GameType = 1 Then
        'because assuming is easy.
        'I could loop through the PATH system variable but
        'I doubt anyone would ever use it in that way
        AddUserChat "I'm assuming the System32 folder because I'm lazy.", "System", False
        dirPath = Environ("WINDIR") & "\System32\"
    Else
        Exit Sub
    End If
End If

If Left(Game(CurrentGameIndex).EXEPath, 1) = "\" Then
    dirPath = FullPathFromLocal(Game(CurrentGameIndex).EXEPath)
End If

If DirExists(dirPath) Then
    
    ExecFile dirPath, ""
    
Else
    
    AddUserChat "Unable to locate Game directory.", "System", False
    
End If
    
End Sub

Private Sub mnuResource_Click(Index As Integer)

    If Index = 0 Then
        ExecFile "http://trollparty.org", ""
    Else
        ExecFile arrResources(Index), ""
    End If

End Sub

Private Sub mnuSaveGames_Click()
    SaveGames
End Sub

Private Sub mnuScanGame_Click()
If Settings.MonitorGame = False Then
    If MsgBox("Game Monitoring is disabled. Enable it?", vbYesNo, "Enable Monitoring") = vbYes Then
        Settings.MonitorGame = True
    Else
        Exit Sub
    End If
End If

    Dim i As Integer
    
    For i = 1 To UBound(Game)
    
        If IsGameRunning(i) Then CheckMonitorGame (i): Exit Sub
    
    Next i

End Sub

Private Sub mnuSettings_Click()

frmSettings.Show vbModal
    
If Not frmSettings.blCancel Then
    Unload frmSettings
    'InitializeSettings
    RefreshSettings
    RefreshBackground
Else
    Unload frmSettings

End If

End Sub

Public Sub UncheckAllSizes()
    mnuLarge.Checked = False
    mnuHuge.Checked = False
    mnuNormal.Checked = False
    mnuList.Checked = False
    mnuSmallIcons.Checked = False
End Sub

Private Sub imgIcon_Click(Index As Integer)
    SelectGame Index
    SetCaption Game(Index).Name
End Sub

Private Sub imgIcon_DblClick(Index As Integer)
    mnuLaunch_Click
End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    'right click
    SelectGame Index
    mnuLaunch.Caption = "Launch " & Game(Index).Name
    mnuEditGame.Caption = "Edit " & Game(Index).Name
    PopupMenu mnuGame
End If

End Sub

Private Sub mnuSmallIcons_Click()
UncheckAllSizes
mnuSmallIcons.Checked = True
Settings.IconSize = 1
IconWidth = 600
IconHeight = 600
UpdateIconList True
End Sub

Private Sub mnuStartPChat_Click()
Dim strKey As String
Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub
    strKey = InputBox("What would you like to call this private chat?", "Private Chat ID", GenUniqueKeySimple(6)) 'GenUniqueKey(21) 'gen a new chat ID
    CreatePChatWindow strKey  'Create the new chat window with the ID, then invite the remote user.
    GetPChatWindow(strKey).AddChatUser User(UserIndex).UniqueID
    AddUserPrivateChat "Sending invite to " & GetUserNameByIndex(UserIndex) & ", please wait..", "System", strKey
    SendCryptTo UserIndex, PrivateChatPacket(1, 0, User(UserIndex).UniqueID, strKey, "")
    
End Sub

Private Sub mnuSuggest_Click()
Dim Command As String
Command = InputBox("Wut?", "Suggest")
    If Len(Command) > 0 Then
        CryptToAll LanAdminPacket(4, Command)
    End If
End Sub

Private Sub mnuSuggestAll_Click()

Dim Command As String
Command = InputBox("Wut?", "Suggest")
    If Len(Command) > 0 Then
        CryptToAll LanAdminPacket(4, Command)
    End If
    
End Sub

Private Sub mnuUpdates_Click()
    CheckUpdate
End Sub

Private Sub sckBroadcast_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddDebug "sckListen: " & Number & " " & Description, True
End Sub

Private Sub sckListen_DataArrival(ByVal bytesTotal As Long)
    IncomingDataUDP (bytesTotal)
End Sub

Private Sub sckListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddDebug "sckListen: " & Number & " " & Description, True
End Sub

Private Sub tmrAdmins_Timer()
If GetUserCount <= 0 Then Exit Sub
    
If IsSyncingAdmins = True Or HasSyncedAdmins = False Then
    CryptToAllAdminSync ReqAdminSyncPacket
Else
    tmrAdmins.Enabled = False
End If

CalculateAdminLists

End Sub

Private Sub tmrBeacon_Timer()
On Error Resume Next
If LenB(Settings.UserName$) = 0 Then Exit Sub
BroadcastUDP BeaconPacket
'AddDebug "Sending beacon.."

'If Settings.MonitorGame Then
'    Dim i As Integer
'    For i = 1 To UBound(Game)
'        If IsGameRunning(i) Then CheckMonitorGame (i): Exit Sub
'    Next i
'End If

tmrBeacon.Interval = DS2.PRNG(1, 10) * 1000
End Sub

'Private Sub Timer1_Timer()
'Me.Caption = "vscroll: " & vsScroll.Value & " max: " & vsScroll.Max & " Cont H: " & picContainer.Height & " T:" & picContainer.Top
'End Sub

Private Sub tmrClose_Timer()
    'If Settings.blDebug = False Then WheelUnHook Me.hwnd
    On Error Resume Next
    WheelUnHook Me.hwnd
    CryptToAll GoodbyePacket
    SaveSettings
    SaveGames
    End
End Sub

Private Sub tmrIconSize_Timer()
On Error GoTo Escape

If ((Not Game) = -1) Then Exit Sub 'catch a crash

If (UBound(Game) = 0) And (LenB(Game(0).Name$) = 0) Then Exit Sub

    UpdateIconList True
    If frmMain.picContainer.Height < frmMain.Height Then frmMain.picContainer.Height = frmMain.Height + 200
    If frmMain.picContainer.Width < frmMain.Width Then frmMain.picContainer.Width = frmMain.Width + 200
    
    Dim NumOfRows As Integer
    Dim NumRowsVisible As Integer
    Dim MaxLocationIcon As Long
    
    NumOfRows = (UBound(Game) / lngLastIconsPerRow)
    'round up..
    If (UBound(Game) / lngLastIconsPerRow) > NumOfRows Then NumOfRows = NumOfRows + 1
    
    'biggest height the box needs to be
    MaxLocationIcon = NumOfRows * (imgIcon(0).Top + imgIcon(0).Height + lblIcon(0).Height)
    
    'num rows visi
    
    'If we have rows drawn beneath the height of the window, we need to be able to scroll.
    If MaxLocationIcon > frmMain.Height Then
    
        Dim i As Integer
        
        For i = 1 To NumOfRows
            If i * (imgIcon(0).Top + imgIcon(0).Height + lblIcon(0).Height) >= frmMain.Height Then
                NumRowsVisible = i - 1
                Exit For
            End If
        Next i
    
        'AddDebug "All rows NOT visible. Rows not visible: " & (NumOfRows - NumRowsVisible)
        'vsScroll.Max = (MaxLocationIcon - frmMain.Height) / NumOfRows
        vsScroll.Max = (NumOfRows - NumRowsVisible)
        lngScrollAmt = (MaxLocationIcon / NumOfRows)
        vsScroll.Visible = True
    Else
    'all rows are visible
        'AddDebug "All rows visible."
        lngScrollAmt = 0
        vsScroll.Max = 0
        vsScroll.Visible = False
    End If
    
    vsScroll.Value = vsScroll.Min
    tmrIconSize.Enabled = False
    
Exit Sub
Escape:

If err.Number = 9 Then err.Clear: Exit Sub
If err.Number = 360 Then Resume Next
If err.Number = 365 Then Resume Next
'Resume Next
AddDebug "tmrIconSize " & err.Number & ": " & err.Description

End Sub



Private Sub tmrMonitorGame_Timer()
    CheckMonitorGame
End Sub

Private Sub tmrPing_Timer()
'loop through computers and increment their lastseen integer
'if it's over, say 1 minutes, send a ping packet.
'if it's over 2 minutes, remove them.
'The beacon packet should be sent every 5 seconds. That will keep them online for eachother.
On Error Resume Next
Dim i As Integer
For i = 1 To UBound(User)
    With User(i)
        If LenB(.UniqueID) > 0 Then
            'they are real!
                Select Case .LastHeard
                
                Case Is = 0
                
                Case Is = 1
                    SendDataToUDP .IP, PingPacket
                Case Is = 2
                    SendDataToUDP .IP, PingPacket
                End Select
            
                'If .LastHeard = 1 Then
                '    'haven't heard from them in 2 mins.
                '    'In one more minute they're removed.
                '    .LastHeard = 2
                'End If
            
            .LastHeard = .LastHeard + 1
            If .LastHeard = 3 Then RemoveUser (i) 'Remove the user after a few minutes
        End If
    End With
Next i

End Sub

Private Sub tmrVotesSync_Timer()
    SyncAllVotes
End Sub

Private Sub vsScroll_Change()
On Error Resume Next
    picContainer.Top = 0 - (vsScroll.Value * lngScrollAmt)
    'picContainer.Top = 0 - (vsScroll.Value * (Fix(picContainer.Height / 5) / vsScroll.Max))
End Sub

Public Sub RefreshBackground()
On Error Resume Next
    If LenB(Settings.BackgroundPath) > 0 Then picContainer.Picture = LoadPicturePlus(FixFilePath(Settings.BackgroundPath))
    'picContainer.Picture = Resize(picContainer.Picture.Handle, vbPicTypeBitmap, 600, 600, , True)
End Sub

Private Sub ResetGameSelect()
On Error Resume Next
Dim i As Integer
On Error Resume Next
    For i = 1 To UBound(Game)
        imgIcon(i).BorderStyle = 0
        'lblIcon(i).BackStyle = 0
        'lblIcon(i).BorderStyle = 0
        'lblIcon(i).FontBold = False
        'lblIcon(i).BackColor = &H8000000F
        'lblIcon(i).ForeColor = &H80000012
    Next

End Sub

Public Sub SelectGame(GameIndex As Integer)
On Error Resume Next

ResetGameSelect

'AddDebug "Last index: " & CurrentGameIndex
'AddDebug "New index: " & GameIndex

    imgIcon(GameIndex).BorderStyle = 1
    'lblIcon(GameIndex).BackStyle = 1
    'lblIcon(GameIndex).BorderStyle = 1
    'lblIcon(GameIndex).FontBold = True
    'lblIcon(GameIndex).ForeColor = vbGreen

CurrentGameIndex = GameIndex
mnuInstallGame.Visible = (Game(GameIndex).InstallFirst And (Not Game(GameIndex).Installed))
End Sub

Public Sub UpdateMenu(UserIndex As Integer)

    mnuMute.Checked = User(UserIndex).Muted

End Sub

Public Sub InitializeNetworkMenu()
Dim strResources() As String
On Error Resume Next
mnuResource(0).Caption = "TrollParty Website"

If Not FileExists(App.Path & "\Resources.txt") Then mnuResources.Visible = False: Exit Sub
If FileLen(App.Path & "\Resources.txt") <= 10 Then mnuResources.Visible = False: Exit Sub
strResources = Split(LoadFile(App.Path & "\Resources.txt"), vbNewLine)

If UBound(strResources) < 1 Then mnuResources.Visible = False: Exit Sub

If (UBound(strResources) + 1) Mod 2 <> 0 Then AddUserChat "Incorrect number of lines in Resources.txt!", "System", False: Exit Sub

'mnuResource(0).Visible = False

'ReDim mnuResource(0)
ReDim arrResources(0)

Dim i As Integer



For i = mnuResource.Count - 1 To 1 Step -1
    Unload mnuResource(i)
Next i

'Load mnuResource(1)

For i = 0 To UBound(strResources) Step 2

    ReDim Preserve arrResources(mnuResource.Count)
    'ReDim Preserve mnuResource(UBound(mnuResource) + 5)
    Load mnuResource(mnuResource.Count)
    mnuResource(mnuResource.Count - 1).Caption = strResources(i)
    arrResources(UBound(arrResources)) = strResources(i + 1)
    
Next


End Sub
