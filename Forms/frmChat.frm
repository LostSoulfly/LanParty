VERSION 5.00
Begin VB.Form frmChat 
   Caption         =   "LanChat"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChat 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   5535
   End
   Begin VB.Timer tmrDock 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7080
      Top             =   240
   End
   Begin VB.Timer tmrLog 
      Interval        =   60000
      Left            =   7560
      Top             =   240
   End
   Begin VB.ListBox lstUsers 
      Height          =   2595
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtEnter 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   5535
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Command1_Click()
''On Error Resume Next
'lstvUsers.ListItems.add , "test", "Test!"
'Set lstvUsers.Icons = frmMain.imgList
'lstvUsers.ListItems(1).Icon = 1
'End Sub

'possible todos:
'Allow creation of custom PChatIDs to simplify chats and invites
'Make a right-click menu on frmChat for users that lists their own open chat windows (based on window's PChatIDs)
'Allow chatrooms to change their PChatID (me.tag), but must update each user? Nah, don't do this.
'


Private Sub Form_GotFocus()
    txtEnter.SetFocus
End Sub

Private Sub Form_Load()
    Resize
    AddUserChat "LanParty Client " & GetAppVersion, "System", False
    AddUserChat "Type /help for a list of supported chat commands.", "System", False
End Sub

Public Sub DockChat()
    tmrDock.Enabled = True
    ShowInTheTaskbar Me.hwnd, False
End Sub

Public Sub UnDockChat()
    tmrDock.Enabled = False
    ShowInTheTaskbar Me.hwnd, True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Settings.ChatWindowHeight = Me.Height
    Settings.ChatWindowWidth = Me.Width
    
    UpdateControlPositions
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
UnDockChat
Settings.DockChat = False
Settings.ShowChat = False

RefreshSettings

'Me.Visible = False

Cancel = 1

End Sub

Private Sub lstUsers_Click()
    UpdateUserCurrentStatus
End Sub

Private Sub lstUsers_DblClick()
On Error Resume Next
Dim strKey As String
Dim UserIndex As Integer
UserIndex = GetUserIndexFromChat
If UserIndex = -1 Then Exit Sub
    strKey = GenUniqueKeySimple(6) 'gen a new chat ID
    CreatePChatWindow strKey   'Create the new chat window with the ID, then invite the remote user.
    GetPChatWindow(strKey).AddChatUser User(UserIndex).UniqueID
    AddUserPrivateChat "Sending invite to " & GetUserNameByIndex(UserIndex) & ", please wait..", "System", strKey
    SendCryptTo UserIndex, PrivateChatPacket(1, 0, User(UserIndex).UniqueID, strKey, "")
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu frmMain.mnuUser
End Sub

Private Sub tmrDock_Timer()
On Error Resume Next
    If Settings.DockChat = False Then UnDockChat
If Me.Visible = False And frmMain.WindowState <> vbNormal Then Exit Sub
If frmMain.WindowState <> vbNormal Then
    If Me.Visible = True Then Me.Visible = False
    Exit Sub
End If
    If Me.Visible = False Then Me.Visible = True
    Me.Top = frmMain.Height + frmMain.Top
    Me.Left = frmMain.Left
    DoEvents
End Sub

Public Sub Resize()
On Error Resume Next
If Settings.ChatWindowHeight = 0 Then Exit Sub
If Settings.ChatWindowWidth = 0 Then Exit Sub
    Me.Width = Settings.ChatWindowWidth
    Me.Height = Settings.ChatWindowHeight
    UpdateControlPositions
End Sub

Private Sub UpdateControlPositions()
On Error Resume Next
    txtChat.Left = 120
    txtChat.Top = txtChat.Left
    txtChat.Width = frmChat.Width / 1.4
    lstUsers.Left = txtChat.Width + (txtChat.Left * 1.5)
    lstUsers.Width = (frmChat.Width - txtChat.Width - (txtChat.Left * 3)) - 150
    txtChat.Height = frmChat.Height - txtEnter.Height - txtChat.Top - 600 - 100
    txtEnter.Top = txtChat.Height + txtEnter.Height + txtChat.Top - 250
    lstUsers.Height = txtEnter.Top + txtEnter.Height - 50
    txtEnter.Width = txtChat.Width

End Sub

Private Sub tmrLog_Timer()

    If Len(txtChat.Text) > 150000 Then

    End If
    
End Sub

Private Sub txtChat_Change()
On Error Resume Next
Dim ChatLen As Long
ChatLen = Len(txtChat.Text)

If ChatLen > 65000 Then
        If Settings.LogChat Then WriteLog txtChat.Text, FixFilePath(App.Path & "\Chat.log")
        DoEvents
        txtChat.Text = vbNullString
End If
    txtChat.SelStart = Len(txtChat.Text) - 1
    
End Sub

Private Sub txtChat_Click()
On Error Resume Next
    txtEnter.SetFocus
End Sub

Private Sub txtEnter_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Dim Text As String
    Dim i As Integer
    Text = Trim$(txtEnter.Text)
    If Len(Text) > 45000 Then txtEnter.Text = ""
    
    If Left$(Text, 1) = "/" Then
        'entering a command.
    Dim strCommand As String
    Dim lSpace As Long
    Dim strData As String
    
    lSpace = InStr(1, Text, " ")
    If lSpace > 0 Then
        strCommand = Mid(Text, 2, lSpace - 2)
        strData = Right(Text, Len(Text) - lSpace)
    Else
        strCommand = Mid(Text, 2, Len(Text))
    End If
    
    Select Case LCase$(strCommand)
    
    Case "help"
        AddChat "[System] Help commands: "
        AddChat "[System] /me <text> to talk in third person. "
        AddChat "[System] /exit, /quit to.. exit or quit the LanParty"
        AddChat "[System] /update to check for updates"
        AddChat "[System] /savegames to save the games list"
        AddChat "[System] /vote to quickly show the newest vote."
        AddChat "[System] /votelist, vl to open the list of current votes."
        AddChat "[System] /launch <GameName> wildcards assumed before and after GameName"
        AddChat "[System] /launch2 <GameName> to quickly launch the first game found, wildcards(*) supported"
                
    Case "me"
        'third person text?
        AddChat "*" & GetUserName(Settings.UniqueID) & " " & strData
        CryptToAll ChatPacket(strData, 1)
        
    Case "pm", "whisper"
        AddUserChat "Meh. Just double click on their name to open a window.", "System", False
    
    Case "killchat", "killchats", "closechats", "closeallchats"
        RemoveAllPChatWindows
        
    Case "exit", "quit"
        frmMain.DoExit
    
    Case "update"
        CheckUpdate
        
    Case "savegames", "savegame"
        AddUserChat "Saving games list..", "System", False
        SaveGames
    
    Case "vote"
        AddUserChat "Showing newest vote..", "System", False
        If UBound(Vote) = 0 Then txtEnter.Text = "": Exit Sub
        ShowVoteWindow UBound(Vote)
        
    Case "votelist"
        AddUserChat "Showing vote window..", "System", False
        frmVoteList.Visible = True
        frmVoteList.RefreshList
    
    Case "launch2"
        If Len(strData$) < 1 Then AddChat "Must specify something else to run this command!": txtEnter.Text = "": Exit Sub
        strData = LCase$(strData)
       AddUserChat "Searching games for pattern: " & strData, "System", False

        For i = 1 To UBound(Game)
            
            If LCase$(Game(i).Name) Like strData Then
                AddDebug strData & " like: " & Game(i).Name
                LaunchGame i, True
            End If
            'If InStr(1, LCase$(Game(i).Name), strData) > 0 Then
            
            'End If
        Next i
    
    Case "launch"
    If Len(strData$) < 1 Then AddUserChat "If I actually did that, it would launch every game. EVERY GAME.", "System", False: Exit Sub
        strData = "*" & LCase$(strData) & "*"
        AddUserChat "Searching games for pattern: " & strData, "System", False

        For i = 1 To UBound(Game)
            
            If LCase$(Game(i).Name) Like strData Then
                AddDebug strData & " like: " & Game(i).Name
                LaunchGame i, True
            End If
            'If InStr(1, LCase$(Game(i).Name), strData) > 0 Then
            
            'End If
        Next i
    
    
    Case Else
        AddUserChat "Unknown command: " & strCommand, "System", False
    
    End Select
    
    txtEnter.Text = ""
    Exit Sub
    End If
    
    If Settings.AltChatType Then
        BroadcastUDP ChatPacket(Text)
    Else
        'SendToAll ChatPacket(Trim$(txtEnter.Text))
        'DoEvents
        'SendCryptChatAll Text
        CryptToAll ChatPacket(Text)
    End If
    AddUserChat Text, Settings.UserName
    txtEnter.Text = ""
End If

End Sub
