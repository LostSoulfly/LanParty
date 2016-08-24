VERSION 5.00
Begin VB.Form frmPrivateChat 
   Caption         =   "PrivateChat"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8745
   Icon            =   "frmPrivateChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUsers 
      Height          =   2595
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame frameInvite 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdDecline 
         Caption         =   "Decline"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblInvite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   6375
      End
   End
   Begin VB.TextBox txtChat 
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
   End
   Begin VB.TextBox txtEnter 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Menu mnuPrivate 
      Caption         =   "Private Chat"
      Begin VB.Menu mnuLog 
         Caption         =   "Save Private Chat Log"
      End
      Begin VB.Menu mnuShowUserList 
         Caption         =   "Show User List"
      End
      Begin VB.Menu mnuInvite 
         Caption         =   "Invite Users"
      End
      Begin VB.Menu mnuLeave 
         Caption         =   "Leave Chat"
      End
   End
End
Attribute VB_Name = "frmPrivateChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ChatUsers() As String

Private Sub cmdAccept_Click()
frameInvite.Visible = False
SendChat "", 4
End Sub

Private Sub cmdDecline_Click()

SendChat "", 5
'SendDataToUDP User(UserIndex).IP, PrivateChatPacket(5, User(UserIndex).UniqueKey, Me.Tag, "Declined.")
DoEvents
Me.Visible = False
Unload Me
    
End Sub

Public Function GetNumChatUsers() As Long

    Dim i As Long
    
    For i = 0 To UBound(ChatUsers)
        If Len(ChatUsers(i)) = 8 Then GetNumChatUsers = GetNumChatUsers + 1
    Next i
'If GetNumChatUsers = 0 Then GetNumChatUsers = 1

End Function

Public Sub SyncPChatUsers(UniqueID As String)
    SendChatUserList UniqueID
End Sub

Private Sub Form_Load()
ReDim ChatUsers(0)
txtChat.ForeColor = Settings.ChatTextColor
txtChat.BackColor = Settings.ChatBGColor
txtChat.FontSize = Settings.ChatTextSize
txtEnter.BackColor = Settings.ChatBGColor
txtEnter.ForeColor = Settings.ChatTextColor
End Sub

Public Sub RequestChat(UserIndex As Integer)
frameInvite.Left = 0
frameInvite.Top = 0
frameInvite.Visible = True
lblInvite.Caption = GetUserNameByIndex(UserIndex) & " wants to initiate a private chat with you." & vbNewLine & vbNewLine & "Do you accept?"

    'SendDataToUDP User(UserIndex).IP, PrivateChatPacket(4, User(UserIndex).UniqueKey, Me.Tag, "Connected.")

End Sub

Private Sub SendChat(Text As String, Optional State As Integer = 2)
Dim UserIndex As Integer
Dim blDidSend As Boolean

For i = 1 To UBound(ChatUsers)
    If LenB(ChatUsers(i)) > 0 Then
        UserIndex = UserIndexByUID(ChatUsers(i))
            If UserIndex > 0 Then SendCryptTo UserIndex, PrivateChatPacket(State, GetNumChatUsers, User(UserIndex).UniqueID, Me.Tag, Text): blDidSend = True
    End If
Next i

If Not blDidSend Then AddChat "Your message echoes endlessly into the void..", "Nobody"

End Sub

Private Sub SendChatUserList(UID As String)
Dim UserIndex As Integer
UserIndex = UserIndexByUID(UID)
AddDebug "Sending PrivateChatUserListPacket 7 of PChat" & Me.Tag & " to: " & UID
If UserIndex > 0 Then SendCryptTo UserIndex, PrivateChatUserListPacket(7, GetNumChatUsers, User(UserIndex).UniqueID, Me.Tag, ChatUsers)

End Sub

Public Sub AddChat(Text As String, Name As String)
    txtChat.Text = txtChat.Text & "<" & Name & "> " & Text & vbCrLf
End Sub

Private Sub Form_Resize()
On Error Resume Next

txtChat.Width = Me.Width - 450
txtChat.Height = Me.Height - txtEnter.Height - 1200
txtEnter.Width = txtChat.Width
txtEnter.Top = txtChat.Height + txtChat.Top + 120

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SendChat "", 3
    DoEvents
End Sub

Private Sub mnuLeave_Click()
Unload Me
DoEvents
End Sub

Private Sub mnuLog_Click()
Dim strUsers As String

For i = 1 To UBound(ChatUsers)
    If Len(ChatUsers(i)) = 8 Then
    'UserIndex = UserIndexByUID(ChatUsers(i))
    If Len(strUsers) = 0 Then strUsers = GetUserName(ChatUsers(i))
        strUsers = strUsers & " - " & GetUserName(ChatUsers(i))
    End If
Next i

    WriteLog "Chat log from: " & Format(Now, "mm-dd-yyyy hh-mm-ss AM/PM") & vbCrLf & "Between you and: " & strUsers & vbCrLf & txtChat.Text, FixFilePath(App.Path & "\Private Chat " & strUsers & ".log")

End Sub

Private Sub txtChat_Change()
Dim ChatLen As Long
ChatLen = Len(txtChat.Text)

If ChatLen > 65000 Then
    If MsgBox("I would recommend clearing the chat box at this point, otherwise we may crash!", vbYesNo, "Clear chat history?") = vbYes Then
        txtChat.Text = ""
    End If
Else
    txtChat.SelStart = ChatLen - 1
End If
    
End Sub

Public Sub AddChatUser(UniqueID As String)
Dim ChatIndex As Integer
Dim i As Integer

If Not Len(UniqueID) = 8 Then Exit Sub

ChatIndex = GetChatIndexFromUID(UniqueID)

If ChatIndex = -1 Then
    For i = 1 To UBound(ChatUsers)
    
        If ChatUsers(i) = vbNullString Then
            ChatIndex = i
            Exit For
        End If
    Next i
    
    If ChatIndex = -1 Then
        ReDim Preserve ChatUsers(UBound(ChatUsers) + 1)
        ChatIndex = UBound(ChatUsers)
    End If
    
    lstUsers.AddItem GetUserName(UniqueID)
    lstUsers.ItemData(lstUsers.ListCount - 1) = ChatIndex
    
    ChatUsers(ChatIndex) = UniqueID
Else
    AddChat UniqueID & " already exists in the chat list?", "ERROR"
End If

End Sub

Public Sub RemoveChatUser(UniqueID As String)

For i = 1 To UBound(ChatUsers)

    If ChatUsers(i) = UniqueID Then
        ChatUsers(i) = vbNullString
        Exit For
    End If
Next i

If GetChatIndexFromUID(UniqueID) > -1 Then
    lstUsers.RemoveItem GetChatIndexFromUID(UniqueID)
    'lstUsers.ItemData(lstUsers.ListCount) = UserIndex
Else
    AddChat UniqueID & " doesn't exist in the chat?", "ERROR"
End If

If UBound(ChatUsers) <= 1 And frameInvite.Visible = True Then DoEvents: Unload Me

End Sub

Private Function GetChatIndexFromUID(UID As String) As Long

Dim i As Long

For i = 0 To lstUsers.ListCount - 1
    If IsNumeric(lstUsers.ItemData(i)) Then
        'ChatUsers(lstUsers.ItemData(i)) = UID
        GetChatIndexFromUID = i
        Exit Function
    End If
Next i

GetChatIndexFromUID = -1

End Function
Private Sub txtEnter_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
    Dim Text As String
    Text = Trim$(txtEnter.Text)
        SendChat Text
        AddChat Text, Settings.UserName
        txtEnter.Text = ""
    End If

End Sub
