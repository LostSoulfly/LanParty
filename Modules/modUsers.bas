Attribute VB_Name = "modUsers"
Option Explicit

Public Type LanUser
    UserName As String
    UniqueID As String
    CompName As String
    CurrentlyPlaying As String
    IP As String
    Muted As Boolean
    LanAdmin As Boolean
    LastHeard As Integer
    Online As Boolean
    UniqueKey As String
    MyUniqueKey As String
    AdminList() As String
    SyncingAdminList As Boolean
    SyncingVotes As Boolean
    HasSyncedVotes As Boolean
    AppVersion As String
    'ConnState As Long
End Type

Public User() As LanUser

Public Sub InitializeUsers()
    ReDim User(0) As LanUser
End Sub

Public Sub RemoveUser(UserIndex As Integer)

AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has left."
RemoveUserFromAllPChats User(UserIndex).UniqueID
RemoveUserFromChat UserIndex

DoEvents

ClearUser (UserIndex)
'HasSyncedAdmins = False
'IsSyncingAdmins = True
'frmMain.tmrAdmins.Enabled = True

End Sub

Public Sub ClearUser(UserIndex As Integer)

With User(UserIndex)
    .UserName = vbNullString
    .CompName = vbNullString
    .CurrentlyPlaying = vbNullString
    .UniqueID = vbNullString
    .IP = vbNullString
    .Muted = False
    .Online = False
    .LastHeard = 0
    .LanAdmin = False
    .UniqueKey = vbNullString
    .MyUniqueKey = vbNullString
    ReDim .AdminList(0)
    '.ConnState = 0
End With

End Sub

Public Function NewUser(Name As String, UID As String, IPAddress As String, HostName As String) As Integer
Dim UserIndex As Integer
UserIndex = UserIndexByUID(UID) 'check if the user is already in the array
If UserIndex = -1 Then 'if not
    If Not VerifyKeyAsString(Name) Then NewUser = -1: Exit Function
    UserIndex = FindEmptyUserSlot 'Check if there is an empty slot (due to user leaving, etc..)
        If UserIndex = -1 Then 'if now..
            ReDim Preserve User(UBound(User) + 1) 'dim a new user array, preserving the known users..
            UserIndex = UBound(User) 'set that userindex as the new array index
        Else
            AddDebug "Using empty UserIndex for NewUser " & Name & ": " & UserIndex
        End If
Else
    AddDebug "NewUser on existing UID: " & UID & " - " & Name
    Exit Function
End If

With User(UserIndex)
    .UserName = Name
    .CompName = HostName
    .CurrentlyPlaying = vbNullString
    .UniqueID = UID
    .IP = IPAddress
    .Muted = False
    .Online = True
    .LastHeard = 0
    .LanAdmin = False
    .UniqueKey = vbNullString
    .MyUniqueKey = vbNullString
    ReDim .AdminList(0)
    '.ConnState = 0
End With

IsSyncingAdmins = True
HasSyncedAdmins = False
frmMain.tmrAdmins.Enabled = True

NewUser = UserIndex
AddUserToChat UserIndex
AddDebug "New computer: " & IPAddress & " " & HostName & " " & UID & " " & Name
End Function

Public Function FindEmptyUserSlot() As Integer
Dim i As Integer

FindEmptyUserSlot = -1

For i = 1 To UBound(User)
    If LenB(User(i).UniqueID) = 0 Then FindEmptyUserSlot = i: Exit Function
Next

End Function



Public Function UserIndexByUID(UID As String) As Integer

    Dim i As Integer
    
    For i = 1 To UBound(User)
        If User(i).UniqueID = UID Then
            UserIndexByUID = i
            Exit Function
        End If
    Next i

UserIndexByUID = -1

End Function

Public Sub SetUserIndexStatus(UserIndex As Integer, blOnline As Boolean)

Select Case blOnline

Case Is = False
    User(UserIndex).LastHeard = 2
    User(UserIndex).Online = False
    
Case Is = True
    If User(UserIndex).Online = False Then frmMain.tmrAdmins.Enabled = True
    User(UserIndex).LastHeard = 0
    User(UserIndex).Online = True
    
End Select

UpdateUserCurrentStatus UserIndex

End Sub

Public Sub RemoveUserFromChat(UserIndex As Integer)
Dim ChatIndex As Integer
ChatIndex = GetUserChatIndex(UserIndex)

If ChatIndex >= 0 Then
    frmChat.lstUsers.RemoveItem (ChatIndex)
    AddDebug "Removing user from chat: " & GetUserNameByIndex(UserIndex)
    'frmChat.lstUsers.ItemData(frmChat.lstUsers.ListCount - 1) = UserIndex
Else
    'update here?
    AddDebug "RemoveUserFromChat: Not in list? - " & GetUserNameByIndex(UserIndex)
End If
End Sub

Public Sub AddUserToChat(UserIndex As Integer)
Dim ChatIndex As Integer
ChatIndex = GetUserChatIndex(UserIndex)

If ChatIndex = -1 Then
    frmChat.lstUsers.AddItem GetUserNameByIndex(UserIndex)
    frmChat.lstUsers.ItemData(frmChat.lstUsers.ListCount - 1) = UserIndex
    UpdateUserCurrentStatus UserIndex
Else
    'update here?
    AddDebug "AddUserToChat: Already in list - " & GetUserNameByIndex(UserIndex)
End If
End Sub

Public Function GetUserNameByIndex(UserIndex As Integer) As String
'This is used far more often than by UID.. so I don't want to
'call GetUserName and do more work by converting the index to a UID and then back to an index
'just to check if they are an admin or not. It's called too often.
GetUserNameByIndex = User(UserIndex).UserName
'If User(UserIndex).LanAdmin Then GetUserNameByIndex = "[Admin] " & GetUserNameByIndex
End Function

Public Function GetUserName(UniqueID As String) As String

    If Settings.UniqueID = UniqueID Then
        GetUserName = Settings.UserName
        'If Settings.LanAdmin Then GetUserName = "[Admin] " & GetUserName
    ElseIf UserIndexByUID(UniqueID) > 0 Then
        GetUserName = User(UserIndexByUID(UniqueID)).UserName
        'If User(UserIndexByUID(UniqueID)).LanAdmin Then GetUserName = "[Admin] " & GetUserName
    Else
        GetUserName = "Unknown??"
    End If
    
End Function

Public Sub ChangeUserName(UserIndex As Integer, Name As String)
    Dim ChatIndex As Integer
    ChatIndex = GetUserChatIndex(UserIndex)
    
    If ChatIndex = -1 Then AddDebug "CUCN: Not in list - " & Name: Exit Sub
    
    AddChat "[System] Changing " & GetUserNameByIndex(UserIndex) & "'s name to: " & Name
    User(UserIndex).UserName = Name
    frmChat.lstUsers.List(ChatIndex) = Name
    
End Sub

Public Function GetUserChatIndex(UserIndex As Integer) As Integer

Dim i As Integer


For i = 0 To frmChat.lstUsers.ListCount - 1
    If frmChat.lstUsers.ItemData(i) = UserIndex Then
        GetUserChatIndex = i
        Exit Function
    End If
Next i

GetUserChatIndex = -1

End Function

Public Function GetUserIndexFromChat() As Integer
On Error GoTo oops
    GetUserIndexFromChat = frmChat.lstUsers.ItemData(frmChat.lstUsers.ListIndex)

Exit Function
oops:

GetUserIndexFromChat = -1

End Function

Public Function GetMyUniqueKeyAsByte(UID As String) As Byte()
Dim UserIndex As Integer
UserIndex = UserIndexByUID(UID)
If UserIndex = -1 Then Exit Function

    GetMyUniqueKeyAsByte = User(UserIndex).MyUniqueKey

End Function

Public Function GetMyUniqueKeyAsByteByIndex(UserIndex As Integer) As Byte()

    GetMyUniqueKeyAsByteByIndex = User(UserIndex).MyUniqueKey

End Function

Public Function GetMyRemoteKeyAsByte(UID As String) As Byte()
Dim UserIndex As Integer
UserIndex = UserIndexByUID(UID)
If UserIndex = -1 Then Exit Function

    GetMyRemoteKeyAsByte = User(UserIndex).UniqueKey

End Function

Public Function GetMyRemoteKeyAsByteByIndex(UserIndex As Integer) As Byte()

    GetMyRemoteKeyAsByteByIndex = User(UserIndex).UniqueKey

End Function

Public Function MuteUser(UserIndex As Integer, Mute As Boolean)

    If UserIndex = -1 Then
        frmChat.txtEnter.Enabled = Not Mute
        AddChat "[System] " & "You have been globally " & IIf(Mute, "muted.", "unmuted.")
        Exit Function
    End If

    AddChat "[System] " & "You have " & IIf(Mute, "muted ", "unmuted ") & GetUserNameByIndex(UserIndex) & "."

    User(UserIndex).Muted = Mute
    frmMain.UpdateMenu UserIndex
End Function

Public Function MuteUserUID(UID As String, Mute As Boolean)

    If UID = Settings.UniqueID Then
        frmChat.txtEnter.Enabled = Not Mute
        AddChat "[System] " & "You have been globally " & IIf(Mute, "muted.", "unmuted.")
        Exit Function
    End If

    User(UserIndexByUID(UID)).Muted = Mute
    frmMain.UpdateMenu UserIndexByUID(UID)
End Function

Public Function GetUserCount() As Integer
Dim i As Integer

    For i = 1 To UBound(User)
        If LenB(User(i).UniqueID) > 0 Then
            GetUserCount = GetUserCount + 1
        End If
    Next i


End Function

Public Sub UpdateUserCurrentStatus(Optional UserIndex As Integer)
Dim ChatIndex As Integer, CurGame As Integer
Dim strAdmin As String
If UserIndex = 0 Then UserIndex = GetUserIndexFromChat
ChatIndex = GetUserChatIndex(UserIndex)

If UserIndex = -1 Or ChatIndex = -1 Then Exit Sub
CurGame = GameIndexByUID(User(UserIndex).CurrentlyPlaying)

If User(UserIndex).LanAdmin Then strAdmin = "[A] "

    If Settings.ShowStatus Then

        If CurGame > 0 Then
            frmChat.lstUsers.ToolTipText = "Currently playing: " & GetGameName(CurGame)
            frmChat.lstUsers.List(ChatIndex) = strAdmin & GetUserNameByIndex(UserIndex) & " [" & GetGameName(CurGame) & "]"
        ElseIf User(UserIndex).Online = False Then
            frmChat.lstUsers.List(ChatIndex) = strAdmin & GetUserNameByIndex(UserIndex) & " [Offline]"
            frmChat.lstUsers.ToolTipText = "Offline.."
        Else
            frmChat.lstUsers.ToolTipText = "Not currently playing a game."
            frmChat.lstUsers.List(ChatIndex) = strAdmin & GetUserNameByIndex(UserIndex) & " [Online]"
        End If
    
    Else
        frmChat.lstUsers.List(ChatIndex) = GetUserNameByIndex(UserIndex)
        frmChat.lstUsers.ToolTipText = ""
    End If
End Sub
