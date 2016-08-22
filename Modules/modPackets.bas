Attribute VB_Name = "modPackets"
Public Function BuildGeneric(PacketNum As Long, Data As String) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong PacketNum
Buffer.WriteBytes PacketHeader
Buffer.WriteString Data

BuildGeneric = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function AuthPacket(Optional State As Integer = 0, Optional Data As String = "") As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

'If State = 0 Then Data = GetAppVersion

Buffer.WriteLong LanPacket.LAuth
Buffer.WriteBytes PacketHeader
Buffer.WriteString Settings.UserName
Buffer.WriteInteger State
Buffer.WriteString Data

AuthPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function GoodbyePacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LGoodbye
Buffer.WriteBytes PacketHeader

GoodbyePacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function NameChangePacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LChangeName
Buffer.WriteBytes PacketHeader
Buffer.WriteString Settings.UserName

NameChangePacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function BeaconPacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LBeacon
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger GetUserCount
Buffer.WriteLong UBound(Vote)

BeaconPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function PrivateChatPacket(State As Integer, NumChatUsers As Long, RemoteKey As String, PChatID As String, Text As String) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer
Dim Temp As String
Buffer.WriteLong LanPacket.LPrivateChat
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger State
Buffer.WriteString PChatID
Buffer.WriteLong NumChatUsers

'not tested
'Temp = DS2.EncryptString(PChatID, RemoteKey)
'Buffer.WriteLong Len(Temp)
'Buffer.WriteBytes Temp

Select Case State

Case Is = 2
    'Buffer.WriteString DS2.EncryptString(Text, PChatID & RemoteKey)    'useless encryption..
    Buffer.WriteString Text
End Select

PrivateChatPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function PrivateChatUserListPacket(State As Integer, NumChatUsers As Long, RemoteKey As String, PChatID As String, Users() As String) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer
Dim Temp As String
Buffer.WriteLong LanPacket.LPrivateChat
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger State
Buffer.WriteString PChatID
Buffer.WriteLong NumChatUsers

Select Case State

Case Is = 6 'requesting a list
    '

Case Is = 7 'sending a list of users

    For i = 0 To UBound(Users)
        Buffer.WriteString
    Next i

End Select

PrivateChatPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function ChatPacket(Text As String, Optional intType As Integer = 0) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LChat
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger intType
Buffer.WriteString Text

ChatPacket = Buffer.ToArray

Set Buffer = Nothing

End Function


'Public Function ChatPacket(Text As String, Optional blCrypt As Boolean = False, Optional RemoteKey As String = "") As Byte()
'Dim Buffer As New clsBuffer
'Set Buffer = New clsBuffer
'
'Buffer.WriteLong LanPacket.LChat
'Buffer.WriteBytes PacketHeader
'Buffer.WriteByte blCrypt
'If blCrypt Then
'    Buffer.WriteString DS2.EncryptString(Text, RemoteKey)
'Else
'    Buffer.WriteString Text
'End If
'
'ChatPacket = Buffer.ToArray
'
'Set Buffer = Nothing
'
'End Function

Public Function PingPacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LPing
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger GetUserCount
Buffer.WriteLong UBound(Vote)

'Debug.Print Buffer.ToString

PingPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function PongPacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LPong
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger GetUserCount
Buffer.WriteLong UBound(Vote)

PongPacket = Buffer.ToArray

Set Buffer = Nothing

End Function


Public Function ReqListPacket(SendList As Boolean) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LReqList
Buffer.WriteBytes PacketHeader
If SendList Then
    'we're generating a list packet to send to the remote client
    Buffer.WriteInteger 2
    Buffer.WriteInteger GetUserCount
    Dim i As Integer
    For i = 1 To UBound(User)
        If Not LenB(User(i).UniqueID$) = 0 Then
            Buffer.WriteString User(i).UniqueID
            Buffer.WriteString User(i).IP
            'Buffer.WriteByte User(i).LanAdmin
        End If
    Next

Else
    Buffer.WriteInteger 1
    Buffer.WriteInteger GetUserCount

End If
ReqListPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function NewVotePacket(VoteID As String, Title As String, Text As String, Option1 As String, _
Option2 As String, Option3 As String, Option4 As String) As Byte()

Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 1
Buffer.WriteString VoteID
Buffer.WriteString Title
Buffer.WriteString Text
Buffer.WriteString Option1
Buffer.WriteString Option2
Buffer.WriteString Option3
Buffer.WriteString Option4

NewVotePacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function NewAdminVotePacket(VoteID As String, AdminUID As String) As Byte()

Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 2
Buffer.WriteString VoteID
Buffer.WriteString AdminUID

NewAdminVotePacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function VotePacket(VoteID As String, Vote As Integer) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

'state should be > 1 as 1 is the new vote packet.
'we should be 3 for new votes?

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 3
Buffer.WriteString VoteID
Buffer.WriteInteger Vote


VotePacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function VoteSyncFinishPacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

'state should be > 1 as 1 is the new vote packet.
'we should be 3 for new votes?

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 7

VoteSyncFinishPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function ReqSyncVotePacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 4


ReqSyncVotePacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function NewVoteSyncPacket(VoteID As String, Title As String, Text As String, Option1 As String, _
Option2 As String, Option3 As String, Option4 As String) As Byte()

Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 5
Buffer.WriteString VoteID
Buffer.WriteString Title
Buffer.WriteString Text
Buffer.WriteString Option1
Buffer.WriteString Option2
Buffer.WriteString Option3
Buffer.WriteString Option4

NewVoteSyncPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function VoteSyncPacket(VoteID As String) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer
Dim VoteIndex As Long
Dim i As Integer

Buffer.WriteLong LanPacket.LVote
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 6
Buffer.WriteString VoteID

VoteIndex = VoteExists(VoteID)
If VoteIndex = -1 Then Exit Function


Buffer.WriteInteger 1                               'Vote option
Buffer.WriteLong UBound(Vote(VoteIndex).Option1)    'number of votes
For i = 1 To UBound(Vote(VoteIndex).Option1)
    Buffer.WriteString Vote(VoteIndex).Option1(i)                      'each vote
Next i

Buffer.WriteInteger 2                               'Vote option
Buffer.WriteLong UBound(Vote(VoteIndex).Option2)    'number of votes
For i = 1 To UBound(Vote(VoteIndex).Option2)
    Buffer.WriteString Vote(VoteIndex).Option2(i)                      'each vote
Next i

Buffer.WriteInteger 3                               'Vote option
Buffer.WriteLong UBound(Vote(VoteIndex).Option3)    'number of votes
For i = 1 To UBound(Vote(VoteIndex).Option3)
    Buffer.WriteString Vote(VoteIndex).Option3(i)                      'each vote
Next i

Buffer.WriteInteger 4                               'Vote option
Buffer.WriteLong UBound(Vote(VoteIndex).Option4)    'number of votes
For i = 1 To UBound(Vote(VoteIndex).Option4)
    Buffer.WriteString Vote(VoteIndex).Option4(i)                      'each vote
Next i


VoteSyncPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function ReqAdminSyncPacket() As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LSyncAdmin
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 1

ReqAdminSyncPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function AdminSyncPacket() As Byte()
Dim Buffer As New clsBuffer
Dim i As Integer
Dim AdminCount As Integer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LSyncAdmin
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger 2

    For i = 1 To UBound(User)
        If User(i).LanAdmin Then AdminCount = AdminCount + 1
    Next i
    If Settings.LanAdmin Then AdminCount = AdminCount + 1
    Buffer.WriteInteger AdminCount
    If Settings.LanAdmin Then Buffer.WriteString Settings.UniqueID
    For i = 1 To UBound(User)
        If User(i).LanAdmin Then Buffer.WriteString User(i).UniqueID
    Next i
    
AdminSyncPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function LanAdminPacket(State As Integer, Optional Data As String = "", Optional UserName As String = "") As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LLanAdmin
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger State
Buffer.WriteString Data
If State = 3 Then Buffer.WriteString UserName

LanAdminPacket = Buffer.ToArray

Set Buffer = Nothing

End Function


Public Function LanAdminExecPacket(State As Integer, Command As String, Args As String, blShell As Boolean) As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LLanAdmin
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger State
Buffer.WriteString Command
Buffer.WriteString Args
Buffer.WriteByte blShell

LanAdminExecPacket = Buffer.ToArray

Set Buffer = Nothing

End Function

Public Function NowPlayingPacket(State As Integer, Optional Data As String = "") As Byte()
Dim Buffer As New clsBuffer
Set Buffer = New clsBuffer

Buffer.WriteLong LanPacket.LNowPlaying
Buffer.WriteBytes PacketHeader
Buffer.WriteInteger State
Buffer.WriteString Data

NowPlayingPacket = Buffer.ToArray

Set Buffer = Nothing

End Function
