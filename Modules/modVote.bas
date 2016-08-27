Attribute VB_Name = "modVote"
Option Explicit

Public Type VoteType
    VoteID As String
    AdminVote As Boolean
    RegularVote As Boolean
    AdminUID As String
    Title As String
    Text As String
    Option1() As String
    Option2() As String
    Option3() As String
    Option4() As String
End Type

Public Vote() As VoteType

Public Sub InitializeVotes()

    ReDim Vote(0) As VoteType
    
End Sub

Public Function VoteExists(VoteID As String) As Long
Dim i As Long

For i = 1 To UBound(Vote)
    If Vote(i).VoteID = VoteID Then VoteExists = i: Exit Function
Next

VoteExists = -1

End Function

Public Sub AddUserVote(VoteID As String, UID As String, intOption As Integer, Optional Syncing As Boolean = False)
'check if they've already voted
Dim VoteIndex As Long

If CheckVoteID(VoteID) Then Exit Sub

VoteIndex = VoteExists(VoteID)

If VoteIndex = -1 Then Exit Sub

If HasUserVoted(VoteIndex, UID) Then AddDebug UID & " already voted: " & VoteID: Exit Sub

Select Case intOption

Case Is = 1
    ReDim Preserve Vote(VoteIndex).Option1(UBound(Vote(VoteIndex).Option1) + 1)
    Vote(VoteIndex).Option1(UBound(Vote(VoteIndex).Option1)) = UID

Case Is = 2
    ReDim Preserve Vote(VoteIndex).Option2(UBound(Vote(VoteIndex).Option2) + 1)
    Vote(VoteIndex).Option2(UBound(Vote(VoteIndex).Option2)) = UID
    
Case Is = 3
    ReDim Preserve Vote(VoteIndex).Option3(UBound(Vote(VoteIndex).Option3) + 1)
    Vote(VoteIndex).Option3(UBound(Vote(VoteIndex).Option3)) = UID
    
Case Is = 4
    ReDim Preserve Vote(VoteIndex).Option4(UBound(Vote(VoteIndex).Option4) + 1)
    Vote(VoteIndex).Option4(UBound(Vote(VoteIndex).Option4)) = UID

End Select

If Not Syncing Then AddUserChat GetUserName(UID) & " has voted in " & Chr(34) & Vote(VoteIndex).Title & Chr(34), "System", False

If Vote(VoteIndex).AdminVote Then
    If CalculateAdminVote(VoteID) Then
        modAdmin.SetUserLanAdmin Vote(VoteIndex).AdminUID, True
    End If
End If

If frmVoteList.Visible = True Then frmVoteList.RefreshList
       
End Sub

Public Function WhatDidUserVote(VoteID As String, UID As String, blByIndex As Boolean) As Integer
'loop through all vote options and make sure they haven't voted yet!
Dim i As Long
Dim VoteIndex As Long

If CheckVoteID(VoteID) Then Exit Function

If blByIndex Then
    VoteIndex = CLng(VoteID)
Else
    VoteIndex = VoteExists(VoteID)
End If

If VoteIndex = -1 Then Exit Function

For i = 1 To UBound(Vote(VoteIndex).Option1)
    If Vote(VoteIndex).Option1(i) = UID Then WhatDidUserVote = 1: Exit Function
Next i

For i = 1 To UBound(Vote(VoteIndex).Option2)
    If Vote(VoteIndex).Option2(i) = UID Then WhatDidUserVote = 2: Exit Function
Next i

For i = 1 To UBound(Vote(VoteIndex).Option3)
    If Vote(VoteIndex).Option3(i) = UID Then WhatDidUserVote = 3: Exit Function
Next i

For i = 1 To UBound(Vote(VoteIndex).Option4)
    If Vote(VoteIndex).Option4(i) = UID Then WhatDidUserVote = 4: Exit Function
Next i

End Function

Private Function HasUserVoted(VoteIndex As Long, UID As String) As Boolean
'loop through all vote options and make sure they haven't voted yet!
Dim i As Long

For i = 1 To UBound(Vote(VoteIndex).Option1)
    If Vote(VoteIndex).Option1(i) = UID Then HasUserVoted = True: Exit Function
Next i

For i = 1 To UBound(Vote(VoteIndex).Option2)
    If Vote(VoteIndex).Option2(i) = UID Then HasUserVoted = True: Exit Function
Next i

For i = 1 To UBound(Vote(VoteIndex).Option3)
    If Vote(VoteIndex).Option3(i) = UID Then HasUserVoted = True: Exit Function
Next i

For i = 1 To UBound(Vote(VoteIndex).Option4)
    If Vote(VoteIndex).Option4(i) = UID Then HasUserVoted = True: Exit Function
Next i

End Function

Public Function NewAdminVote(VoteID As String, UID As String) As String
Dim VoteIndex As Long

If CheckVoteID(VoteID) Then Exit Function

VoteIndex = VoteExists(VoteID)

If VoteIndex > 0 Then Exit Function

VoteIndex = GetEmptyVote()

If VoteIndex = -1 Then
    VoteIndex = UBound(Vote) + 1
    ReDim Preserve Vote(VoteIndex)
End If

With Vote(VoteIndex)
    .AdminUID = UID
    .AdminVote = True
    
    If Len(VoteID$) = 0 Then
        .VoteID = GenUniqueKey(21) 'These don't need to be too long
    Else
        .VoteID = VoteID
    End If
    If UID = Settings.UniqueID Then
        .Title = "Admin Vote: " & GetUserName(UID) & " on PC " & frmMain.sckListen.LocalHostName
        .Text = "Do you vote for or against " & GetUserName(UID) & " on PC " & frmMain.sckListen.LocalHostName & " becoming a LAN Admin?"
    Else
        .Title = "Admin Vote: " & GetUserName(UID) & " on PC " & User(UserIndexByUID(UID)).CompName
        .Text = "Do you vote for or against " & GetUserName(UID) & " on PC " & User(UserIndexByUID(UID)).CompName & " becoming a LAN Admin?"
    End If
    '.Text = "Yes or no, for admin??"
    ReDim .Option1(0)
    ReDim .Option2(0)
    ReDim .Option3(0)
    ReDim .Option4(0)
    .Option1(0) = "I am for this user being a LAN Admin."
    .Option2(0) = "I am against this user being a LAN Admin."

    NewAdminVote = .VoteID
    AddDebug "New Admin Vote: " & .VoteID
    If frmVoteList.Visible = True Then frmVoteList.RefreshList
End With

End Function

Public Function NewVote(VoteID As String, UID As String, Title As String, Text As String, Option1 As String, Option2 As String, Option3 As String, Option4 As String) As String
Dim VoteIndex As Long

'If CheckVoteID(VoteID) Then Exit Function

VoteIndex = VoteExists(VoteID)
    
If VoteIndex > 0 Then Exit Function

VoteIndex = GetEmptyVote()

If VoteIndex = -1 Then
    VoteIndex = UBound(Vote) + 1
    ReDim Preserve Vote(VoteIndex)
End If

With Vote(VoteIndex)
    .RegularVote = True
    ReDim .Option1(0)
    ReDim .Option2(0)
    ReDim .Option3(0)
    ReDim .Option4(0)
    .Option1(0) = Option1
    .Option2(0) = Option2
    .Option3(0) = Option3
    .Option4(0) = Option4
    .Text = Text
    .Title = Title
    If Len(VoteID$) = 0 Then
        .VoteID = GenUniqueKey(21) 'These don't need to be too long. BUT WHAT THE HELL WHY NOT
    Else
        .VoteID = VoteID
    End If
    
    AddDebug "New Vote: " & .VoteID
    NewVote = .VoteID
End With



End Function

Public Function CalculateAdminVote(VoteID As String) As Boolean
Dim VoteIndex As LogEventTypeConstants
Dim VoteYes As Long
Dim VoteNo As Long

If CheckVoteID(VoteID) Then Exit Function

VoteIndex = VoteExists(VoteID)

If VoteIndex = -1 Then Exit Function

'If the YES vote has more array indexes (should be an easy and fairly accurate way)
'then we can assume that vote won.
If Vote(VoteIndex).RegularVote Or Vote(VoteIndex).AdminVote = False Then Exit Function

VoteYes = UBound(Vote(VoteIndex).Option1)
VoteNo = UBound(Vote(VoteIndex).Option2)

If VoteYes + VoteNo <= 1 Then Exit Function

If VoteYes > VoteNo Then
    CalculateAdminVote = True
Else
    CalculateAdminVote = False
End If
End Function

Public Sub RemoveVote(VoteID As String)
Dim VoteIndex As Long

If CheckVoteID(VoteID) Then Exit Sub

VoteIndex = VoteExists(VoteID)

If VoteIndex = -1 Then Exit Sub

With Vote(VoteIndex)
    .RegularVote = False
    .AdminVote = False
    .AdminUID = vbNullString
    ReDim .Option1(0)
    ReDim .Option2(0)
    ReDim .Option3(0)
    ReDim .Option4(0)
    .Option1(0) = vbNullString
    .Option2(0) = vbNullString
    .Option3(0) = vbNullString
    .Option4(0) = vbNullString
    .Text = vbNullString
    .VoteID = vbNullString 'These don't need to be too long
End With
AddDebug "Removing Vote: " & VoteID
End Sub

Private Function GetEmptyVote() As Long
Dim i As Long


For i = 1 To UBound(Vote)
    If Len(Vote(i).VoteID$) = 0 Then GetEmptyVote = i: AddDebug "Using Empty Vote.": Exit Function
Next i

GetEmptyVote = -1

End Function

Public Sub CalculatePercentage(VoteID As String, ByRef Op1 As String, ByRef Op2 As String, _
ByRef Op3 As String, ByRef Op4 As String)
On Error Resume Next
Dim VoteIndex As Long
Dim TotalVotes As Long

If CheckVoteID(VoteID) Then Exit Sub

VoteIndex = VoteExists(VoteID)

If VoteIndex = -1 Then Exit Sub

'add up all votes
TotalVotes = UBound(Vote(VoteIndex).Option1)
TotalVotes = TotalVotes + UBound(Vote(VoteIndex).Option2)
TotalVotes = TotalVotes + UBound(Vote(VoteIndex).Option3)
TotalVotes = TotalVotes + UBound(Vote(VoteIndex).Option4)

Op1 = FormatNumber(UBound(Vote(VoteIndex).Option1) / TotalVotes * 100)
Op2 = FormatNumber(UBound(Vote(VoteIndex).Option2) / TotalVotes * 100)
Op3 = FormatNumber(UBound(Vote(VoteIndex).Option3) / TotalVotes * 100)
Op4 = FormatNumber(UBound(Vote(VoteIndex).Option4) / TotalVotes * 100)

If Op1 = "" Then Op1 = "0"
If Op2 = "" Then Op2 = "0"
If Op3 = "" Then Op3 = "0"
If Op4 = "" Then Op4 = "0"


End Sub

Public Function ShowVoteWindow(VoteIndex As Long)

Dim frmNewVote As New frmVote

Load frmNewVote
frmNewVote.Tag = Vote(VoteIndex).VoteID
frmNewVote.RefreshVote
frmNewVote.Visible = True

End Function

Public Function GetAdminUIDFromVote(VoteID As String) As String

If CheckVoteID(VoteID) Then Exit Function

    GetAdminUIDFromVote = VoteExists(VoteID)
End Function

Private Function CheckVoteID(VoteID As String) As Boolean
On Error Resume Next

If Len(VoteID$) < 19 Then CheckVoteID = True: AddDebug "VoteID not >= 19 len"

End Function

Public Sub SyncAllVotes(Optional UserIndex As Integer = -1, Optional blForceSync As Boolean = False)
Dim i As Integer
Dim numLeft As Integer

If UserIndex = -1 Then

    For i = 1 To UBound(User)
        'If Not User(i).SyncingVotes Then
        If blForceSync Or User(i).HasSyncedVotes = False Then
            numLeft = numLeft + 1
            User(i).SyncingVotes = True
            User(i).HasSyncedVotes = False
            SendCryptTo i, ReqSyncVotePacket
            DoEvents
        End If
    Next
    
    If numLeft = 0 Then frmMain.tmrVotesSync.Enabled = False: Exit Sub
    AddDebug "Users left to sync votes with: " & numLeft
    
Else
    'User(UserIndex).SyncingVotes = False
    If blForceSync = False Or User(UserIndex).HasSyncedVotes Then Exit Sub
    
    User(UserIndex).SyncingVotes = True
    User(UserIndex).HasSyncedVotes = False
    SendCryptTo UserIndex, ReqSyncVotePacket
    
End If

End Sub
