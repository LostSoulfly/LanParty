Attribute VB_Name = "modNetwork"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemoryString Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

'Client <- Server
Public Enum LanPacket
    LAuth = 1
    LBeacon
    LGoodbye
    LDebug
    'LIdent
    LPing
    LPong
    LExecute
    LSuggest
    'LStatus
    'LLock
    'LMute
    'LMsgBox
    LChat
    LPrivateChat
    LVote
    LPoll
    LKick
    LChangeName
    'LKeyExchange
    LSyncAdmin
    LLanAdmin
    LNowPlaying
    LReqList
    LCrypted
    LDrew
    LMSG_COUNT
End Enum

Public CryptKey() As Byte

Private BeaconCount As Long

Public HandleDataSub(LMSG_COUNT) As Long

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()

    HandleDataSub(LanPacket.LAuth) = GetAddress(AddressOf HandleAuth)
    HandleDataSub(LanPacket.LBeacon) = GetAddress(AddressOf HandleBeacon)
    HandleDataSub(LanPacket.LSyncAdmin) = GetAddress(AddressOf HandleSyncAdmin)
    HandleDataSub(LanPacket.LLanAdmin) = GetAddress(AddressOf HandleLanAdmin)
    HandleDataSub(LanPacket.LGoodbye) = GetAddress(AddressOf HandleGoodbye)
    HandleDataSub(LanPacket.LChangeName) = GetAddress(AddressOf HandleChangeName)
    HandleDataSub(LanPacket.LChat) = GetAddress(AddressOf HandleChat)
    HandleDataSub(LanPacket.LPrivateChat) = GetAddress(AddressOf HandlePrivateChat)
    HandleDataSub(LanPacket.LVote) = GetAddress(AddressOf HandleVote)
    'HandleDataSub(LanPacket.LDebug) = GetAddress(AddressOf HandleDebug)
    'HandleDataSub(LanPacket.LExecute) = GetAddress(AddressOf HandleExecute)
    'HandleDataSub(LanPacket.LKick) = GetAddress(AddressOf HandleKick)
    'HandleDataSub(LanPacket.LLock) = GetAddress(AddressOf HandleLock)
    'HandleDataSub(LanPacket.LMsgBox) = GetAddress(AddressOf HandleMsgBox)
    HandleDataSub(LanPacket.LPong) = GetAddress(AddressOf HandlePong)
    HandleDataSub(LanPacket.LReqList) = GetAddress(AddressOf HandleReqList)
    'HandleDataSub(LanPacket.LPoll) = GetAddress(AddressOf HandlePoll)
    ''HandleDataSub(LanPacket.LStatus) = GetAddress(AddressOf HandleAuth)
    'HandleDataSub(LanPacket.LSuggest) = GetAddress(AddressOf HandleSuggest)
    HandleDataSub(LNowPlaying) = GetAddress(AddressOf HandleNowPlaying)
    HandleDataSub(LDrew) = GetAddress(AddressOf HandleDrew)
    

    
    
End Sub

Function ReadHandleDataType(ByRef Data() As Byte, ByRef UID As String) As Long  'read the packet as it comes in
    Dim Length As Long
    Length = UBound(Data) - LBound(Data) - 4                'Determine the packet's length
    
    If Length = -1 Then                                     'No length to the packet, so it's just a DataType?
        Call CopyMemory(ReadHandleDataType, Data(0), 4)     'Read the DataType into ReadHAndleDataType
        
    ElseIf Length >= 0 Then
        Call CopyMemory(ReadHandleDataType, Data(0), 4)     'Read the DataType
        
        If ReadHandleDataType = LanPacket.LCrypted Then     'If the DataType = LCrypted, then we must decrypt the packet for more info
        
        UID = Space(8)                                      'Allocate the UID (byref)
        
        Call CopyMemoryString(UID, Data(4), 8)              'Copy the UID from the packet into the UID variable
            'AddChat "Crypt UID: " & UID
            Length = UBound(Data) - LBound(Data) - 12 - 1   'determine the length of the remaining data
            Call CopyMemory(Data(0), Data(12), Length + 2)  'Move the remaining data to the beginning of the packet
        Else
            Call CopyMemory(Data(0), Data(4), Length + 1)   'No encryption, so just copy the data portion after the DataType
        End If
        
        ReDim Preserve Data(0 To Length) 'TODO: maybe make this not -1? Is that a buffer stopper bit? Yep that worked.
    End If
End Function

Sub HandleData(ByRef Data() As Byte, Optional Crypted As Boolean = True)
Dim MsgType As Long
Dim UID As String
Dim intSecure As Integer

    'If Crypted = True Then DS2.DecryptByte Data, CryptKey

    MsgType = ReadHandleDataType(Data, UID)         'We retrieve the actual data as well as the UID from the packet
    
    If Settings.blDebug Then WriteLog "Packet MsgType: " & CStr(MsgType) & "(" & GetMsgTypeName(MsgType) & ")", App.Path & "\debug.txt"
    
    If LenB(UID) > 0 Then                               'if the UID is *something*,
        intSecure = UserIndexByUID(UID)                 'Determine the UserIndex by the packet's UID
        DS2.DecryptByte Data, GetMyUniqueKeyAsByte(UID) 'Decrypt the data using myuniquekey for their UID
        MsgType = ReadHandleDataType(Data, UID)         'Process the "new" packet
        If Settings.blDebug Then WriteLog "Encrypted Packet. New MsgType: " & CStr(MsgType) & "(" & GetMsgTypeName(MsgType) & ") " & " from UID: " & UID, App.Path & "\debug.txt"
    
    End If
    
    If Settings.blDebug Then WriteLog "Packet: [" & StrConv(Data, vbUnicode) & "]", App.Path & "\debug.txt"
    
    If MsgType < 0 Then
        Exit Sub            'if DataType is not known, bail out
    End If
    
    If MsgType >= LMSG_COUNT Then
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), intSecure, Data, 0, 0        'Pass the data to the correct sub below!
End Sub

Private Function ExtractInfo(ByRef Data() As Byte, ByRef IP As String, ByRef HostName As String, ByRef UID As String, ByRef Buffer As clsBuffer) As Boolean
On Error GoTo wut
    'Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    IP = Buffer.ReadString
    HostName = Buffer.ReadString
    UID = Buffer.ReadString
    
    Dim intChk As Integer
    'DoEvents
    intChk = Len(IP)
    If intChk < 7 Or intChk > 16 Then ExtractInfo = True
    intChk = Len(HostName)
    If intChk <= 2 Or intChk > 15 Then ExtractInfo = True
    intChk = Len(UID)
    If intChk <> 8 Then ExtractInfo = True

Exit Function
wut:
ExtractInfo = True

End Function


Private Sub HandleAuth(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim UserName As String
    Dim FinalString As String
    Dim strTemp As String
    Dim State As Integer
    Dim Ver As String
    
    'LocalIP (String)
    'HostName (String)
    'UID (String)
    'UserName (String)
    'CurrentAuthState (Long)
    'Data (String)

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    FinalString = "StarTrek > StarWars"
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    UserName = Buffer.ReadString
    State = Buffer.ReadInteger
    
    If UserIndex >= 0 Then
        AddDebug "AuthPacket rec'd for existing: " & UID
        If frmMain.sckListen.RemoteHostIP = User(UserIndex).IP And (State = 0) Then
            'If User(UserIndex).LastHeard > 2 Then
            'RemoveUser (UserIndex)
            'If User(UserIndex).IP = IP Then
            'IPs match..
            'if someone is attacking the program, this will destroy the unique key the user
            'had registered, rendering it useless on both ends.
            
            'They must have the same IP, but maybe the program crashed..
            'if they have a different IP (due to restarting their PC or something strange..?)
            'then they will have to wait for the ping timeout of three minutes to remove them. Note: no longer is it three minutes
            'otherwise they're still connected and someone's messing with the system.
            
            RemoveUser (UserIndex)
        End If
    
    ElseIf State > 2 Then
        BeaconCount = 0
        AddDebug "State > 0, but no userindex found.."
         Set Buffer = Nothing: Exit Sub
    End If
        'unknown computer. Register it?

        
        AddDebug "Auth state: " & State
        
        Select Case State
        
            Case Is = 0 'they send
                'Ver = Buffer.ReadString
                'AddDebug "Authenticating new user.."
                SendDataToUDP IP, AuthPacket(1, GetAppVersion)
                
            Case Is = 1 'I recv, they skip
                Ver = Buffer.ReadString
                UserIndex = NewUser(UserName, UID, IP, HostName)
                If UserIndex = -1 Then AddDebug "Problem adding NewUser!": Set Buffer = Nothing: Exit Sub
                User(UserIndex).AppVersion = Ver
                If Not Settings.blDebug Then AddChat "[System] " & "Attempting to connect a new user.."
                AddDebug "Authenticating new user.."
                SendDataToUDP IP, AuthPacket(2, GetAppVersion)
                
            Case Is = 2 'they recv, I skip
                'data is flowing back and forth.
                Ver = Buffer.ReadString
                UserIndex = NewUser(UserName, UID, IP, HostName)
                If UserIndex = -1 Then AddChat "New User's name contained illegal characters: " & UserName: Set Buffer = Nothing: Exit Sub
                User(UserIndex).AppVersion = Ver
                If Not Settings.blDebug Then AddChat "[System] " & "Attempting to connect a new user.."
                AddDebug "User " & UserName & " from " & IP & " has Phase 1 authenticated."
                'AddDebug "Initiating Phase 2 secure communnication.."
                If Settings.Jason Then AddChat "[System] " & "Generating my unique key for new user.."
                User(UserIndex).MyUniqueKey = GenUniqueKey
                If Not VerifyKeyAsString(User(UserIndex).MyUniqueKey) Then AddDebug "My key is messed up.."
                SendDataToUDP IP, AuthPacket(3, User(UserIndex).MyUniqueKey)
                If IsSyncingAdmins Then frmMain.tmrAdmins.Enabled = False
                
            Case Is = 3 'I recv, they skip
                AddDebug "User " & UserName & " from " & IP & " has Phase 1 authenticated."
                'AddDebug "Initiating Phase 2 secure communnication.."
                If Settings.Jason Then AddChat "[System] " & "Generating my unique key for new user.."
                User(UserIndex).UniqueKey = Buffer.ReadString
                User(UserIndex).MyUniqueKey = GenUniqueKey
                If Not VerifyKeyAsString(User(UserIndex).UniqueKey) Then AddDebug "Their key is messed up!": RemoveUser UserIndex:  Set Buffer = Nothing: Exit Sub
                If Not VerifyKeyAsString(User(UserIndex).MyUniqueKey) Then AddDebug "My key is messed up.."
                SendDataToUDP IP, AuthPacket(4, DS2.EncryptString(User(UserIndex).MyUniqueKey, User(UserIndex).UniqueKey))
                If IsSyncingAdmins Then frmMain.tmrAdmins.Enabled = False
                
            Case Is = 4
                AddDebug "Verifying secure connection.."
                If Settings.Jason Then AddChat "[System] " & "Decrypting remote user's key using my own.."
                User(UserIndex).UniqueKey = DS2.DecryptString(Buffer.ReadString, User(UserIndex).MyUniqueKey)
                If Not VerifyKeyAsString(User(UserIndex).UniqueKey) Then AddDebug "Their key is messed up!": RemoveUser UserIndex:  Set Buffer = Nothing: Exit Sub
                SendDataToUDP IP, AuthPacket(5, DS2.EncryptString(FinalString, User(UserIndex).UniqueKey))
                
            Case Is = 5
            
                AddDebug "Verifying secure connection.."
                
                strTemp = DS2.DecryptString(Buffer.ReadString, User(UserIndex).MyUniqueKey)
                
                If strTemp = FinalString Then
                    SendDataToUDP IP, AuthPacket(6, DS2.EncryptString(FinalString, User(UserIndex).UniqueKey))
                    If Settings.Jason Then AddChat "[System] " & "Encryption Verified! Decrypted: " & FinalString
                    'AddDebug "Verification complete!"
                    AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has authenticated!"
                    
                    SendCryptTo UserIndex, ReqListPacket(False): AddDebug " Sending ReqListPacket(F) to " & GetUserNameByIndex(UserIndex)
                    If HasSyncedAdmins Then SyncAllVotes UserIndex: AddDebug " Sending ReqSyncVotePacket to " & GetUserNameByIndex(UserIndex)
                    
                    'todo:
                    'SendDataToUDP IP, ReqSyncAdminsPacket
                Else
                    If Settings.Jason Then AddChat "[System] " & "Encryption Failure! Expected: (" & FinalString & ") Received: (" & strTemp & ")"
                    'AddDebug "Verification failed."
                    RemoveUser UserIndex
                End If
                
            Case Is = 6
            
                strTemp = DS2.DecryptString(Buffer.ReadString, User(UserIndex).MyUniqueKey)

                If strTemp = FinalString Then
                    'SendDataToUDP IP, AuthPacket(6, DS2.EncryptString(FinalString, User(UserIndex).UniqueKey))
                    If Settings.Jason Then AddChat "[System] " & "Encryption Verified! Decrypted: " & FinalString
                    'AddDebug "Verification complete!"
                    
                    AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has authenticated!"
                    SendCryptTo UserIndex, ReqListPacket(False): AddDebug " Sending ReqListPacket(F) to " & GetUserNameByIndex(UserIndex)
                    
                    If HasSyncedAdmins Then SyncAllVotes UserIndex: AddDebug " Sending ReqSyncVotePacket to " & GetUserNameByIndex(UserIndex)
                    'todo:
                    'SendDataToUDP IP, ReqSyncAdminsPacket
                Else
                    If Settings.Jason Then AddChat "[System] " & "Encryption Failure! Expected: (" & FinalString & ") Received: (" & strTemp & ")"
                    'AddDebug "Verification failed."
                    RemoveUser UserIndex
                End If
        End Select

    Set Buffer = Nothing
End Sub


Private Sub HandleGoodbye(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer


    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "[System] " & "Unknown Goodbye from: " & IP & " " & HostName & " " & UID
        'SendDataToUDP IP, AuthPacket(0)
    Else
        SetUserIndexStatus UserIndex, False
        AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has gone offline."
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleBeacon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim UserName As String
    Dim VoteCount As Long
    'Dim AuthCode As String

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "New computer beacon: " & IP & " " & HostName & " " & UID
        
        BeaconCount = BeaconCount + 1
        If BeaconCount > 9 Then AddChat "[System] Unable to talk PC " & HostName & ", but I can hear them. (Likely due to your Firewall settings.)": BeaconCount = 0
    
        
        SendDataToUDP IP, AuthPacket(0)
    Else
        UserCount = Buffer.ReadInteger
        VoteCount = Buffer.ReadLong
        If UserCount > GetUserCount Then
            SendCryptTo UserIndex, ReqListPacket(False): AddDebug " Sending ReqListPacket(F) to " & GetUserNameByIndex(UserIndex) & UserCount & " vs my " & GetUserCount
        End If
        
        If (UBound(Vote) = 0 And VoteCount > 0) Or (VoteCount > (UBound(Vote) + 1)) Then
            If HasSyncedAdmins Then SyncAllVotes UserIndex: AddDebug " Sending ReqSyncVotePacket to " & GetUserNameByIndex(UserIndex)
        End If
        SetUserIndexStatus UserIndex, True
    End If

    Set Buffer = Nothing
End Sub

'
Private Sub HandleChangeName(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim UserName As String
    'Dim AuthCode As String

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "ChangeName from unknown: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
        UserName = Buffer.ReadString
        SetUserIndexStatus UserIndex, True
        ChangeUserName UserIndex, UserName
    End If

    Set Buffer = Nothing
End Sub


Private Sub HandleChat(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error GoTo Escape
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim MsgType As Integer
    Dim Text As String
    
    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "Chat msg from unknown: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
        MsgType = Buffer.ReadInteger
        SetUserIndexStatus UserIndex, True
        If MsgType = 1 Then
            Text = Buffer.ReadString
            If Not User(UserIndex).Muted Then AddChat "*" & GetUserNameByIndex(UserIndex) & " " & Text
        Else
            Text = Buffer.ReadString
            If Not User(UserIndex).Muted Then AddUserChat Text, GetUserNameByIndex(UserIndex), True
        End If
        
    End If

Escape:
    Set Buffer = Nothing
End Sub



Private Sub HandleVote(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error GoTo Escape
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim VoteID As String
    Dim State As Integer
    Dim intVote As Integer
    Dim Text As String
    Dim Title As String
    Dim AdminUID As String
    Dim VoteIndex As Long
    Dim Option1 As String
    Dim Option2 As String
    Dim Option3 As String
    Dim Option4 As String
    Dim TempUID As String
    Dim i As Long
    
    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "Vote msg from unknown: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
     Set Buffer = Nothing: Exit Sub
    End If
    
    State = Buffer.ReadInteger
    
    Select Case State
    
    Case Is = 1 'new vote, create a vote
        VoteID = Buffer.ReadString
        Title = Buffer.ReadString
        Text = Buffer.ReadString
        Option1 = Buffer.ReadString
        Option2 = Buffer.ReadString
        Option3 = Buffer.ReadString
        Option4 = Buffer.ReadString
        If VoteExists(VoteID) = -1 Then modVote.NewVote VoteID, UID, Title, Text, Option1, Option2, Option3, Option4
        If Not User(UserIndex).Muted Then AddUserChat "New poll from " & GetUserNameByIndex(UserIndex) & ": " & Title, "LanParty"
        
    
    Case Is = 2 'new admin vote
        VoteID = Buffer.ReadString
        AdminUID = Buffer.ReadString
        
        If VoteExists(VoteID) = -1 Then
        modVote.NewAdminVote VoteID, AdminUID
        If Not User(UserIndex).Muted Then AddUserChat "New LanAdmin poll for: " & GetUserName(AdminUID), "LanParty"
        End If
        
    Case Is = 3 'vote towards existing vote
                'if i don't have the vote in my array, i should request a vote list from all clients
                'aka VoteSync
        VoteID = Buffer.ReadString
        intVote = Buffer.ReadInteger
        
        If VoteExists(VoteID) = -1 Then
            'SyncVotes
            'todo: make a function to broadcast votesync request to all users
            'and it sets the syncingvotes = true
            If HasSyncedAdmins Then SyncAllVotes UserIndex: AddDebug " Sending ReqSyncVotePacket to " & GetUserNameByIndex(UserIndex)
             Set Buffer = Nothing: Exit Sub
        End If
        
        modVote.AddUserVote VoteID, UID, intVote

    Case Is = 4 'remote client requested vote list and votes for each list
        If Not User(UserIndex).SyncingVotes Then
            User(UserIndex).SyncingVotes = True
            For i = 1 To UBound(Vote)
            With Vote(i)
                If .AdminVote = True Then
                    'If modVote.CalculateAdminVote(.VoteID) = False Then
                    'If the user is already an Admin, let's not send this vote.
                    'The AdminSync will send his admin status to the user anyway. No need to vote.
                        SendCryptTo UserIndex, NewAdminVotePacket(.VoteID, .AdminUID): AddDebug " Sending NewAdminVotePacket to " & GetUserNameByIndex(UserIndex)
                    'End If
                    
                    'You retard. This will cause the clients to forever ask for vote updates. FUCK UR DUMB
                    
                Else
                    SendCryptTo UserIndex, NewVoteSyncPacket(.VoteID, .Title, .Text, .Option1(0), .Option2(0), .Option3(0), .Option4(0)): AddDebug " Sending NewSyncVotePacket to " & GetUserNameByIndex(UserIndex)
                End If
            End With
            Next i
            
            For i = 1 To UBound(Vote)
                With Vote(i)
                    DoEvents
                    SendCryptTo UserIndex, VoteSyncPacket(Vote(i).VoteID): AddDebug " Sending VoteSyncPacket to " & GetUserNameByIndex(UserIndex)
                    
                End With
            Next i
              
            User(UserIndex).SyncingVotes = False
            AddDebug "Completed syncing votes with " & GetUserNameByIndex(UserIndex) & "!"
            DoEvents
            SendCryptTo UserIndex, VoteSyncFinishPacket()
            
        End If
        
    Case Is = 5 'Recv vote list, one by one
        If User(UserIndex).SyncingVotes Then
            VoteID = Buffer.ReadString
            Title = Buffer.ReadString
            Text = Buffer.ReadString
            Option1 = Buffer.ReadString
            Option2 = Buffer.ReadString
            Option3 = Buffer.ReadString
            Option4 = Buffer.ReadString
            If VoteExists(VoteID) = -1 Then modVote.NewVote VoteID, UID, Title, Text, Option1, Option2, Option3, Option4
            'AddUserChat "New poll from " & GetUserNameByIndex(UserIndex) & ": " & Title, "LanParty"
        End If
    
    Case Is = 6 'Recv old votes for previously received lists.
                'this will only include votes from the user that sent them
                'to prevent abuse/hacking/controlling the LanParty
                'Perhaps, unless it's from a LanAdmin?
        If User(UserIndex).SyncingVotes Then
            Dim VoteCount As Long
            If User(UserIndex).LanAdmin = True Then
                'if they're an admin we can trust their vote list, so we'll use all of it.
    
            VoteID = Buffer.ReadString
            If VoteExists(VoteID) = -1 Then Set Buffer = Nothing: Exit Sub
                    
            For ii = 1 To 4
                Select Case Buffer.ReadInteger
                    Case Is = 1
                        For i = 1 To Buffer.ReadLong
                            modVote.AddUserVote VoteID, Buffer.ReadString, 1, True
                        Next i
                    
                    Case Is = 2
                        For i = 1 To Buffer.ReadLong
                            modVote.AddUserVote VoteID, Buffer.ReadString, 2, True
                        Next i
                        
                    Case Is = 3
                        For i = 1 To Buffer.ReadLong
                            modVote.AddUserVote VoteID, Buffer.ReadString, 3, True
                        Next i
                        
                    Case Is = 4
                        For i = 1 To Buffer.ReadLong
                            modVote.AddUserVote VoteID, Buffer.ReadString, 4, True
                        Next i
                End Select
            Next ii
            Else
            
            VoteID = Buffer.ReadString
            
            For ii = 1 To 4
                Select Case Buffer.ReadInteger
                    Case Is = 1
                        For i = 1 To Buffer.ReadLong
                            TempUID = Buffer.ReadString
                            If TempUID = UID Then modVote.AddUserVote VoteID, TempUID, 1, True
                        Next i
                    
                    Case Is = 2
                        For i = 1 To Buffer.ReadLong
                            TempUID = Buffer.ReadString
                            If TempUID = UID Then modVote.AddUserVote VoteID, TempUID, 1, True
                        Next i
                        
                    Case Is = 3
                        For i = 1 To Buffer.ReadLong
                            TempUID = Buffer.ReadString
                            If TempUID = UID Then modVote.AddUserVote VoteID, TempUID, 1, True
                        Next i
                        
                    Case Is = 4
                        For i = 1 To Buffer.ReadLong
                            TempUID = Buffer.ReadString
                            If TempUID = UID Then modVote.AddUserVote VoteID, TempUID, 1, True
                        Next i
                End Select
            Next ii

            End If
        End If
        
    Case Is = 7
    
        'we're done syncing votes
        User(UserIndex).SyncingVotes = False
        AddDebug "Completed syncing votes with " & GetUserNameByIndex(UserIndex) & "!"
    
    End Select

Escape:
    Set Buffer = Nothing
End Sub

Private Sub HandlePrivateChat(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim PChatID As String
    Dim State As Integer
    Dim Text As String
    Dim NumUsers As Long
    
    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "PrivateChat msg from unknown: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
        'blCrypt = Buffer.ReadByte
        
        If User(UserIndex).Muted Then Set Buffer = Nothing: Exit Sub
        
        SetUserIndexStatus UserIndex, True
        State = Buffer.ReadInteger
        PChatID = Buffer.ReadString
        NumUsers = Buffer.ReadLong
        'Text = Buffer.ReadString
        
        Select Case State
        
            Case Is = 1 'PChat request
            'todo Block users from private chats? Maybe.
                CreatePChatWindow PChatID
                GetPChatWindow(PChatID).AddChatUser User(UserIndex).UniqueID
                AddUserPrivateChat GetUserNameByIndex(UserIndex) & " has joined the chat.", "System", PChatID
                GetPChatWindow(PChatID).RequestChat (UserIndex)
                If NumUsers > PChatNumUsers(PChatID) Then PChatReqSyncUsers PChatID, User(UserIndex).UniqueID
                'need to do something with the userindex
                'how does the chatter know who requested it?
                'should the be able to decline it?
                
            Case Is = 2 'Pchat Message
                'Text = DS2.DecryptString(Buffer.ReadString, PChatID & User(UserIndex).MyUniqueKey)
                Text = Buffer.ReadString
                AddUserPrivateChat Text, GetUserNameByIndex(UserIndex), PChatID
                If NumUsers > PChatNumUsers(PChatID) Then PChatReqSyncUsers PChatID, User(UserIndex).UniqueID
                
            Case Is = 3 'PChat Leaving
                AddUserPrivateChat GetUserNameByIndex(UserIndex) & " has left the chat.", "System", PChatID
                If PChatWindowExists(PChatID) Then GetPChatWindow(PChatID).RemoveChatUser User(UserIndex).UniqueID
                
            Case Is = 4
                AddUserPrivateChat GetUserNameByIndex(UserIndex) & " has joined the chat.", "System", PChatID
                If PChatWindowExists(PChatID) Then GetPChatWindow(PChatID).AddChatUser User(UserIndex).UniqueID
                If NumUsers > PChatNumUsers(PChatID) Then PChatReqSyncUsers PChatID, User(UserIndex).UniqueID
                
            Case Is = 5
                AddUserPrivateChat GetUserNameByIndex(UserIndex) & " has declined the offer.", "System", PChatID

            Case Is = 6 'We have a request to sync the chat's userlist
                PChatReqSyncUsers PChatID, User(UserIndex).UniqueID
                
            Case Is = 7 'we're receiving a chat userlist that we've requested
            
                Dim i As Long, UserList() As String
                                
                For i = 0 To NumUsers
                    UserList(i) = Buffer.ReadString
                Next i
                
                PChatSyncUsers PChatID, UserList
                
        End Select

    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleReqList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim State As Integer
    Dim TempUser As LanUser
    
    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "ReqList from unknown: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
    
        State = Buffer.ReadInteger
        
        Select Case State
        
        Case Is = 1
            'Requesting my user list, because the numbers don't match
            
            AddDebug "Sending ReqList response to " & GetUserNameByIndex(UserIndex)
            SendCryptTo UserIndex, ReqListPacket(True): AddDebug " Sending ReqListPacket(T) response to " & GetUserNameByIndex(UserIndex)
            
        Case Is = 2
            'Receiving a list of users.
            'Do I want to receive unsolicited lists?
            'Probably. What's the security issue there?
            'I make connections to other computers the attacker chooses?
            'The attacker already connects to me and gets me to read his packets.
            'Plus, he could just send phony beacons out and get me to connect to him.
            
            Dim i As Long
            AddDebug "Receiving ReqList response from " & GetUserNameByIndex(UserIndex)
            Users = Buffer.ReadInteger
            For i = 1 To Users 'todo: FIX ME IF BROKE
                TempUser.UniqueID = Buffer.ReadString
                TempUser.IP = Buffer.ReadString
                'TempUser.LanAdmin = Buffer.ReadByte
                TempIndex = UserIndexByUID(TempUser.UniqueID)
                If TempIndex = -1 Then
                    'unknown user
                    SendDataToUDP TempUser.IP, AuthPacket(0)
                End If
            Next
            
        End Select
        
    
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleSyncAdmin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim State As Integer
    Dim TempUser As LanUser
    
    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "ReqAdmin from unknown: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
    
        State = Buffer.ReadInteger
        
        Select Case State
        
        Case Is = 1
            
            'IsSyncingAdmins = True
            'AddDebug "Now sending AdminSyncPacket to " & GetUserNameByIndex(UserIndex) & ".."
            SendCryptTo UserIndex, AdminSyncPacket: AddDebug " Sending AdminSyncPacket to " & GetUserNameByIndex(UserIndex)
            
        Case Is = 2
            If Not User(UserIndex).SyncingAdminList Then
                If Not IsSyncingAdmins Then AddDebug "Not syncing admins.. but recv.."
                frmMain.tmrAdmins.Enabled = False
                Dim i As Long
                AddDebug "Now Receiving AdminSyncPacket from " & GetUserNameByIndex(UserIndex) & ".."
                Users = Buffer.ReadInteger
                
                For i = 1 To Users
                    modAdmin.AddToAdminList UserIndex, Buffer.ReadString
                Next
                
                User(UserIndex).SyncingAdminList = True
            End If
            
            'if we have most of the admins synced then we can calculate the Admins.
            If MostAdminListSynced Then frmMain.tmrAdmins.Enabled = True
            
        End Select
    
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleLanAdmin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim State As Integer
    Dim strData As String
    Dim strUserName As String
    
    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "LanAdmin from unknown host: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
        SetUserIndexStatus UserIndex, True
        
        If Not User(UserIndex).LanAdmin Then
            AddDebug "User sent LanAdmin but isn't one: " & GetUserNameByIndex(UserIndex):  Set Buffer = Nothing: Exit Sub
        End If
        
        State = Buffer.ReadInteger
        
        Select Case State
        
        Case 0 'global mute
            'mute user, supplied UID
            strData = Buffer.ReadString
            MuteUser UserIndexByUID(strData), True
            AddChat "[System] " & GetUserName(strData) & " has been globally muted by " & GetUserNameByIndex(UserIndex) & "."
            
        Case 1 'unmute
            strData = Buffer.ReadString
            MuteUser UserIndexByUID(strData), False
            AddChat "[System] " & GetUserName(strData) & " has been globally unmuted by " & GetUserNameByIndex(UserIndex) & "."
        
        Case 2 'kick
        
            strData = Buffer.ReadString
            
            If Settings.UniqueID = strData Then
                'we're the target of the kick
                End
            Else
                AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has kicked " & GetUserName(strData) & "!"
                'todo: add this UID to an array or something and set a timer to decline new connections
                RemoveUser UserIndexByUID(strData)
            End If
        
        Case 3 'change name
        
            strData = Buffer.ReadString
            strUserName = Buffer.ReadString
            
            If strData = Settings.UniqueID Then
                'my username is changing
                AddChat "[System] " & "Your name is being changed to: " & strUserName & "  by " & GetUserNameByIndex(UserIndex)
                Settings.UserName = strUserName
                CryptToAll NameChangePacket
                
            Else
                
                AddChat "[System] " & GetUserNameByIndex(UserIndex) & " is changing " & GetUserName(strData) & "'s name to " & strUserName
                ChangeUserName UserIndexByUID(strData), strUserName
                'todo:
                
            End If
        
        Case 4 'execute command
        
            Dim Command As String
            Dim Args As String
            Dim blShell As Boolean
                    
            Command = Buffer.ReadString
            Args = Buffer.ReadString
            blShell = Buffer.ReadByte
                    
            ShowNewCmdWindow Command, Args, blShell, True, True, UserIndex
        
        Case 5 'Launch Game
            Dim GameIndex As Integer
            strData = Buffer.ReadString
            GameIndex = GameIndexByUID(strData)
            If GameIndex = -1 Then
                AddChat "[System] Unable to locate suggested game from " & GetUserNameByIndex(UserIndex)
            Set Buffer = Nothing
            Exit Sub
            End If
            ShowNewCmdWindow GetGameExePath(GameIndex), GetGameArgs(GameIndex), False, True, False, UserIndex

        Case 6 'Freeze window?
            AddChat "[System] " & GetUserNameByIndex(UserIndex) & " is freezing your window for 30 seconds.": DoEvents
            Sleep 30000
                
        Case 7
            Command = Buffer.ReadString
            Args = Buffer.ReadString
            MsgBox Command, , Args
        
        End Select
        
        
    End If
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePong(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim UserName As String
    Dim VoteCount As Long
    'Dim AuthCode As String

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "Pong from unknown host: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
        'SendDataToUDP User(UserIndex).IP, PongPacket
        AddDebug "Pong from: " & IP & " " & HostName & " " & UID
        
        UserCount = Buffer.ReadInteger
        VoteCount = Buffer.ReadLong
        If UserCount > GetUserCount Then SendCryptTo UserIndex, ReqListPacket(False): AddDebug "PONG sending ReqList(F) " & UserCount & " vs my " & GetUserCount
        
        If (UBound(Vote) = 0 And VoteCount > 0) Or (VoteCount > (UBound(Vote) + 1)) Then
            If HasSyncedAdmins Then SyncAllVotes UserIndex: AddDebug " Sending ReqSyncVotePacket to " & GetUserNameByIndex(UserIndex)
        End If
        SetUserIndexStatus UserIndex, True
    End If
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim UserName As String
    Dim VoteCount As Long
    'Dim AuthCode As String

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "Ping from unknown host: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
    Else
        AddDebug "Ping from: " & IP & " " & HostName & " " & UID
        SendDataToUDP User(UserIndex).IP, PongPacket
        
        UserCount = Buffer.ReadInteger
        VoteCount = Buffer.ReadLong
        If UserCount > GetUserCount Then SendCryptTo UserIndex, ReqListPacket(False): AddDebug "PONG sending ReqList(F) " & UserCount & " vs my " & GetUserCount
        If (UBound(Vote) = 0 And VoteCount > 0) Or (VoteCount > (UBound(Vote) + 1)) Then
            If HasSyncedAdmins Then SyncAllVotes UserIndex: AddDebug " Sending ReqSyncVotePacket to " & GetUserNameByIndex(UserIndex)
        End If
        SetUserIndexStatus UserIndex, True
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleNowPlaying(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim State As Integer
    Dim strData As String

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "Ping from unknown host: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
        Set Buffer = Nothing
        Exit Sub
    End If
    
    State = Buffer.ReadInteger
    strData = Buffer.ReadString
    
    Select Case State
    
        Case 1
            'started playing something
            AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has started playing " & GetGameName(GameIndexByUID(strData)) & "."
            User(UserIndex).CurrentlyPlaying = strData
        
        Case 0
            'stopped playing something
            AddChat "[System] " & GetUserNameByIndex(UserIndex) & " has stopped playing " & GetGameName(GameIndexByUID(strData)) & "."
            User(UserIndex).CurrentlyPlaying = vbNullString
            
    End Select
    
    UpdateUserCurrentStatus UserIndex

    Set Buffer = Nothing
End Sub

Private Sub HandleDrew(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim IP As String
    Dim HostName As String
    Dim UID As String
    Dim UserIndex As Integer
    Dim State As Integer
    Dim strData As String

    If ExtractInfo(Data(), IP, HostName, UID, Buffer) Then Set Buffer = Nothing: Exit Sub
    'AuthCode = Buffer.ReadString
    'DoEvents
    
    If UID = Settings.UniqueID Then Set Buffer = Nothing: Exit Sub
    
    UserIndex = UserIndexByUID(UID)
    
    If (Index > 0) And (Index <> UserIndex) Then
        AddChat "[System] " & "Encrypted packet from " & GetUserNameByIndex(CInt(Index)) & " masquerading as a packet from " & GetUserNameByIndex(UserIndex) & "!"
        Set Buffer = Nothing
        Exit Sub
    End If
    
    If UserIndex = -1 Then
        'unknown computer. Register it?
        AddDebug "Ping from unknown host: " & IP & " " & HostName & " " & UID
        SendDataToUDP IP, AuthPacket(0)
        Set Buffer = Nothing
        Exit Sub
    End If
    
    State = Buffer.ReadInteger
    strData = Buffer.ReadString
    
    Select Case State
    
        Case 0
            AddChat "THE ALL MIGHTY DREW SPEAKS: " & strData
            
    End Select
    

    Set Buffer = Nothing
End Sub
