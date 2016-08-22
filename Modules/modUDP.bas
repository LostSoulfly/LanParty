Attribute VB_Name = "modUDP"
Global Const UDPClientPort = 21968

Public bCryptHeader() As Byte
Private UDPBuffer As clsBuffer

Public Sub LoadUDP()
'init UDP Listen/server
On Error GoTo Escape
Set UDPBuffer = New clsBuffer

If Settings.DisableLan Then Exit Sub

With frmMain
    .sckBroadcast.Protocol = sckUDPProtocol
    .sckListen.Protocol = sckUDPProtocol
    
    .sckListen.LocalPort = UDPClientPort
    .sckBroadcast.RemotePort = UDPClientPort
    'AddDebug "Remote Port: " & UDPClientPort
    .sckBroadcast.RemoteHost = "255.255.255.255"  ' Broadcast IP - All network
    'AddDebug "RemoteHost default: " & .sckBroadcast.RemoteHost
    'AddDebug "sckListen binding.."
    .sckListen.Bind                ' UDP must be bound to a port... We shouldn't command 'Listen' like a TCP connection
    'AddDebug "sckBroadcast binding.."
    .sckBroadcast.Bind                ' The broadcast object must be bound also.
    AddDebug "Listening on: " & .sckListen.LocalIP
    Settings.CurrentIP = .sckListen.LocalIP
    'end UDP
End With
Exit Sub
Escape:
If err.Number = 10048 Then
    AddChat "[System] Unable to bind UDP to port " & UDPClientPort & "; port is in use."
    AddChat "[System] Network functions will not work, please restart the program after freeing the port in question."
    
Else
    AddChat "[System] " & "LoadUDP Error: " & err.Number & " " & err.Description
End If

err.Clear
End Sub

Public Sub IncomingDataUDP(ByVal bytesTotal As Long)
Dim Buffer() As Byte
Dim pLength As Long

    frmMain.sckListen.GetData Buffer, vbUnicode, bytesTotal
    UDPBuffer.WriteBytes Buffer()
    If UDPBuffer.Length >= 4 Then
        pLength = UDPBuffer.ReadLong(False)
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= UDPBuffer.Length - 4
        If pLength <= UDPBuffer.Length - 4 Then
            UDPBuffer.ReadLong
            HandleData UDPBuffer.ReadBytes(pLength)

        End If

        pLength = 0
        If UDPBuffer.Length >= 4 Then
            pLength = UDPBuffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop

    UDPBuffer.Trim
    DoEvents
End Sub

Public Function PacketHeader() As Byte()
On Error GoTo Continue
'If UBound(bPacketHeader) > 0 Then PacketHeader = bPacketHeader: Exit Function

Continue:
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteString frmMain.sckListen.LocalIP
    Buffer.WriteString frmMain.sckListen.LocalHostName
    Buffer.WriteString Settings.UniqueID
    
    PacketHeader = Buffer.ReadBytes(Buffer.Length - 1)
    
    'bPacketHeader = Buffer.ReadBytes(Buffer.Length - 1)
    'PacketHeader = bPacketHeader
    
    Set Buffer = Nothing
End Function

Public Function CryptHeader() As Byte()
On Error GoTo Continue
If UBound(bCryptHeader) > 0 Then CryptHeader = bCryptHeader: Exit Function

Continue:
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong LanPacket.LCrypted
    Buffer.WriteConstString Settings.UniqueID
    
    'CryptHeader = Buffer.ReadBytes(Buffer.Length - 1)
    
    'trim off the last piece because we'll be adding it to a different byte array
    bCryptHeader = Buffer.ReadBytes(Buffer.Length - 1)
    CryptHeader = bCryptHeader
    
    Set Buffer = Nothing
End Function

Public Sub SendToAll(ByRef Data() As Byte)
    Dim i As Integer
    
    For i = 1 To UBound(User)
        SendDataToUDP User(i).IP, Data()
    Next
End Sub

Public Sub CryptToAll(ByRef Data() As Byte)
On Error GoTo err

    Dim i As Integer

    For i = 1 To UBound(User)
        If LenB(User(i).UniqueID) > 0 Then
            'AddDebug "Sending cryptto: " & User(i).IP
            SendCryptTo i, Data
        End If
    Next
    
Exit Sub

err:

AddDebug "CryptToAll " & err.Number & ": " & err.Description
    
End Sub

Public Sub CryptToAllAdminSync(ByRef Data() As Byte)
    Dim i As Integer
    
    For i = 1 To UBound(User)
        If LenB(User(i).UniqueID) > 0 Then
            If Not User(i).SyncingAdminList Then
                'AddDebug "Sending cryptto: " & User(i).IP
                SendCryptTo i, Data
            End If
        End If
    Next i

End Sub

Public Sub SendCryptTo(UserIndex As Integer, ByRef Data() As Byte)
    Dim Buffer As New clsBuffer
    
    'Because of the way VB6 handles Byte arrays, I couldn't pass the Data to this sub unless it was passed BY REFERENCE,
    'which means that when the byte array is encrypted (which is done for EACH USER) it is modified all the way back to its origin,
    'so to avoid that I had to copy the array from Data to a newData array after initializing it and CopyMemory should be a
    'very fast api call to do the actual copying. Alternatively I could have done a strconv and maybe encrypted the string.
    'I don't feel like determining which is faster. If you know or find out, let me know!
    
    Dim newData() As Byte
    ReDim newData(UBound(Data))
    CopyMemory newData(0), Data(0), UBound(Data) + 1 'son of a bitch. Without this +1 here, we lose the last character.
    
        If LenB(User(UserIndex).UniqueID) = 0 Then Exit Sub
        Buffer.WriteBytes CryptHeader                                   'write the header with crypt identifier and our uid
        DS2.EncryptByte newData, GetMyRemoteKeyAsByteByIndex(UserIndex)    'encrypt the actual packet with remote unique key
        Buffer.WriteBytes newData                                          'write the encrypted data after the crypt header
        SendDataToUDP User(UserIndex).IP, Buffer.ReadBytes(Buffer.Length) ' send the data!

    Set Buffer = Nothing
    DoEvents
End Sub

Public Sub SendCryptChatAll(Text As String)

    Dim i As Integer
    
    For i = 1 To UBound(User)
        'SendDataToUDP User(i).IP, ChatPacket(Text, True, User(i).UniqueKey)
        'Ugh. I wrote the chat encryption before I decided to encrypt the entire packet inside of an encrypted packet
        'with CryptoToAll/SendCryptTo.
        If LenB(User(i).UniqueID) > 0 Then Exit Sub
        'SendCryptTo i, ChatPacket(Text, True, GetMyRemoteKeyAsByteByIndex(i))
        SendCryptTo i, ChatPacket(Text)
    Next

End Sub

Public Sub BroadcastUDP(ByRef Data() As Byte)
    SendDataToUDP "255.255.255.255", Data()
End Sub


Public Sub SendDataToUDP(ByVal IP As String, ByRef Data() As Byte)
Dim Buffer As clsBuffer
'Dim TempData() As Byte

On Error GoTo wut

If LenB(IP) = 0 Then Exit Sub

        Set Buffer = New clsBuffer
        'TempData = Data
        
        DS2.EncryptByte Data, CryptKey  'Encrypt the data with the generic key, note it's byref
        
        Buffer.PreAllocate 4 + (UBound(Data) - LBound(Data)) + 1    'allocate the new buffer
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1          'write data length to buffer
        Buffer.WriteBytes Data()                                    'write encrypted data

        'max UDP size:  65,507bytes
        'frmMain.sckBroadcast.Close
        frmMain.sckBroadcast.RemoteHost = IP
        'frmMain.sckBroadcast.RemotePort = UDPClientPort
        frmMain.sckBroadcast.SendData Buffer.ToArray()
                
    Set Buffer = Nothing
Exit Sub

wut:

Set Buffer = Nothing

AddDebug err.Number & err.Description
err.Clear

End Sub

Public Function DisableNetwork()

AddDebug "[System] Disabling Network Features.."

With frmMain
    .tmrAdmins.Enabled = False
    .tmrAdminSync.Enabled = False
    .tmrBeacon.Enabled = False
    .tmrPing.Enabled = False
End With

IsSyncingAdmins = False
HasSyncedAdmins = False

DoEvents

frmMain.sckBroadcast.Close
frmMain.sckListen.Close

End Function
