Attribute VB_Name = "modAdmin"
Option Explicit
Public HasSyncedAdmins As Boolean
Public IsSyncingAdmins As Boolean


Public Function CalculateAdminLists()
Dim UID As String
Dim i As Integer
Dim ii As Integer

For i = 1 To UBound(User)
    'for each user.. loop through their admin list
    With User(i)
    
        For ii = 0 To UBound(.AdminList)
            UID = .AdminList(ii)
            If Len(UID) > 0 Then
                CalculateIsUserAdmin UID
                'If CalculateIsUserAdmin(UID) Then
                    'SetUserLanAdmin UID, True
                'End If
            End If
        
        Next ii
        
    'reset this bool so they can sync again if necessary
    .SyncingAdminList = False
    End With

Next i

AddDebug "Finished syncing AdminLists.."

HasSyncedAdmins = True
IsSyncingAdmins = False

'SyncAllVotes
frmMain.tmrVotesSync.Enabled = True

End Function

Public Function MostAdminListSynced() As Boolean
Dim i As Integer
Dim Count As Integer

For i = 1 To UBound(User)
    If User(i).SyncingAdminList Then Count = Count + 1
Next i

AddDebug "Admin Synced Percent: " & ((Count / GetUserCount) * 100)

If ((Count / GetUserCount) * 100) >= 75 Then
    MostAdminListSynced = True
End If

End Function

Private Function CalculateIsUserAdmin(UID As String) As Boolean
Dim i As Integer
Dim ii As Integer
Dim SeeAsAdmin As Integer
Dim OnlyOnce As Boolean
For i = 1 To UBound(User)
    'for each user.. loop through their admin list
    With User(i)
    OnlyOnce = False
        For ii = 0 To UBound(.AdminList)    'We check each array item, increment the SeeAs counter, and then erase the UID
                                            'so that we won't check the same user in each adminlist.
            If .AdminList(ii) = UID Then
                If OnlyOnce = False Then
                    SeeAsAdmin = SeeAsAdmin + 1
                    OnlyOnce = True
                End If
                
                .AdminList(ii) = vbNullString
            End If
        Next ii
    End With
Next i

'FormatNumber(UBound(Vote(VoteIndex).Option1) / TotalVotes * 100)
If GetUserCount = 1 Then Exit Function
AddDebug GetUserName(UID) & " admin sync percentage: " & ((SeeAsAdmin / GetUserCount) * 100)
If ((SeeAsAdmin / GetUserCount) * 100) > 75 Then
    'there is a 75% vote FOR this user being an admin.
    SetUserLanAdmin UID, True
End If

End Function

Public Function RemoveFromAdminList(UserIndex As Integer, UID As String)
Dim i As Integer
    With User(UserIndex)
        
        For i = 0 To UBound(.AdminList)
            If .AdminList(i) = UID Then .AdminList(i) = vbNullString
        Next
    
    End With

End Function

Public Function AddToAdminList(UserIndex As Integer, UID As String)
Dim i As Integer
    With User(UserIndex)
    
        For i = 0 To UBound(.AdminList)
                If Len(.AdminList(i)) <> 0 Then
                    If .AdminList(i) = UID Then Exit Function
                Else
                    .AdminList(i) = UID
                Exit Function
            End If
        Next
        
        ReDim .AdminList(UBound(.AdminList) + 1)
        
        .AdminList(i) = UID
    
    End With

End Function

Public Sub SetUserLanAdmin(UID As String, isAdmin As Boolean)

    If User(UserIndexByUID(UID)).LanAdmin = isAdmin Then Exit Sub

    If UID = Settings.UniqueID Then
        Settings.LanAdmin = True
        UpdateAdminMenus (True)
    Else
        User(UserIndexByUID(UID)).LanAdmin = isAdmin
    End If
    
    AddChat "<--- " & GetUserName(UID) & " has been elected a LanAdmin --->"
    
End Sub

Public Sub UpdateAdminMenus(isAdmin As Boolean)

With frmMain

    .mnuChangeUserName.Visible = isAdmin
    '.mnuExecAll.Visible = isAdmin
    .mnuExecute.Visible = isAdmin
    .mnuGlobalMuteMnu.Visible = isAdmin
    '.mnuGlobalUnMute.Visible = isAdmin
    .mnuKick.Visible = isAdmin
    '.mnuSuggest.Visible = isAdmin
    .mnuSuggestAll.Visible = isAdmin
    .mnuLaunchAll.Visible = isAdmin
    .mnuFreeze.Visible = isAdmin
    .mnuMsg.Visible = isAdmin
    
End With

End Sub
