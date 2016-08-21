Attribute VB_Name = "modPChatMgr"
Private m_Forms As Collection

Public Function CountPChatWindows() As Long
    CountPChatWindows = m_Forms.Count
End Function

Public Function GetPChatWindow(PChatID As String) As Form
'Dim frm As Form
If Len(PChatID) = 0 Then Exit Function
Set GetPChatWindow = m_Forms(PChatID)
    
    'frm.Caption = "success!"

End Function

Public Sub RemovePChatWindow(PChatID As String)
On Error Resume Next
    Unload m_Forms(PChatID)
    m_Forms.Remove (PChatID)
End Sub
Public Sub RemoveAllPChatWindows()
Dim i As Long

'we must do it in reverse otherwise the .count will be incorrect as we erase items!
For i = m_Forms.Count To 1 Step -1
    Unload m_Forms(i)
    m_Forms.Remove (i)
Next i

'ZeroMemory m_Forms, LenB(m_Forms)

Set m_Forms = New Collection
End Sub

Public Function PChatWindowExists(PChatID As String) As Boolean
Dim frm As Form
    For Each frm In Forms
        If frm.Tag = PChatID Then
            If frm.Visible = False Then RemovePChatWindow PChatID: Exit Function
            PChatWindowExists = True
            Exit For
        End If
    Next
    Set frm = Nothing
End Function

Public Function PChatNumUsers(PChatID As String) As Long
    
If Not PChatWindowExists(PChatID) Then AddChat "PChat " & PChatID & " doesn't exist!"

PChatNumUsers = GetPChatWindow(PChatID).GetNumUsers

End Function

Public Sub PChatSyncUsers(PChatID As String, UniqueID As String)

If Not PChatWindowExists(PChatID) Then AddChat "PChat " & PChatID & " doesn't exist!"

GetPChatWindow(PChatID).SyncPChatUsers UniqueID

End Sub

Public Function CreatePChatWindow(PChatID As String)
Dim f As New frmPrivateChat
If m_Forms Is Nothing Then Set m_Forms = New Collection
If Len(PChatID) = 0 Then Exit Function
If PChatWindowExists(PChatID) = True Then
    If GetPChatWindow(PChatID).Visible = False Then
        RemovePChatWindow PChatID
    Else
        'todo: bring window to front?
        GetPChatWindow(PChatID).Show
        Exit Function
    End If
Else
    RemovePChatWindow PChatID
End If

f.Caption = "Private Chat"
f.Tag = PChatID
m_Forms.add f, PChatID
f.Show
End Function
