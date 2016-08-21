Attribute VB_Name = "modProcesses"

Public Proc As Collection
Public blSentProcesses As Boolean

Public Sub GetFullProcesses(Optional SendToServer As Boolean = False)

Set Proc = Nothing
Set Proc = New Collection

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String
Dim pName As String
Dim pPid As Long
    Set cTemp = cProc.Process_Enumerate
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
            pName = sUData(0)  'name
            pPid = Val(sUData(1))  'ID
            AddToCollection pName, pPid
    Next vItem

On Error GoTo 0

End Sub

Public Sub GetNewProcesses()

If Proc Is Nothing Then Set Proc = New Collection

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String
Dim pName As String
Dim pPid As Long
    Set cTemp = cProc.Process_Enumerate
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
            pName = sUData(0)  'name
            pPid = Val(sUData(1))  'ID
            If ProcessInCol(pName, pPid) = True Then
                Proc.Remove pName & pPid 'remove the ones that exist currently
            Else
                'SendData NewProcess(pName, pPid)
            End If
    Next vItem

    For Each vItem In Proc
            pName = vItem(0)  'name
            pPid = Val(vItem(1))  'ID
        SendData DelProcess(pName, pPid)
    Next vItem

'erase the process list and start fresh
GetFullProcesses

'compare it to the collection and find the new processes, process the new ones
'and then we erase the collection and re-add all processes

End Sub

Public Function ProcessCount(Name As String) As Long

If Proc Is Nothing Then Set Proc = New Collection

Dim Count As Long
    For Each Item In Proc
        If UCase(Item(0)) = UCase(Name) Then
            Count = Count + 1
        End If
    Next
    
    ProcessCount = Count
    
End Function

Public Function GetPIDFromName(Name As String) As Long
    
If Proc Is Nothing Then Set Proc = New Collection
    
    For Each Item In Proc
        If UCase(Item(0)) = UCase(Name) Then
            GetPIDFromName = Item(1)
            Exit Function
        End If
    Next
    
    GetPIDFromName = 0
    
End Function

Public Function TerminateAllByName(Name As String) As Boolean

If Proc Is Nothing Then Set Proc = New Collection

Dim PID As Long
PID = GetPIDFromName(Name)
    Do While PID > 0
        TerminateAllByName = True
        cProc.Process_Terminate PID
        GetNewProcesses 'refresh the process collection
        DoEvents
        PID = GetPIDFromName(Name)
    Loop

End Function

Private Function AddToCollection(Name As String, PID As Long) As Boolean
Dim strProc(1) As String

If Proc Is Nothing Then Set Proc = New Collection

    If ProcessInCol(Name, PID) = True Then Exit Function
    
    strProc(0) = Name
    strProc(1) = PID
    
    Proc.Add strProc, Name & PID
    
End Function
