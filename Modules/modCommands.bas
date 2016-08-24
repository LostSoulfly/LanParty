Attribute VB_Name = "modCommands"
Option Explicit


Public Sub ShowNewCmdWindow(Command As String, Args As String, blShell As Boolean, ExecIfAdmin As Boolean, blWait As Boolean, UserIndex As Integer)

If ExecIfAdmin And Settings.AllowCommands Then
    'exec the commands
    AddDebug "Allowing command because LanAdmin: " & Command & " " & Args
    
        If blShell Then
            Shell Command & " " & Args
        Else
            ExecFile Command, Args
        End If
    
    Exit Sub
End If

Dim f As New frmCmd

Load f
f.Caption = f.Caption & " - From " & GetUserNameByIndex(UserIndex)
f.SetupCMD Command, Args, blShell, True, blWait
f.Visible = True

End Sub
