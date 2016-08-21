VERSION 5.00
Begin VB.Form frmCmd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Command Window"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClose 
      Caption         =   "Close After Execution"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton cmdDecline 
      Caption         =   "Decline"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtCmd 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intDelay As Integer
Private Const TotalDelay = 5
Private strCommand As String
Private strArgs As String
Private blShell As Boolean

Private Sub cmdDecline_Click()
Unload Me
End Sub

Private Sub cmdExecute_Click()

If InStr(1, LCase$(txtCmd.Text), "cmd.exe") > 0 Then
    If MsgBox("There exists a reference to the Windows Command-Line Interface. You should not execute this unless you trust the issuer of the command implicitly.", vbYesNo, "Execute?") = vbNo Then Exit Sub
End If

If blShell Then
    Shell strCommand & " " & strArgs
Else
    ExecFile strCommand, strArgs
End If

If chkClose.Value = vbChecked Then Unload Me


End Sub

Private Sub Form_Load()
'cmdExecute.Enabled = False
End Sub

Public Sub SetupCMD(Command As String, Args As String, toShell As Boolean, DelayExec As Boolean, blWait As Boolean)
    
    intDelay = 0
    cmdExecute.Enabled = DelayExec
    strCommand = Command
    strArgs = Args
    txtCmd.Text = strCommand & " " & strArgs
    blShell = toShell
    If blWait Then
        cmdExecute.Enabled = False
        tmrDelay.Enabled = True
    End If
    
End Sub

Private Sub tmrDelay_Timer()

cmdExecute.Caption = "Wait.. " & (TotalDelay - intDelay)

intDelay = intDelay + 1

If intDelay >= TotalDelay + 1 Then cmdExecute.Caption = "Execute": cmdExecute.Enabled = True: tmrDelay.Enabled = False

End Sub
