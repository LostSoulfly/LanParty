VERSION 5.00
Begin VB.Form frmVoteList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vote List"
   ClientHeight    =   4620
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstVote 
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblOption4 
      Caption         =   "<Option4>"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   5535
   End
   Begin VB.Label lblOption3 
      Caption         =   "<Option3>"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Label lblOption2 
      Caption         =   "<Option2"
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label lblOption1 
      Caption         =   "<Option1>"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label lblText 
      Caption         =   "<Text>"
      Height          =   1575
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.Menu mnuVoteMenu 
      Caption         =   "Vote Menu"
      Begin VB.Menu mnuSync 
         Caption         =   "Re-Sync Votes"
      End
      Begin VB.Menu mnuNewVote 
         Caption         =   "Start New Vote"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Vote List"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmVoteList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub RefreshList()
'refresh the list, lel
Dim i As Long
Dim OldIndex As Long
If lstVote.ListIndex > 0 Then OldIndex = lstVote.ItemData(lstVote.ListIndex)
lstVote.Clear

For i = 1 To UBound(Vote)

If Vote(i).AdminVote Then
    With Vote(i)
        If .AdminUID = Settings.UniqueID Then
            .Title = "Admin Vote: " & GetUserName(.AdminUID) & " on PC " & frmMain.sckListen.LocalHostName
            .Text = "Do you vote for or against " & GetUserName(.AdminUID) & " on PC " & frmMain.sckListen.LocalHostName & " becoming a LAN Admin?"
        Else
            If UserIndexByUID(.AdminUID) > 0 Then
                .Title = "Admin Vote: " & GetUserName(.AdminUID) & " on PC " & User(UserIndexByUID(.AdminUID)).CompName
                .Text = "Do you vote for or against " & GetUserName(.AdminUID) & " on PC " & User(UserIndexByUID(.AdminUID)).CompName & " becoming a LAN Admin?"
            Else
                .Title = "Admin Vote: <Offline/Unknown User> UID: " & .AdminUID
                .Text = "Do you vote for or against <Offline/Unknown User> on PC <Offline/Unknown User> UID: (" & .AdminUID & ") becoming a LAN Admin?"
            End If
        End If
    End With
End If

    lstVote.AddItem i & ": " & Vote(i).Title
    lstVote.ItemData(lstVote.ListCount - 1) = i
    If OldIndex = i - 1 Then lstVote.ListIndex = i - 1

Next

UpdateLabels

End Sub

Public Sub UpdateLabels()
On Error Resume Next

If lstVote.ListCount = 0 Then Exit Sub

With Vote(lstVote.ItemData(lstVote.ListIndex))
    
    lblText.Caption = .Text
    
    Dim Op1 As String
    Dim Op2 As String
    Dim Op3 As String
    Dim Op4 As String
    
    modVote.CalculatePercentage Vote(lstVote.ItemData(lstVote.ListIndex)).VoteID, Op1, Op2, Op3, Op4
    
    lblOption1.Caption = "(" & Op1 & "%) " & .Option1(0)
    lblOption2.Caption = "(" & Op2 & "%) " & .Option2(0)
    lblOption3.Caption = "(" & Op3 & "%) " & .Option3(0)
    lblOption4.Caption = "(" & Op4 & "%) " & .Option4(0)
    

End With

Me.Width = 8775

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Width = 3105
RefreshList
'lstVote.ListIndex = lstVote.ListCount - 1
'UpdateLabels
End Sub

Private Sub lstVote_Click()

UpdateLabels

End Sub

Private Sub lstVote_DblClick()
ShowVoteWindow (lstVote.ItemData(lstVote.ListIndex))
End Sub

Private Sub mnuClose_Click()
Me.Visible = False
End Sub

Private Sub mnuNewVote_Click()
Dim Text As String
Dim Title As String
Dim Option1 As String
Dim Option2 As String
Dim Option3 As String
Dim Option4 As String


MsgBox "This is a ghetto way to do this because I didn't want to spend any more time on it, sorry!", vbOKOnly, "New Vote"
Title = InputBox("What's the title of your poll?", "Poll Title")

Text = InputBox("What is the question of your poll?", "Poll Question")
Option1 = InputBox("First poll option: (The first two must not be left blank; enter nothing to cancel.)", "Option 1")
Option2 = InputBox("Second poll option: (The first two must not be left blank; enter nothing to cancel.)", "Option 2")
Option3 = InputBox("Third poll option: (enter nothing to cancel.)", "Option 3")
Option4 = InputBox("Final poll option: (enter nothing to cancel.)", "Option 4")

If Len(Option1$) = 0 Then Exit Sub
If Len(Option2$) = 0 Then Exit Sub

Text = Settings.UserName & " asks: " & Trim$(Text)

Dim VoteID As String

'generate a new voteid with the supplied vote info.
VoteID = modVote.NewVote(vbNullString, Settings.UniqueID, Title, Text, Option1, Option2, Option3, Option4)

If Len(VoteID$) < 19 Then MsgBox "Something went wrong; an incorrect VoteId was generated!?", vbOKOnly, "What?": Exit Sub

'send the new vote packet to all.
CryptToAll NewVotePacket(VoteID, Trim$(Title), Text, Trim$(Option1), Trim$(Option2), Trim$(Option3), Trim$(Option4))
DoEvents
RefreshList
'lstVote.ListIndex = lstVote.ListCount - 1
'UpdateLabels
End Sub

Private Sub mnuRefresh_Click()
RefreshList
'lstVote.ListIndex = lstVote.ListCount - 1
'UpdateLabels
End Sub

Private Sub mnuSync_Click()

    SyncAllVotes , True

End Sub
