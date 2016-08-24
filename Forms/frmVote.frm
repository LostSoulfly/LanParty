VERSION 5.00
Begin VB.Form frmVote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vote"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Submit Vote"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   7095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   7095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   7095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   7095
   End
   Begin VB.Label lblText 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()
'send vote packet to all
'no verification :(
Dim Vote As Integer
Dim VoteIndex As Long

If Option1.Value = True Then Vote = 1
If Option2.Value = True Then Vote = 2
If Option3.Value = True Then Vote = 3
If Option4.Value = True Then Vote = 4
VoteIndex = modVote.VoteExists(Me.Tag)

If VoteIndex = -1 Then Exit Sub

'If Vote(VoteIndex).AdminVote = True Then
    'cast my vote in my own vote array
    modVote.AddUserVote Me.Tag, Settings.UniqueID, Vote

    'then send it to everyone!
    CryptToAll VotePacket(Me.Tag, Vote)
        
    DoEvents
    Unload Me
    frmVoteList.UpdateLabels
End Sub

Public Sub RefreshVote()
Dim VoteIndex As Long
VoteIndex = VoteExists(Me.Tag)
lblText.Caption = Vote(VoteIndex).Text
Option1.Caption = Vote(VoteIndex).Option1(0)
Option2.Caption = Vote(VoteIndex).Option2(0)
Option3.Caption = Vote(VoteIndex).Option3(0)
Option4.Caption = Vote(VoteIndex).Option4(0)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

