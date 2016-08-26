VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Debug Stuff"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Ramblings?"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   9495
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInfo_Click()

Text1.Text = ""

AddText "The LanParty Client generates and exchanges a random and unique key of between 28-38 characters for each new connection." & _
"The initial locating of other LanParty clients is done through UDP broadcasts and is encrypted with a default key. This is the default encryption key: " & _
vbNewLine & CStr(CryptKey) & vbNewLine & _
"However, subsequent communication is encrypted twice, once with the default key, and once with the unique key that was exchanged with the intended recipient. The advantage to this is that each " & _
"packet is encrypted differently depending on who it is being sent to, because each client has a unique key for and from every other client, " & _
"even in the case of a client receiving a packet it was not supposed to, it wouldn't be able to decrypt it as the key that it has does not match " & _
"what the packet was encrypted with."
AddText "Perhaps some day I'll re-write this program (If it's successful and useful) in a more modern language and utilize public-key cryptography for even better security and set up a sort of mesh network."
AddText "I have a lot of ideas of how awesome a LanParty client I could write including file sharing and more.."

RefreshDebug

End Sub

Private Sub cmdRefresh_Click()
    Text1.Text = ""
    RefreshDebug
End Sub

Private Sub Form_Load()

    Text1.Text = ""
    RefreshDebug

End Sub

Private Sub AddText(Text As String, Optional NL As Boolean = True)

    Text1.Text = Text1.Text & Text & IIf(NL, vbNewLine, "")

End Sub

Private Sub RefreshDebug()

Dim strSep As String
strSep = "---------------------------------------------"

AddText strSep
AddText "User Name" & vbTab & "HardwareID" & vbTab & "CurrentIP"
AddText Settings.UserName & vbTab & vbTab & Settings.UniqueID & vbTab & Settings.CurrentIP
AddText "IsSyncing: " & IsSyncingAdmins & vbTab & " HasSyncedAdmins: " & HasSyncedAdmins _
& vbTab & " tmrAdmin: " & frmMain.tmrAdmins.Enabled

AddText strSep

Dim i As Integer

For i = 1 To UBound(User)
AddText GetUserNameByIndex(i) & vbTab & vbTab & User(i).UniqueID & vbTab & User(i).IP & vbTab & "Ver:" & User(i).AppVersion & vbTab & "CompName: " & User(i).CompName
AddText "MyUniqueKey: " & User(i).MyUniqueKey & vbNewLine & "Their UniqueKey: " & User(i).UniqueKey
AddText "Playing: " & GetGameName(GameIndexByUID(User(i).CurrentlyPlaying)) & " (" & User(i).CurrentlyPlaying & ")" & vbTab & "LanAdmin: " & User(i).LanAdmin & vbTab & "LastHeard: " & User(i).LastHeard & vbTab & "SyncingVotes: " & User(i).SyncingVotes & vbTab & "HasSyncedVotes: " & User(i).HasSyncedVotes
AddText strSep
Next i

End Sub
