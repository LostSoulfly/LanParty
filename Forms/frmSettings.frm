VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10110
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   10110
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmScript 
      Caption         =   "Script Options"
      Height          =   1095
      Left            =   6840
      TabIndex        =   40
      Top             =   1680
      Width           =   3135
      Begin VB.CheckBox chkAllowScriptDL 
         Caption         =   "Allow File Downloads In Scripts"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkAllowScriptExec 
         Caption         =   "Allow File Execution In Scripts"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkAllowScripts 
         Caption         =   "Allow Scripts"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save && Close"
      Height          =   495
      Left            =   8400
      TabIndex        =   17
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text Options"
      Height          =   1215
      Left            =   3600
      TabIndex        =   24
      Top             =   2160
      Width           =   3135
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   315
         Left            =   2280
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.PictureBox picChatBGColor 
         Height          =   300
         Left            =   2640
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   31
         Top             =   320
         Width           =   375
      End
      Begin VB.PictureBox picIconTextColor 
         Height          =   300
         Left            =   1750
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   30
         Top             =   700
         Width           =   375
      End
      Begin VB.PictureBox picChatColor 
         Height          =   300
         Left            =   1750
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   29
         Top             =   320
         Width           =   375
      End
      Begin VB.TextBox txtIconSize 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "8"
         Top             =   720
         Width           =   375
      End
      Begin VB.VScrollBar vsIconSize 
         Height          =   255
         Left            =   1320
         Max             =   72
         Min             =   6
         TabIndex        =   27
         Top             =   720
         Value           =   6
         Width           =   255
      End
      Begin VB.TextBox txtChatSize 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "8"
         Top             =   330
         Width           =   375
      End
      Begin VB.VScrollBar vsChatSize 
         Height          =   255
         Left            =   1320
         Max             =   72
         Min             =   6
         TabIndex        =   25
         Top             =   330
         Value           =   6
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "BG:"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Icon Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Chat Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog dlgFileOpen 
      Left            =   2640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Network Options"
      Height          =   1455
      Left            =   6840
      TabIndex        =   21
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox chkVersion 
         Caption         =   "Version-specific Encryption Key"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   $"frmSettings.frx":151DE
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox chkChatType 
         Caption         =   "Broadcast, Not Direct"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   $"frmSettings.frx":15288
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chkAllowCommands 
         Caption         =   "Allow Admin to Run Commands"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   $"frmSettings.frx":1539B
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkShareLanGame 
         Caption         =   "Share Game Status"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkDisableLan 
         Caption         =   "Disable ALL Network Features"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "LanParty Options"
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox chkDebug 
         Caption         =   "Enable Debug Output"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CheckBox chkScanStartup 
         Caption         =   "Scan For Games At Startup"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "Auto Check For Updates"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Check for updates at program startup."
         Top             =   2160
         Width           =   3015
      End
      Begin VB.PictureBox picLanBG 
         Height          =   300
         Left            =   2880
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   33
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkJason 
         Caption         =   "I'm Jason Copple."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "This provides extra information where necessary to make Jason Copple more comfortable."
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox chkShowIcons 
         Caption         =   "Load and Show Icons In Launcher"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   $"frmSettings.frx":1544C
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox chkMinimize 
         Caption         =   "Minimize LanParty Upon Game Launch"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Minimize the LanParty client upon launching a game."
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   300
         Left            =   2400
         TabIndex        =   22
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtBG 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkMonitorGame 
         Caption         =   "Monitor Currently Launched Game"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   $"frmSettings.frx":154D9
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "If you don't know what this is, you should probably leave."
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "LanParty Background Image:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "UserName:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LanChat Options"
      Height          =   1935
      Left            =   3600
      TabIndex        =   18
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox chkAcceptPrivateChat 
         Caption         =   "Auto-Accept Private Chats"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "If you receive an unsolicited Private Chat message a new window will be opened, displaying it."
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkShowStatus 
         Caption         =   "Show User's Status Next To Name"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chkLogChat 
         Caption         =   "Save Chat Log On Close"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox chkDisableChat 
         Caption         =   "Disable Chat"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkDockChat 
         Caption         =   "Dock Chat To Main Window"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "This keeps the LanChat window stuck beneath the main window."
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkShowChat 
         Caption         =   "Show Chat Window"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Show the LanChat. It will remain connected in the background."
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldName As String
Public blCancel As Boolean


Private Sub chkAllowCommands_Click()
If chkAllowCommands.Value = vbUnchecked Then
    MsgBox "This option is disabled by default and does not prevent a Host from *recommending* commands to be executed on your computer." & vbNewLine & vbNewLine & _
    "However, you make the decision whether to allow the recommended commands to be run.", vbInformation, "Admin Commands"
    
End If

End Sub

Private Sub chkDebug_Click()
'chkJason.Value = vbChecked
End Sub

Private Sub chkDisableLan_Click()
    If chkDisableLan.Value = vbChecked Then
        chkShowChat.Value = vbUnchecked
        chkDockChat.Value = vbUnchecked
        chkShareLanGame.Value = vbUnchecked
        chkAllowCommands.Value = vbUnchecked
        chkMonitorGame.Value = vbUnchecked
    End If
End Sub

Private Sub chkMonitorGame_Click()
chkShareLanGame.Value = chkMonitorGame.Value
End Sub

Private Sub chkShareLanGame_Click()
chkMonitorGame.Value = chkShareLanGame.Value
End Sub

Private Sub cmdCancel_Click()

If Len(Settings.UserName$) < 3 Then
    MsgBox "You need to enter and save a username before using this program!"
    txtUser.SetFocus
    Exit Sub
End If

blCancel = True
Me.Visible = False
'Unload Me
End Sub

Private Sub cmdOpen_Click()
With dlgFileOpen
    .DefaultExt = ""
    .InitDir = App.Path
    .DialogTitle = "Select a Background Image.."
    .Filter = "Image Files" & "|" & "*gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.wmf;*.emf;*.png;*.tga" & "|" & _
    "Bitmaps" & "|" & "*.bmp" & "|" & "GIFs" & "|" & "*.gif" & "|" & _
                            "Icons/Cursors" & "|" & "*.ico;*.cur" & "|" _
                            & "JPGs" & "|" & "*.jpg;*.jpeg" & "|" _
                            & "Meta Files" & "|" & "*.wmf;*.emf" & "|" _
                            & "PNGs" & "|" & "*.png" & "|" _
                            & "TGAs (Targa)" & "|" & "*.tga" & "|" _
                            & "TIFFs" & "|" & "*.tiff" & "|" _
                            & "All Files" & "|" & "*.*"
    .ShowOpen
    If FileExists(.FileName) Then txtBG.Text = .FileName
End With
End Sub

Private Sub cmdReset_Click()

    txtChatSize.Text = 8
    txtIconSize.Text = 8
    picChatBGColor.BackColor = 16777215
    picChatColor.BackColor = 0
    picIconTextColor.BackColor = 0
    picLanBG.BackColor = -2147483633

End Sub

Private Sub cmdSave_Click()
If Trim$(txtUser.Text) = "" Then
    MsgBox "You must enter a username!", vbCritical, "Don't be difficult."
    Exit Sub
End If

    If VerifyKeyAsString(Trim$(txtUser.Text)) Then
        Settings.UserName = Trim$(txtUser.Text)
    Else
        MsgBox "Your username contains illegal characters!", vbInformation, "Try Again"
        txtUser.SetFocus
        Exit Sub
    End If

If Not IIf(chkVersion.Value, True, False) = Settings.SameVersion Then MsgBox "Changing the 'Version-specific Encryption' requires a restart!" & vbNewLine & _
"Also, only other users that have this enabled AND have the same exact version of the LanParty as you will be able to connect!", vbInformation, "Restart Required"

SetSettings
SaveSettings

If Not OldName = txtUser.Text And Len(OldName$) > 0 Then
MsgBox "I will attempt to notify other users of your name change one time." & vbNewLine & vbNewLine & _
            "However, due to the connectionless nature of UDP (The protocol we are using here) I cannot guarantee all clients " _
            & "will receive the name change. " & vbNewLine & "I mean, I totally could make it work that way by having remote clients confirm.. Well, nevermind." & _
            vbNewLine & vbNewLine & "Deal with it.", vbInformation, "Name change? No guarantees."

CryptToAll NameChangePacket

End If

DoEvents
Me.Visible = False
End Sub

Private Sub Form_Load()
DisplaySettings
OldName = txtUser.Text
End Sub


Private Sub DisplaySettings()

chkDebug.Value = IIf(Settings.blDebug, vbChecked, vbUnchecked)
txtUser.Text = Settings.UserName
txtBG.Text = Settings.BackgroundPath
chkShowIcons.Value = IIf(Settings.ShowIcons, vbChecked, vbUnchecked)
chkMinimize.Value = IIf(Settings.MinimizeAfterLaunch, vbChecked, vbUnchecked)
chkMonitorGame.Value = IIf(Settings.MonitorGame, vbChecked, vbUnchecked)
chkShowChat.Value = IIf(Settings.ShowChat, vbChecked, vbUnchecked)
chkDockChat.Value = IIf(Settings.DockChat, vbChecked, vbUnchecked)
chkChatType.Value = IIf(Settings.AltChatType, vbChecked, vbUnchecked)
'chkDisableChat = IIf(Settings.DisableChat, vbChecked, vbUnchecked)
chkAllowCommands.Value = IIf(Settings.AllowCommands, vbChecked, vbUnchecked)
chkLogChat.Value = IIf(Settings.LogChat, vbChecked, vbUnchecked)
chkShowStatus.Value = IIf(Settings.ShowStatus, vbChecked, vbUnchecked)
'chkShareLanGame = IIf(Settings.MinimizeAfterLaunch, vbChecked, vbUnchecked)
chkDisableLan.Value = IIf(Settings.DisableLan, vbChecked, vbUnchecked)
chkJason.Value = IIf(Settings.Jason, vbChecked, vbUnchecked)
chkUpdate.Value = IIf(Settings.AutoUpdate, vbChecked, vbUnchecked)
chkScanStartup.Value = IIf(Settings.ScanAtStartup, vbChecked, vbUnchecked)
chkVersion.Value = IIf(Settings.SameVersion, vbChecked, vbUnchecked)
chkAcceptPrivateChat.Value = IIf(Settings.AcceptPrivateChat, vbChecked, vbUnchecked)
chkAllowScripts.Value = IIf(Settings.AllowScripts, vbChecked, vbUnchecked)
chkAllowScriptExec.Value = IIf(Settings.ScriptExecute, vbChecked, vbUnchecked)
chkAllowScriptDL.Value = IIf(Settings.ScriptDownload, vbChecked, vbUnchecked)

txtChatSize.Text = Settings.ChatTextSize: vsChatSize.Value = Settings.ChatTextSize
txtIconSize.Text = Settings.IconTextSize: vsIconSize.Value = Settings.IconTextSize
picChatBGColor.BackColor = Settings.ChatBGColor
picChatColor.BackColor = Settings.ChatTextColor
picIconTextColor.BackColor = Settings.IconTextColor
picLanBG.BackColor = Settings.LanBGColor

End Sub

Private Sub SetSettings()
    
    Settings.AltChatType = chkChatType.Value
    Settings.UserName = Trim$(txtUser.Text)
    Settings.BackgroundPath = FormatToLocalPath(Trim$(txtBG.Text))
    Settings.blDebug = chkDebug.Value
    Settings.ShowIcons = chkShowIcons.Value
    Settings.MinimizeAfterLaunch = chkMinimize.Value
    Settings.MonitorGame = chkMonitorGame.Value
    Settings.ShowChat = chkShowChat.Value
    Settings.DockChat = chkDockChat.Value
    Settings.LogChat = chkLogChat.Value
    Settings.ShowStatus = chkShowStatus.Value
    Settings.Jason = chkJason.Value
    Settings.AutoUpdate = chkUpdate.Value
    Settings.ScanAtStartup = chkScanStartup.Value
    Settings.SameVersion = chkVersion.Value
    Settings.DisableLan = chkDisableLan.Value
    'chkDisableChat.Value = IIf(Settings.DisableChat, vbChecked, vbUnchecked)
    Settings.AllowCommands = chkAllowCommands.Value
    Settings.AllowScripts = chkAllowScripts.Value
    Settings.ScriptDownload = chkAllowScriptExec.Value
    Settings.ScriptExecute = chkAllowScriptDL.Value

    
    'chkShareLanGame.Value = IIf(Settings.MinimizeAfterLaunch, vbChecked, vbUnchecked)
    'chkDisableLan.Value = IIf(Settings.MinimizeAfterLaunch, vbChecked, vbUnchecked)
    Settings.AcceptPrivateChat = chkAcceptPrivateChat.Value
    Settings.ChatTextSize = Val(txtChatSize.Text)
    Settings.IconTextSize = Val(txtIconSize.Text)
    Settings.ChatBGColor = Val(picChatBGColor.BackColor)
    Settings.ChatTextColor = Val(picChatColor.BackColor)
    Settings.IconTextColor = Val(picIconTextColor.BackColor)
    Settings.LanBGColor = Val(picLanBG.BackColor)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Visible = True Then Cancel = 1
End Sub

Private Sub picChatBGColor_Click()
Dim Color As Long
Color = ShowColorDialog

If Color = -1 Then Exit Sub

picChatBGColor.BackColor = Color

End Sub

Private Sub picChatColor_Click()
Dim Color As Long
Color = ShowColorDialog

If Color = -1 Then Exit Sub

picChatColor.BackColor = Color

End Sub

Private Sub picIconTextColor_Click()
Dim Color As Long
Color = ShowColorDialog

If Color = -1 Then Exit Sub

picIconTextColor.BackColor = Color

End Sub

Private Sub picLanBG_Click()

Dim Color As Long
Color = ShowColorDialog

If Color = -1 Then Exit Sub

picLanBG.BackColor = Color

End Sub

Private Sub txtChatSize_Change()
If Not IsNumeric(txtChatSize.Text) Then txtChatSize.Text = 8
End Sub

Private Sub txtIconSize_Change()
If Not IsNumeric(txtIconSize.Text) Then txtIconSize.Text = 8
End Sub

Private Sub vsChatSize_Change()
txtChatSize.Text = vsChatSize.Value
End Sub

Private Sub vsIconSize_Change()
txtIconSize.Text = vsIconSize.Value
End Sub
