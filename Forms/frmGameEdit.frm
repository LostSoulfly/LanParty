VERSION 5.00
Begin VB.Form frmGameEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Editor"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4530
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPlayers 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   25
      Text            =   "0"
      ToolTipText     =   "Multiple games can be separated by semicolons ( ; )"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Game"
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkCommand 
      Caption         =   "Run As Command"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox chkMonitor 
      Caption         =   "Monitor Game Running"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CheckBox chkInstallFirst 
      Caption         =   "Require Install First"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtInstallerPath 
      Height          =   285
      Left            =   1200
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1920
      Width           =   5535
   End
   Begin VB.CheckBox chkGameUID 
      Caption         =   "Gen New GUID"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdLocalPath 
      Caption         =   "Convert To Local Paths"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "= Update ="
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "+ New +"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "> Next >"
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton chkPrevious 
      Caption         =   "< Previous <"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtCMDArgs 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox txtMonitorEXE 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "Monitor EXE(s)"
      ToolTipText     =   "Multiple games can be separated by semicolons ( ; )"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtGameEXE 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtIconPath 
      Height          =   285
      Left            =   1200
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox txtEXEPath 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblPlayers 
      Caption         =   "Max Players:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   1560
      Picture         =   "frmGameEdit.frx":0000
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   5400
      Picture         =   "frmGameEdit.frx":0442
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4440
      Picture         =   "frmGameEdit.frx":0884
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   315
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   6840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label label8 
      Caption         =   "Installer Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Game Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   6720
      X2              =   6720
      Y1              =   3120
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   5880
      X2              =   6720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   5880
      X2              =   6720
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   5880
      Y1              =   3120
      Y2              =   3960
   End
   Begin VB.Label Label5 
      Caption         =   "Cmd Args:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Monitor EXE:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Game EXE:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Icon Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "EXE Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Drag and Drop Game File Here:"
      Height          =   495
      Left            =   4680
      TabIndex        =   21
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frmGameEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GameIndex As Integer

Private Sub UpdateGame(Index As Integer, Optional blDelete As Boolean = False)
AddDebug "UpdateGame: " & Index & " blDelete: " & blDelete
If Not blDelete Then
    With Game(Index)
        .EXEPath = Trim$(txtEXEPath.Text)
        .GameEXE = Trim$(txtGameEXE.Text)
        .CommandArgs = Trim$(txtCMDArgs.Text)
        .IconPath = Trim$(txtIconPath.Text)
        .MonitorRunning = IIf(chkMonitor.Value = vbChecked, True, False)
        .MonitorEXE = Trim$(txtMonitorEXE.Text)
        .Name = Trim$(txtName.Text)
        .InstallerPath = Trim$(txtInstallerPath.Text)
        .InstallFirst = IIf(chkInstallFirst.Value = vbChecked, True, False)
        .GameType = IIf(chkCommand.Value = vbChecked, 1, 0)
        .MaxPlayers = IIf(IsNumeric(txtPlayers.Text), CInt(txtPlayers.Text), 0)
        'AddDebug "UpdateGame - CalcGameUID"
        CalcGameUID Index
    End With
Else
    With Game(Index)
        .EXEPath = ""
        .GameEXE = ""
        .CommandArgs = ""
        .IconPath = ""
        .MonitorRunning = False
        .MonitorEXE = ""
        .Name = ""
        .InstallerPath = ""
        .InstallFirst = False
        .GameType = 0
        .GameUID = ""
        .MaxPlayers = 0
    End With
End If
End Sub

Public Sub RefreshGame(Index As Integer)

With Game(Index)
    
    txtEXEPath.Text = .EXEPath
    txtGameEXE.Text = .GameEXE
    txtCMDArgs.Text = .CommandArgs
    txtIconPath.Text = .IconPath
    chkMonitor.Value = IIf(.MonitorRunning, vbChecked, vbUnchecked)
    txtMonitorEXE.Text = .MonitorEXE
    txtName.Text = .Name
    txtInstallerPath.Text = .InstallerPath
    chkInstallFirst.Value = IIf(.InstallFirst, vbChecked, vbUnchecked)
    chkCommand.Value = IIf(.GameType = 0, vbUnchecked, vbChecked)
    Me.Caption = "Game Editor - " & Index & " - " & .Name
    txtPlayers.Text = .MaxPlayers
    
End With

UpdateCmdChk

GameIndex = Index
End Sub

Private Sub UpdateCmdChk()

If chkCommand.Value = vbChecked Then
    Label1.Caption = "Working Dir:"
    Label3.Caption = "File Path:"
    Label7.Caption = "Cmd Name:"
    chkInstallFirst.Enabled = False
    txtInstallerPath.Enabled = False
Else
    Label1.Caption = "EXE Path:"
    Label3.Caption = "Game EXE:"
    Label7.Caption = "Game Name:"
    chkInstallFirst.Enabled = True
    txtInstallerPath.Enabled = True
End If

End Sub

Private Sub chkCommand_Click()
    UpdateCmdChk
End Sub

Private Sub chkPrevious_Click()

GameIndex = GameIndex - 1

If GameIndex <= LBound(Game) Then GameIndex = LBound(Game) + 1

RefreshGame GameIndex
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you certain you wish to delete this game?", vbYesNoCancel, "Delete Game") = vbYes Then
    UpdateGame GameIndex, True
    SaveGames
    DoEvents
    InitializeGameArray True
    UpdateIconList True
    DoEvents
    
    cmdNext_Click
End If
End Sub

Private Sub cmdLocalPath_Click()

    txtEXEPath.Text = FormatToLocalPath(txtEXEPath.Text)
    txtGameEXE.Text = FormatToLocalPath(txtGameEXE.Text)
    txtCMDArgs.Text = FormatToLocalPath(txtCMDArgs.Text)
    txtIconPath.Text = FormatToLocalPath(txtIconPath.Text)
    txtMonitorEXE.Text = txtMonitorEXE.Text
    txtInstallerPath.Text = FormatToLocalPath(txtInstallerPath.Text)

End Sub

Private Sub cmdNew_Click()
'scan through games to find an empty one first
'then if no empty ones found, save as a new one
'expand the game array and pass the index to updategame
Dim NewIndex As Integer

NewIndex = UBound(Game) + 1
ReDim Preserve Game(NewIndex)
RefreshGame NewIndex
GameIndex = NewIndex

End Sub

Private Sub cmdNext_Click()

GameIndex = GameIndex + 1


If GameIndex >= UBound(Game) Then GameIndex = UBound(Game)

RefreshGame GameIndex
End Sub

Private Sub cmdUpdate_Click()

txtGameEXE.Text = Trim(txtGameEXE.Text$)

    If Len(Game(GameIndex).GameEXE$) > 0 Then
        If Not LCase(Game(GameIndex).GameEXE$) = LCase(txtGameEXE.Text$) Then
            If MsgBox("Do you want to update the Game entry " & Game(GameIndex).Name & "?", vbYesNo, "Update") = vbYes Then
                UpdateGame GameIndex
            End If
        Else
            UpdateGame GameIndex
        End If
    Else
        UpdateGame GameIndex
    End If
    
    If chkGameUID.Value = vbChecked Then CalcGameUID GameIndex, True
End Sub

Private Sub Form_Load()
    If UBound(Game) = 0 Then
        cmdNew_Click
    Else
        RefreshGame 1
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Dim FullFileData        As String
Dim NoFilePath          As String
Dim NoFileExtension     As String
Dim GameEXE             As String
'Dim DataSize            As Long

For Each DroppedFile In Data.Files

    FullFileData = DroppedFile
    NoFilePath = Mid(FullFileData, 1, InStrRev(FullFileData, "\"))
    
    GameEXE = Mid(FullFileData, InStrRev(FullFileData, "\") + 1, Len(FullFileData) - InStrRev(FullFileData, "\") + 1)
    NoFileExtension = Mid(GameEXE, InStrRev(GameEXE, "\") + 1, InStrRev(GameEXE, ".") - 1)
    'DataSize = FileLen(FullFileData) / 1024

txtEXEPath.Text = NoFilePath
txtGameEXE.Text = GameEXE
If Len(txtMonitorEXE.Text) = 0 Then txtMonitorEXE.Text = GameEXE
'txtCMDArgs.Text = ""
txtIconPath.Text = FullFileData
If Len(txtName.Text) = 0 Then txtName.Text = StrConv(NoFileExtension, vbProperCase)

'If Len(NoFileExtension) > 0 Then
'If the file has no extension.. or is a folder, we don't want it.
'End If

Next DroppedFile

    Data.Files.Clear

End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

On Error Resume Next

If Data.GetFormat(vbCFFiles) Then

    Effect = vbDropEffectCopy And Effect

Else

    Effect = vbDropEffectNone
 
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    UpdateIconList True
    UpdateMaxPlayersMenu
End Sub

Private Sub Image1_Click()
MsgBox "In order to monitor the currently running game, we need to supply the EXE(s) that we expect to be running while that game is being played." & vbNewLine & _
"For instance, if we wanted to run the Windows Calculator, we could put in the Monitor EXE box: 'calc.exe;calculator.exe'" & vbNewLine & _
"It is possible to search for multiple EXEs by separating them with a semicolon.", vbInformation, "Monitor EXE"
End Sub

Private Sub Image2_Click()
MsgBox "Drag and drop a game's main executable into the box in order to populate the fields above with the correct information." & vbNewLine & _
"It is recommended that you use the ""Convert To Local Paths"" button to ensure that the EXE can be found when used on a different computer.", vbInformation, "Drag and Drop"
End Sub

Private Sub Image3_Click()
MsgBox "Checking this box generates a new GUID." & vbNewLine & "GUID, short for Game Unique ID, is calculated by hashing the contents of the GameEXE together with the Game Name." & vbNewLine & _
"If your GUID for a game doesn't match another LanParty user's GUID, they won't be able to see what you're playing.", vbInformation, "Game Unique Identifer"
End Sub

Private Sub txtIconPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

For Each DroppedFile In Data.Files

txtIconPath.Text = DroppedFile

Next DroppedFile

Data.Files.Clear

End Sub

Private Sub txtIconPath_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next

If Data.GetFormat(vbCFFiles) Then

    Effect = vbDropEffectCopy And Effect

Else

    Effect = vbDropEffectNone
 
End If


End Sub

Private Sub txtInstallerPath_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next

For Each DroppedFile In Data.Files

txtInstallerPath.Text = DroppedFile

Next DroppedFile

Data.Files.Clear

chkInstallFirst.Value = vbChecked

End Sub
