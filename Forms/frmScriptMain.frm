VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScriptMain 
   Caption         =   "Script Editor"
   ClientHeight    =   5400
   ClientLeft      =   7800
   ClientTop       =   3330
   ClientWidth     =   8250
   Icon            =   "frmScriptMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8250
   Begin MSComDlg.CommonDialog cd 
      Left            =   5310
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   15
      X2              =   6615
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   6615
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open file.."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save to file.."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveGameScript 
         Caption         =   "Save GameScript"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMake 
         Caption         =   "Make EXE..."
         Shortcut        =   {F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Begin VB.Menu mnuEnableDebug 
         Caption         =   "Enable Debug"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSnapDebug 
         Caption         =   "Snap Debug Window To Bottom"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowErrors 
         Caption         =   "Show Script Errors"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDebugLevel 
         Caption         =   "Debug Level"
         Begin VB.Menu mnuLevel 
            Caption         =   "Level 0 (Default)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuLevel 
            Caption         =   "Level 1"
            Index           =   1
         End
         Begin VB.Menu mnuLevel 
            Caption         =   "Level 2"
            Index           =   2
         End
         Begin VB.Menu mnuLevel 
            Caption         =   "Level 3 (All Output)"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
      Begin VB.Menu mnuRunRun 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuStep 
         Caption         =   "Step"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset Script/Variables"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpCommands 
         Caption         =   "Quick Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmScriptMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' JEL Script v1.0 : frmScriptMain.frm
' modified by LostSoulFly


Public WithEvents myScript As clsScript
Attribute myScript.VB_VarHelpID = -1

Option Explicit
Private blDebug As Boolean
Private blShowErrors As Boolean
Private intDebug As Integer
Public blCancel As Boolean
Public EditScript As Boolean

Private Sub Form_Load()
    Set myScript = New clsScript
    blDebug = True
    SetDebug
End Sub

Public Sub EditGameScript(GameScript As String)
    mnuFileSep2.Visible = True
    mnuSaveGameScript.Visible = True
    txtScript.Text = GameScript
    EditScript = True
End Sub

Private Sub SetDebug()
    myScript.blDebug = blDebug
    myScript.blShowErrors = blShowErrors
    myScript.DebugLevel = intDebug
    Set myScript.myParent = Me
    myScript.ShowDebug
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtScript.Width = Me.Width - 300
txtScript.Height = Me.Height - 950
End Sub

Private Sub Form_Unload(Cancel As Integer)
myScript.myDebug.Visible = False
Unload myScript.myDebug

If EditScript Then
    If MsgBox("Do you want to save the changes to this GameScript?", vbYesNo, "Save changes?") = vbNo Then blCancel = True
    Me.Visible = False
    Cancel = 1
    Exit Sub
End If


End Sub

Private Sub mnuEnableDebug_Click()
    blDebug = Not blDebug
    myScript.blDebug = blDebug
    mnuEnableDebug.Checked = blDebug
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMake_Click()
    On Error GoTo ErrorCatch:
    
    If Not FileExists(App.Path & "\LSLexe.exe") Then
        MsgBox "Program could not find the second LSL executable, aborting.", vbCritical, "Error"
        Exit Sub
    End If
    
    cd.FileName = ""
    cd.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
    cd.ShowSave
    If cd.FileName <> "" Then
        If FileExists(cd.FileName) Then
            If MsgBox("Overwrite existing file?", vbQuestion + vbYesNo, "LSL") = vbNo Then
                Exit Sub
            Else
                Kill cd.FileName
            End If
        End If
        
        Dim nFile As Integer
        nFile = FreeFile
        FileCopy App.Path & "\LSLexe.exe", cd.FileName
        Open cd.FileName For Output As #nFile
        Print #nFile, "|*LSL*|" & txtScript.Text
        Close #nFile
        MsgBox "File Compiled!", vbInformation, "LSL"
        
        Dim sTemp As String, sTemp2 As String
        
        Open cd.FileName For Output As #1
        Open App.Path & "\LSLexe.exe" For Binary As #2
        
        ' Copy data from LSLexe into new exe
        While Not EOF(2)
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
            If Len(sTemp) > 2000 Then
                sTemp = ""
            End If
        Wend
        
        ' Append the script
        Print #1, "|*LSL*|" & txtScript.Text
        
        Close #2
        Close #1
        
        
    End If
    
    Exit Sub
ErrorCatch:
    MsgBox "Error has occured: " & err.Description, vbCritical, "Error"
    Resume Next
End Sub

Private Sub mnuHelpCommands_Click()
frmHelp.Visible = True
End Sub

Private Sub mnuLevel_Click(Index As Integer)

    Dim i As Integer
    intDebug = Index
    
    mnuLevel(Index).Checked = True
    
    For i = 0 To mnuLevel.Count - 1
        If Not Index = i Then mnuLevel(i).Checked = False
    Next i
    
    myScript.DebugLevel = intDebug

End Sub

Private Sub mnuOpen_Click()
    cd.FileName = ""
    cd.Filter = "LSL Source Files (*.LSL)|*.lsl|All Files (*.*)|*.*"
    cd.ShowOpen
    If cd.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open cd.FileName For Input As nFile
        txtScript.Text = Input(LOF(nFile), nFile)
        Close nFile
    End If
End Sub

Private Sub mnuReset_Click()
    'Set myScript = New clsScript
    myScript.ShowDebug
    myScript.ExitScript
End Sub

Private Sub mnuRunRun_Click()
    RunScript
End Sub

Public Sub RunScript()
    If myScript.Script <> "" Then
        If MsgBox("Would you like to reset the variables and arrays?", vbYesNo, "Clear State") = vbYes Then
            Call myScript.ExitScript
        End If
    End If
    myScript.Script = txtScript.Text
    'Set myScript.theForm = Me
    myScript.ScriptExecute
End Sub

Private Sub mnuSave_Click()
    cd.FileName = ""
    cd.Filter = "LSL Source Files (*.LSL)|*.lsl|All Files (*.*)|*.*"
    cd.ShowSave
    If cd.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open cd.FileName For Output As nFile
        Print #nFile, txtScript.Text
        Close nFile
    End If
End Sub

Private Sub mnuSaveGameScript_Click()
    blCancel = False
    Me.Visible = False
End Sub

Private Sub mnuShowErrors_Click()
    blShowErrors = Not blShowErrors
    myScript.blShowErrors = blShowErrors
    mnuShowErrors.Checked = blShowErrors
End Sub

Private Sub mnuSnapDebug_Click()
    mnuSnapDebug.Checked = Not mnuSnapDebug.Checked
    myScript.myDebug.tmrSnap.Enabled = mnuSnapDebug.Checked
End Sub

Private Sub mnuStep_Click()
    StepScript
End Sub

Public Sub StepScript()
    If myScript.nCurrentline > myScript.nNumCurrentLines Then
        If MsgBox("You've reached the end of the script. Reset?", vbYesNo, "Reset Script?") = vbYes Then
            'Set myScript = New clsScript
            myScript.ExitScript
        End If
    End If
    
    myScript.Script = txtScript.Text
    myScript.ScriptExecute True
End Sub

Private Sub myScript_CommandOut(Command As String, ArgList() As String)
    Debug.Print "Command: " & Command
End Sub

Private Sub myScript_LineChange(nLine As Long, sContents As String)
    Me.Caption = "Script Editor - Parsing Line: " & (nLine + 1)
    DoEvents
End Sub

Private Sub myScript_ScriptError(Msg As String, Line As Long)
    MsgBox Msg, vbCritical, "Line: " & CStr(Line)
End Sub

Private Sub txtScript_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Then
        KeyAscii = Asc(" ")
    End If
End Sub


