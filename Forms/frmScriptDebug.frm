VERSION 5.00
Begin VB.Form frmScriptDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   3315
   ClientLeft      =   7800
   ClientTop       =   11910
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7650
   Begin VB.Timer tmrSnap 
      Interval        =   10
      Left            =   6120
      Top             =   480
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   7095
   End
   Begin VB.TextBox txtOutput 
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuLvl 
         Caption         =   "Level"
         Begin VB.Menu mnuDebugLevel 
            Caption         =   "Level 1 (Default)"
            Index           =   0
         End
         Begin VB.Menu mnuDebugLevel 
            Caption         =   "Level 2"
            Index           =   1
         End
         Begin VB.Menu mnuDebugLevel 
            Caption         =   "Level 3"
            Index           =   2
         End
         Begin VB.Menu mnuDebugLevel 
            Caption         =   "Level 4"
            Index           =   3
         End
      End
      Begin VB.Menu mnuErrors 
         Caption         =   "Show Script Errors"
      End
   End
End
Attribute VB_Name = "frmScriptDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public theScript As clsScript
Private txtOutLen As Long

Private Sub Form_Load()
    'LockWindowUpdate Me.hWnd
End Sub

Private Sub Form_Resize()
    txtOutput.Height = Me.Height - 1100
    txtOutput.Width = Me.Width - 350
    txtInput.Top = txtOutput.Height + txtOutput.Top + 100
    txtInput.Width = txtOutput.Width
End Sub

Public Sub PrintDebug(Msg As String)
    
    txtOutput.Text = txtOutput.Text & "[" & Time & "] " & Msg & vbNewLine
    txtOutLen = txtOutLen + Len(Msg) + 15
    DoEvents
    
    If txtOutLen > 62000 Then _
        If MsgBox("The Debug Console is nearing its character limit. Would you like to clear it now?" & vbNewLine & _
        "(Failure to do so may result in a crash).", vbYesNo, "Debug") = vbYes Then txtOutput.Text = "": txtOutLen = 0
        
    
End Sub

Private Sub tmrSnap_Timer()
On Error Resume Next
If Me.Visible = False Then Exit Sub
    Me.Top = frmMain.Top + frmMain.Height
    Me.Width = frmMain.Width
    Me.Left = frmMain.Left
End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strTemp As String
Dim lngTemp As Long

    If KeyCode = 13 Then
        lngTemp = InStr(1, txtInput.Text, " ")
        strTemp = Trim$(Left(txtInput.Text, IIf(lngTemp > 0, lngTemp, Len(txtInput.Text))))
        
        If InStr(1, strTemp, "(") > 0 Then
            theScript.DoFunction txtInput.Text
            txtInput.Text = ""
            Exit Sub
        End If
        
        
        Select Case LCase$(strTemp)
        
            Case "help", "info"
                PrintDebug "--- LSL Scripting Debug Console ---"
                PrintDebug "Commands available in this window:"
                PrintDebug "Restart, Run" & vbTab & vbTab & " - Restart/Run the current script from the beginning."
                PrintDebug "Reset, Stop" & vbTab & vbTab & " - Reset all script variables and start fresh."
                PrintDebug "Close, Exit" & vbTab & vbTab & " - Closes Debug Window and sets DebugLevel to 0."
                PrintDebug "Variables, ShowVars" & vbTab & " - Enumerate and print all current variables for the loaded script."
                PrintDebug "FileVars, ShowFileVars" & vbTab & " - Enumerate and print all current file variables for the loaded script."
                PrintDebug "clear, cls" & vbTab & vbTab & " - Clear the debug console."
                PrintDebug ""
                PrintDebug "Also, any single-line functions can also be called from this window."
                PrintDebug "You may also dim/set variables."
                
            Case Is = "restart", "run"
                Call theScript.ScriptExecute
                
            Case Is = "reset", "stop"
                Call theScript.ExitScript
                            
            Case Is = "close", "exit"
                theScript.DoFunction ("debug(""hide"")")
                theScript.DoFunction ("debug(""0"")")
                Unload Me
                
            Case Is = "variables", "showvars"
                EnumVars
                
            Case Is = "filevars", "showfilevars"
                EnumFileVars
                
            Case Is = "clear", "cls"
                theScript.DoFunction ("debug(""clear"")")
                
            Case Else
                theScript.LineExecute txtInput.Text
                
        End Select
        txtInput.Text = ""
    End If
    
End Sub

Private Sub txtOutput_Change()
txtOutput.SelStart = 65535
End Sub

Public Sub Clear()
    txtOutput.Text = ""
End Sub

Private Sub EnumVars()
On Error Resume Next
Dim strNames() As String
Dim strVars() As String
Dim i As Long

PrintDebug "--- Variable Names and contents ---"

theScript.GetAllVars strNames, strVars

    For i = 1 To UBound(strNames)
    
        txtOutput.Text = txtOutput.Text & "Name: " & strNames(i) & ": [" & strVars(i) & "]" & vbNewLine
    
    Next i
    
PrintDebug "--- End Variables and contents ---"
End Sub

Private Sub EnumFileVars()
On Error Resume Next
Dim strNames() As String
Dim strVars() As String
Dim i As Long

PrintDebug "--- FileVar Names and contents ---"

theScript.GetAllFileVars strNames, strVars

    For i = 1 To UBound(strNames)
    
        txtOutput.Text = txtOutput.Text & "Name: " & strNames(i) & ": [" & strVars(i) & "]" & vbNewLine
    
    Next i
    
PrintDebug "--- End FileVars and contents ---"
End Sub
