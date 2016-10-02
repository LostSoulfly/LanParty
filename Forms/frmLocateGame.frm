VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLocateGame 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Locate Game"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   2400
      Top             =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   1680
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Timer tmrLocate 
      Interval        =   1000
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.ListBox lstResults 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   6855
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdDlg 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Caption         =   "Search Results (Select the correct one and click Accept below):"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "Search File:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Search Folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status.."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Game EXE Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmLocateGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is only the FindFirstFile code, and it's been modified
' to only look for what I want, naturally..
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const vbDot = 46
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const vbBackslash = "\"
Private Const ALL_FILES = "*.*"

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   nCount As Long
   nSearched As Long
   sFileNameExt As String
   sFileRoot As String
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long


Private Declare Function lstrlen Lib "kernel32" _
    Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Declare Function PathMatchSpec Lib "shlwapi" _
   Alias "PathMatchSpecW" _
  (ByVal pszFileParam As Long, _
   ByVal pszSpec As Long) As Long

Private fp As FILE_PARAMS  'holds search parameters

Private intTimer As Integer

Private blUpdate As Boolean
Private blCancel As Boolean

Public Sub LocateGame(Optional Path As String)

lstResults.Clear

Status "Starting.."
tmrUpdate.Enabled = True

   'Dim tstart As Single   'timer var for this routine only
   'Dim tend As Single     'timer var for this routine only
      If Len(Path$) = 0 Then Path = Environ("HOMEDRIVE")
      
   With fp
      .sFileRoot = QualifyPath(Path) 'start path
      .sFileNameExt = Trim$(txtFile.Text) 'Game(GameIndex).GameEXE           'file type(s) of interest
      .bRecurse = True       'True = recursive search
      .nCount = 0                          'results
      .nSearched = 0                       'results
   End With
  
   'tstart = GetTickCount()
   Call SearchForFiles(fp.sFileRoot)
   'tend = GetTickCount()
   tmrUpdate.Enabled = False
   
   Status "Searched " & Format$(fp.nSearched, "###,###,###,##0") & " files. Please select the correct file below."
   
    cmdSearch.Caption = "Start"
    blCancel = True
    txtFolder.Enabled = True
    txtFile.Enabled = True
    cmdFolder.Enabled = True
   
   'List1.Visible = True
   'Text3.Text = Format$(fp.nSearched, "###,###,###,##0")
   'Text4.Text = Format$(fp.nCount, "###,###,###,##0")
   'Text5.Text = FormatNumber((tend - tstart) / 1000, 2) & "  seconds"

End Sub

Private Sub cmdAccept_Click()
'check that fileexists
'and set it to the game array object

    txtLocation.Text = Trim$(txtLocation.Text)

    Game(Me.Tag).EXEPath = Mid(txtLocation.Text, 1, InStrRev(txtLocation.Text, "\"))
    'NoFileExtension = Mid(NoFilePath, InStrRev(NoFilePath, "\") + 1, InStrRev(NoFilePath, ".") - 1)
    Game(Me.Tag).GameEXE32Bit = Mid(txtLocation.Text, InStrRev(txtLocation.Text, "\") + 1, Len(txtLocation.Text) - InStrRev(txtLocation.Text, "\") + 1)
    blCancel = True
    
    Me.Visible = False

End Sub

Private Sub cmdCancel_Click()
blCancel = True
Me.Visible = False
End Sub

Private Sub cmdDlg_Click()
tmrLocate.Enabled = False

With comDialog
    .DefaultExt = GetMyBitGameEXE(Me.Tag)
    .InitDir = App.Path
    .DialogTitle = "Locate " & GetMyBitGameEXE(GameIndex)(Me.Tag) & ".."
    .Filter = GetMyBitGameEXE(Me.Tag) & "|" & GetMyBitGameEXE(Me.Tag) & "|" _
                            & "All Files" & "|" & "*.*"
    .ShowOpen
    If FileExists(.FileName) Then txtLocation.Text = .FileName
End With

'if game has been located, we can close this form after calling an update game function
'and storing the new game exe path into the game object.
End Sub

Private Sub cmdFolder_Click()
tmrLocate.Enabled = False

    txtFolder.Text = GetFolder(Me.hwnd, txtFolder.Text, "Select a folder to start searching..")
End Sub

Private Sub cmdSearch_Click()
tmrLocate.Enabled = False

Select Case cmdSearch.Caption

Case Is = "Cancel"
    cmdSearch.Caption = "Start"
    blCancel = True
    txtFolder.Enabled = True
    txtFile.Enabled = True
    cmdFolder.Enabled = True
    
Case Is = "Start"
    cmdSearch.Caption = "Cancel"
    blCancel = False
    txtFolder.Enabled = False
    txtFile.Enabled = False
    cmdFolder.Enabled = False
    LocateGame txtFolder.Text
End Select

End Sub

Public Sub SetSearchFile(strFile As String)
    txtFile.Text = strFile
End Sub

Private Sub Form_Click()
    tmrLocate.Enabled = False
End Sub

Private Sub Form_Load()
txtFolder.Text = Environ("HOMEDRIVE") & "\"
Status "Starting Auto Locate.. Click above to cancel."
End Sub

Private Function Status(Text As String)
    lblStatus.Caption = Text
End Function

Private Sub SearchForFiles(sRoot As String)

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
  
   hFile = FindFirstFile(sRoot & ALL_FILES, WFD)
  
   'Status "Searching.. " & Chr(34) & fp.sFileNameExt & Chr(34)
  
   If hFile <> INVALID_HANDLE_VALUE Then
    
    If blUpdate Then
        
        If Len(sRoot) > 60 Then
            If Len(sRoot) < 100 Then
                Status "Searching.. " & Left(sRoot, 10) & ".." & Right(sRoot, Len(sRoot) / 2)
            ElseIf Len(sRoot) < 150 Then
                Status "Searching.. " & Left(sRoot, 5) & ".." & Right(sRoot, Len(sRoot) / 3)
            Else
                Status "Searching.. " & Left(sRoot, 5) & ".." & Right(sRoot, Len(sRoot) / 5)
            End If
        Else
            Status "Searching.. " & sRoot
        End If
    blUpdate = False
    End If
      Do
      
      If blCancel Then FindClose (hFile): Exit Sub
      DoEvents
        'if a folder, and recurse specified, call
        'method again
         If (WFD.dwFileAttributes And vbDirectory) Then
            If Asc(WFD.cFileName) <> vbDot Then

             If fp.bRecurse Then
                  SearchForFiles sRoot & TrimNull(WFD.cFileName) & vbBackslash
               End If
            End If
            
         Else
         
           'must be a file..
            If MatchSpec(WFD.cFileName, fp.sFileNameExt) Then
               fp.nCount = fp.nCount + 1
               lstResults.AddItem sRoot & TrimNull(WFD.cFileName)
               If lstResults.ListCount = 1 Then lstResults.ListIndex = 0
            End If  'If MatchSpec
      
         End If 'If WFD.dwFileAttributes
      
         fp.nSearched = fp.nSearched + 1
      
      Loop While FindNextFile(hFile, WFD)
   
   End If 'If hFile
  
   Call FindClose(hFile)

End Sub

Private Function QualifyPath(sPath As String) As String

   If Right$(sPath, 1) <> vbBackslash Then
      QualifyPath = sPath & vbBackslash
   Else
      QualifyPath = sPath
   End If
      
End Function

Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlen(StrPtr(startstr)))
   
End Function

Private Function MatchSpec(sFile As String, sSpec As String) As Boolean

   MatchSpec = PathMatchSpec(StrPtr(sFile), StrPtr(sSpec))
   
End Function

Private Sub Form_Unload(Cancel As Integer)
    UpdateIconList True
End Sub

Private Sub Label1_Click()
tmrLocate.Enabled = False
End Sub

Private Sub Label2_Click()
tmrLocate.Enabled = False
End Sub

Private Sub Label3_Click()
tmrLocate.Enabled = False
End Sub

Private Sub lblStatus_Click()
tmrLocate.Enabled = False
End Sub

Private Sub lstResults_Click()
txtLocation.Text = lstResults.Text

End Sub

Private Sub tmrLocate_Timer()

Status "Starting Auto Locate in " & (5 - intTimer) & " Seconds.. Click above to cancel."

If intTimer = 5 Then

    cmdSearch_Click
    
tmrLocate.Enabled = False
End If

intTimer = intTimer + 1
End Sub

Private Sub tmrUpdate_Timer()
blUpdate = True
End Sub

Private Sub txtFolder_Click()
tmrLocate.Enabled = False
End Sub

Private Sub txtLocation_Change()
If FileExists(txtLocation.Text) Then
    cmdAccept.Enabled = True
    txtLocation.BackColor = vbGreen
Else
    cmdAccept.Enabled = False
    txtLocation.BackColor = vbRed
End If
End Sub

Private Sub txtLocation_Click()
    tmrLocate.Enabled = False
End Sub
