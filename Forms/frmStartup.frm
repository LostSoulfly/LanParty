VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Startup"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   120
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   4320
      Y1              =   650
      Y2              =   650
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading, please wait.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label lblSkip 
      BackStyle       =   0  'Transparent
      Caption         =   "Skip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmStartup"
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
Private strGameFiles() As String

Public Sub LocateGameFiles(GameFileName As String)
Dim Path As String

tmrUpdate.Enabled = True
Path = App.Path & "\"

   'Dim tstart As Single   'timer var for this routine only
   'Dim tend As Single     'timer var for this routine only
      
   With fp
      .sFileRoot = QualifyPath(Path) 'start path
      .sFileNameExt = GameFileName 'Game(GameIndex).GameEXE           'file type(s) of interest
      .bRecurse = True       'True = recursive search
      .nCount = 0                          'results
      .nSearched = 0                       'results
   End With
  
   'tstart = GetTickCount()
   Call SearchForFiles(fp.sFileRoot)
   'tend = GetTickCount()
   tmrUpdate.Enabled = False
   
   SetStatus "Searched " & Format$(fp.nSearched, "###,###,###,##0") & " files.."
   Pause 100

    blCancel = True

End Sub

Private Sub cmdCancel_Click()
blCancel = True
Me.Visible = False
End Sub

Private Sub Form_Click()
    'show a loading screen of sorts, with progress indications
    'scan through all subdirs and look for a specific file in each folder
    'read that file and determine if the file is a supported format
    'determine if the game in the file has already been added to the list
    'if not, add it to the list.
    'keep looping through all subdirs until complete.
    'save the list and continue on with the startup process.
    'decide whether to keep the loaded games or re-load the list.
    'decide whether to
End Sub

Private Sub Form_Load()
Me.BackColor = Settings.ChatBGColor
lblLoading.ForeColor = Settings.ChatTextColor
lblStatus.ForeColor = Settings.ChatTextColor
lblSkip.ForeColor = Settings.ChatTextColor
ReDim strGameFiles(0)
Me.Visible = True
Me.Show
Pause 100
If Settings.ScanAtStartup Then
    SetCaption "Scanning for games, please wait.."
    SetStatus "Searching for games.."
    LocateGameFiles "GameData.lan"
    lblSkip.Visible = False
    InitializeGameArray True
    'SetStatus "Checking for new games.."
    If Not ((Not strGameFiles) = -1) Then ParseGameFiles strGameFiles
End If
Pause 500
'Allow startup to continue
blContinueStartup = True
End Sub

Public Sub SetStatus(Text As String)
    If LenB(Text) = 0 Then Exit Sub
    lblStatus.Caption = Text
    DoEvents
End Sub

Public Sub SetCaption(Text As String)
    If LenB(Text) = 0 Then Exit Sub
    lblLoading.Caption = Text
    DoEvents
End Sub

Private Sub SearchForFiles(sRoot As String)

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sRootTemp As String
  
   hFile = FindFirstFile(sRoot & ALL_FILES, WFD)
  
   'Status "Searching.. " & Chr(34) & fp.sFileNameExt & Chr(34)
  
   If hFile <> INVALID_HANDLE_VALUE Then
    
    If blUpdate Then
        sRootTemp = Right(sRoot, Len(sRoot) - Len(App.Path & "\"))
        If Len(sRootTemp) > 60 Then
            If Len(sRootTemp) < 100 Then
                SetStatus Left(sRootTemp, InStr(1, sRootTemp, "\")) & ".." & Right(sRootTemp, Len(sRootTemp) / 2)
            ElseIf Len(sRootTemp) < 150 Then
                SetStatus Left(sRootTemp, InStr(1, sRootTemp, "\")) & ".." & Right(sRootTemp, Len(sRootTemp) / 3)
            Else
                SetStatus Left(sRootTemp, InStr(1, sRootTemp, "\")) & ".." & Right(sRootTemp, Len(sRootTemp) / 5)
            End If
        Else
            SetStatus sRootTemp
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
               AddGameDataFile sRoot & TrimNull(WFD.cFileName)
            End If  'If MatchSpec
      
         End If 'If WFD.dwFileAttributes
      
         fp.nSearched = fp.nSearched + 1
      
      Loop While FindNextFile(hFile, WFD)
   
   End If 'If hFile
  
   Call FindClose(hFile)

End Sub

Private Function AddGameDataFile(sFile As String)

    If LenB(sFile) < 2 Then Exit Function
    If LenB(strGameFiles(0)) > 0 Then ReDim Preserve strGameFiles(UBound(strGameFiles) + 1)
    strGameFiles(UBound(strGameFiles)) = sFile

End Function

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
    If blContinueStartup = False Then End
    If blBootComplete = False Then End
End Sub

Private Sub lblSkip_Click()
blCancel = True
blContinueStartup = True

End Sub

Private Sub tmrUpdate_Timer()
blUpdate = True
End Sub

