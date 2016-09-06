Attribute VB_Name = "modDrawGames"
Global Const DEFAULT_LEFT = 120
Global Const DEFAULT_TOP = DEFAULT_LEFT * 2
Global Const DEFAULT_SPACING = DEFAULT_LEFT
Public lngLastIconsPerRow As Long
Public lngScrollAmt As Long
Public IconHeight As Long
Public IconWidth As Long
Public NumPlayers As Integer

Private Type PlayersUDT
    NumPlayers As Integer
    GameCount As Integer
End Type

Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Sub InitIcons()
On Error GoTo oops

Dim i As Integer

frmMain.lblIcon(0).Caption = ""

'this is sure to error often.
For i = 1 To frmMain.imgIcon.Count
    Unload frmMain.imgIcon(i)
    Unload frmMain.lblIcon(i)
Next i

Exit Sub

oops:

If err.Number = 340 Then err.Clear: Exit Sub
If err.Number = 360 Then Resume Next
If err.Number = 365 Then Resume Next
'Resume Next
AddDebug err.Number & ": " & err.Description

End Sub

Public Sub UpdateIconList(Optional ForceUpdate As Boolean = False)
On Error GoTo Escape
Dim CurrentIconsPerRow As Integer
Dim LoopCurrentRow As Integer
Dim CurrentNumberOfRows As Integer
Dim NextLeft As Long
Dim NextTop As Long
Dim IconTotalWidth As Long
Dim IconTotalHeight As Long
Dim DisplayIcons As Integer
Dim i As Integer
Dim ii As Integer

If ((Not Game) = -1) Then Exit Sub 'prevent an early error, though it would be caught with err 9 at the bottom

If (UBound(Game) = 0) And (LenB(Game(0).Name$) = 0) Then Exit Sub

If NumPlayers > -1 Then
    'loop through the list of games and find each game with NumPlayers as .maxplayers
    'set it to displayicons
    
    For i = 1 To UBound(Game)
        If Game(i).MaxPlayers = NumPlayers Then DisplayIcons = DisplayIcons + 1
    Next i
    
Else
    DisplayIcons = UBound(Game)
End If


'If frmMain.mnuList.Checked = True Then
'    IconWidth = 600
'    IconHeight = 600
    'this is a special case where the label is wider, not the actual icon.
'End If

'calculate stuff we will use later
IconTotalWidth = DEFAULT_SPACING + IconWidth
IconTotalHeight = DEFAULT_SPACING + IconHeight + frmMain.lblIcon(0).Height

'Calculate the number of icons of a certain width per ROW
CurrentIconsPerRow = Fix((frmMain.Width - (frmMain.vsScroll.Width * 2)) / IconTotalWidth)

'se the default positions
NextLeft = DEFAULT_LEFT
NextTop = DEFAULT_TOP

If ForceUpdate = False Then
'don't redraw unless we really need to.
    If CurrentIconsPerRow = lngLastIconsPerRow Then Exit Sub 'AddDebug "No need to update rows.": Exit Sub
    'If lngCurrentNumIcons > 0 Then If lngCurrentNumIcons = UBound(Game) Then Exit Sub
End If

'set the current number of icons per row for next redraw
lngLastIconsPerRow = CurrentIconsPerRow

If CurrentIconsPerRow = 0 Then CurrentIconsPerRow = 1 ': AddDebug "NumIcons 0"

Dim intRnd As Integer
intRnd = (DisplayIcons / CurrentIconsPerRow)

'we need to round up.. otherwise some of the icons will be off the screen on the last row.
If (DisplayIcons / CurrentIconsPerRow) > intRnd Then intRnd = intRnd + 1

'set our number of rows to the rounded up count of rows games / games per row = number of rows rounded up.
CurrentNumberOfRows = intRnd

'Reinitialize the icons and labels..
InitIcons


'Start a loop for each game, and spawn a new icon/lbl for each one and calculate its location.

    For i = 1 To UBound(Game)
    DoEvents
        If NumPlayers = -1 Or (NumPlayers > -1 And Game(i).MaxPlayers = NumPlayers) Then
            If LoopCurrentRow = CurrentIconsPerRow Then
                NextTop = NextTop + IconTotalHeight
                NextLeft = DEFAULT_LEFT
                LoopCurrentRow = 0
            End If
            
            If i >= frmMain.imgIcon.Count Then
                ii = frmMain.imgIcon.Count
                Load frmMain.imgIcon(ii)
                Load frmMain.lblIcon(ii)
            Else
                ii = i
            End If
            'determine this icon's placement.
            
            With frmMain.imgIcon(ii)
                
                If Settings.ShowIcons Then
                'AddDebug "GetGamePicture: " & GetGameName(i)
                    If FileExists(FixFilePath(Game(i).IconPath)) Then
                        .Picture = GetGamePicture(Game(i).IconPath)
                    ElseIf FileExists(GetGameExePath(i)) Then
                        .Picture = GetGamePicture(GetGameExePath(i))
                    ElseIf FileExists(FixFilePath(Game(i).InstallerPath)) Then
                        .Picture = GetGamePicture(Game(i).InstallerPath)
                    End If
                End If
                'If .Picture <> 0 Then frmMain.imgList.ListImages.add , Game(i).GameUID, .Picture
                    '.Picture = GetGamePicture(IIf(Len(Game(i).IconPath) = 0, GetGameExePath(i), Game(i).IconPath)) 'if the icon path is blank, use the EXE instead..
                .Stretch = True
                .BorderStyle = 0
                .Left = NextLeft
                .Top = NextTop
                .Height = IconHeight
                .Width = IconWidth
                .Visible = True
                .Tag = i
            End With
                
            With frmMain.lblIcon(ii)
                .Left = NextLeft
                .Top = NextTop + frmMain.imgIcon(ii).Height
                .Caption = Game(i).Name
                .Width = IconWidth
                .ForeColor = Settings.IconTextColor
                .FontSize = Settings.IconTextSize
                .Visible = True
            End With
                
            NextLeft = NextLeft + IconTotalWidth
            
            'If i = 1 Then NextLeft = IconTotalWidth
            lngCurrentNumIcons = lngCurrentNumIcons + 1
            LoopCurrentRow = LoopCurrentRow + 1
        End If
    Next i
    
    Dim iconTopTotal As Long
    
    iconTopTotal = frmMain.imgIcon(DisplayIcons).Top + IconTotalHeight
    
    frmMain.picContainer.Height = iconTopTotal

Exit Sub

Escape:

If err.Number = 9 Then err.Clear: Exit Sub
If err.Number = 360 Then Resume Next
If err.Number = 365 Then Resume Next
'Resume Next
AddDebug err.Number & ": " & err.Description
'err.Clear

End Sub

Public Function UpdateMaxPlayersMenu()
    Dim i As Integer
    Dim ii As Integer
    Dim NotEmpty As Boolean
    Dim arrGames() As PlayersUDT
    ReDim arrGames(UBound(Game))

    For i = 0 To UBound(Game)
    DoEvents
        With Game(i)
            If LenB(.GameEXE) > 0 Then
                If .MaxPlayers > 0 Then 'the maxPlayers has been set for this game.
                    NotEmpty = True
                    For ii = 0 To UBound(arrGames)
                    DoEvents
                        If arrGames(ii).NumPlayers = 0 Then arrGames(ii).NumPlayers = .MaxPlayers
                        If arrGames(ii).NumPlayers = .MaxPlayers Then _
                            arrGames(ii).GameCount = arrGames(ii).GameCount + 1: Exit For
                    Next ii
                End If
            End If
        End With
    Next i
    
    If NotEmpty = False Then frmMain.mnuPlayers.Visible = False: Exit Function
    
    

    'then load/unload the menus here and populate their captions.
    'need to set their tag data. Set it to their actual player number?
    
    'unload all but last existing menus
    For i = (frmMain.mnuNumPlayers.Count - 1) To 1 Step -1
        Unload frmMain.mnuNumPlayers(i)
    Next i
    
    frmMain.mnuNumPlayers(0).Tag = -1
    frmMain.mnuNumPlayers(0).Caption = "Clear Filter"

    Dim UnkPlayers As Integer
    For i = 0 To UBound(arrGames)
    DoEvents
        If arrGames(i).GameCount > 0 Then
            ii = frmMain.mnuNumPlayers.Count
                Load frmMain.mnuNumPlayers(ii)
                frmMain.mnuNumPlayers(ii).Tag = arrGames(i).NumPlayers
                frmMain.mnuNumPlayers(ii).Caption = arrGames(i).NumPlayers & " Players (" & arrGames(i).GameCount & " Games)"
        Else
            UnkPlayers = UnkPlayers + 1
        End If
    Next
    
    If UnkPlayers > 0 Then
        ii = frmMain.mnuNumPlayers.Count
        Load frmMain.mnuNumPlayers(ii)
        frmMain.mnuNumPlayers(ii).Tag = 0
        frmMain.mnuNumPlayers(ii).Caption = "Unknown Players (" & (UnkPlayers - 1) & " Games)"
    End If
    
    frmMain.mnuPlayers.Visible = True
    
End Function

