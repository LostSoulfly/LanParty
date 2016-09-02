Attribute VB_Name = "modGameScan"
Option Explicit


Public blContinueStartup As Boolean
Public blBootComplete As Boolean

'check game if game exists with GameIndexByUID, if it's 1 or more, it exists

Public Sub ParseGameFiles(strGameFiles() As String)
Dim i As Integer
Dim TempGame() As GameData
ReDim TempGame(0)
    
    For i = 0 To UBound(strGameFiles)
        
        Debug.Print strGameFiles(i)
        
        GetGamesFromText LoadFile(strGameFiles(i)), TempGame    'Remember, the passed UDT is cleared!
        AddGamesToFrom Game, TempGame                           'Compare and add (To, From)
        
    Next i


End Sub
