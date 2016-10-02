VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Possible to create an input array and reference the array number in commands
'if more than one input is necessary..

Dim tempInput As String, tempListInput As String, tempVar As Long, tempBool as Boolean
Dim screenWidth As String
Dim screenHeight As String
Dim DoNextLine As Boolean

Public Function ParseScript(Script() As String) As Boolean
Dim i As Integer
Dim Line As String

For i = 0 To UBound(Script)

    Line = Trim$(Script(i))
    If Not Left$(Line, 2) = "//" Then 'skip comments
        ParseCommand (SplitString(Line))
    End If

Next i


End Function

Public Function ParseCommand(Command() As String) As String

'vars to use in code:
'%name% replaces with chat name
'%gamepath% replaces with game's path
'%resx% replaces with screen width
'%resy% replaces with screen height
'%ip% replaces with the current network IP of this machine
'%var% replaces withe the current tempVar
'%tinput% replaces with current tempInput
'%tlist% replaces with the current tempListInput

'to be determined, the best way to verify this..
If UBound(Command) <= 1 Then ParseCommand = "Array length check failed": Exit Function
If LenB(Command(1)) = 0 Then ParseCommand = "Command check failed": Exit Function

Select Case LCase$(Command(1))
    
    Case "clearallvars"
        tempInput = ""
        tempListInput = ""
        tempVar = ""
        tempBool = false
		
    Case "clearinput"
        tempInput = ""
        
    Case "clearlistinput"
        tempListInput = ""
        
    Case "clearvar"
        tempVar = 0
    
	Case "clearyesno"
		tempBool = false
		
    Case "setvar"
        If IsNumeric(Command(2)) Then
            tempVar = Command(2)
        Else
            ParseCommand = "Error in setting (" & Command(1) & Command(2) & "). Not a number. "
            Exit Function
        End If
        
    Case "incvar"
        tempVar = tempVar + 1
        
    Case "decvar"
        tempVar = tempVar - 1
        
	Case "ifvar="
		'if the tempVar is this
		DoNextLine = true
	
	case "ifvar>"
		'if tempVar is > this
		DoNextLine = true
		
	case "ifvar<"
		'if tempVar < this
		DoNextLine = True

	Case "getyesno"
		'msgbox with title, text, and yes/no buttons
		'set tempBool
		
	case "ifyes"
		DoNextLine = True

	case "ifno"
		DoNextLine = True
			
    Case "getinput"
        'show input window, specify caption, text, and default
        'set tempInput
    
    Case "getinputwithlist"
        'show an input window with a dropdown containing options specified in the command, and sets the temp
        'set TempInputList
        
    Case "getinputres"
        'show an input window with widthxheight selection list populated
        
    Case "getres"
        'get screen's res. Might have the class get this when initialized..
    
    Case "createfile"
        'create a file
    
    Case "writetofile"
        'append to end of file
    
    Case "copyfile"
    
    Case "renamefile"
    
    Case "renamefolder"
    
    Case "createfolder"
    
    Case "deletefile"
    
    Case "deletefolder"
    
    Case "iffileexists"
        DoNextLine = True
    
    Case "ifstringinfile"
        DoNextLine = True
        
    Case "if64bit"
        DoNextLine = True
        
    Case "if32bit"
        DoNextLine = True
        
    Case Else
        ParseCommand = "Unknown command: " & Command(1)

End Select

End Function

Public Function SplitString(StringToSplit As String) As String()
    Dim sTmp() As String, sTemp As String
    Dim iCounter As Integer
    sTmp = Split(" " & StringToSplit, Chr(34))
    For iCounter = 0 To UBound(sTmp)
        If Trim$(sTmp(iCounter)) = "" Then sTmp(iCounter) = ""
        If Left$(sTmp(iCounter), 1) = " " Then
            sTemp = sTemp & Replace$(sTmp(iCounter), " ", Chr(34)) & Chr(34)
        Else
            sTemp = sTemp & sTmp(iCounter) & Chr(34)
        End If
    Next iCounter
    sTemp = Replace(sTemp, Chr(34) & Chr(34), Chr(34))
    SplitString = Split(sTemp, Chr(34))
End Function

Private Function IsPathAllowed(Path As String) As Boolean

    'this should be made into its own functin and called every time
    'we need to get an EXEPath.. It's used in a few places.
    If Left(Path, 1) = "\" Then
        Path = FullPathFromLocal(Path)
    End If

    If InStr(1, LCase$(Path), LCase$(App.Path)) > 0 Then IsPathAllowed = True

End Function

