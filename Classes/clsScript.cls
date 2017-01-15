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
' Original code: Justin's Elite Language, AKA JEL.
' Found this on PlanetSourceCode.
'
' Edited/Modified heavily by LostSoulfly
' It's been a challenge editing and working with this code.
' It's horribly optimized and wasn't designed for speed at all.
' I've attempted to add recursive function processing as well as
' simple arrays and more. Large portions from the original code have
' been changed. I've been adding functions and (hopefully) improving the usability of the language.
' I split up the ScriptExecute into two functions so that each line can be input and processed
' outside of the ScriptExecute function, enabling the Debug console.
' However, I didn't want to rewrite it, so I just ripped the variables out of it.

Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'// Insert and delete declarations
Private Const GENERIC_READ_WRITE As Long = &HC0000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_BEGIN As Long = 0
Private Const NO_ERROR As Long = 0
 
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
 
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
     ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
     ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
     ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
     ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, _
     lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" ( _
     ByVal hFile As Long, ByVal lDistanceToMove As Long, _
     lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
 
Private Declare Function WriteFileA Lib "kernel32" (ByVal hFile As Long, _
     lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
     lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
     
Private Declare Function ReadFileA Lib "kernel32" (ByVal hFile As Long, _
     lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
     lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
 
Private Declare Function GetLastError Lib "kernel32" () As Long

'\\ End


Public Script As String
Private SplitScript() As String
Private VarNames As New Collection
Private FileNames As New Collection
Private FilePaths As New Collection
Private FileContents As New Collection
Private Variables As New Collection
Public nCurrentline As Long
Public nNumCurrentLines As Long

Public Event LineChange(nLine As Long, sContents As String)
Public Event ScriptError(Msg As String, Line As Long)
Public Event DebugOut(Msg As String)
Public Event CommandOut(Command As String, ArgList() As String)
Public AllowDownloads As Boolean
Public AllowExecute As Boolean

Private IfSkip As Boolean
Private ScrExeSkip As Boolean
Private ScrExeiTemp As Integer, ScrExeiTemp2 As Integer, ScrExei As Integer
Private ScrExesTemp As String, ScrExesTemp2 As String
Private ScrExearTemp2() As String
Private InsideLoop As Collection
Private ReturnVar() As String
Private blReturn As Boolean
Private StopExecution As Boolean
Public blDebug As Boolean
Public DebugLevel As Integer
Public blShowErrors As Boolean

Public myDebug As Form
Public myParent As Form
'Public theForm As Form

Public Sub ScriptExecute(Optional blStep As Boolean = False)
Dim sCurrentLine As String

    If LenB(Script) = 0 Then Call NotifyError("No script has been loaded!", 0): Exit Sub
        
    StopExecution = False
        
    'performance increase to be had here for large scripts simply
    'set the Script variable to private and only allow access through
    'a method or function, in which case the SplitScript would be updated
    SplitScript = Split(Script, vbCrLf)
    
    If Not blStep Then nCurrentline = LBound(SplitScript)
    
    nNumCurrentLines = UBound(SplitScript)
    
    For nCurrentline = nCurrentline To UBound(SplitScript)
        If StopExecution = True Then StopExecution = False: Exit For
        sCurrentLine = TrimChar$(SplitScript(nCurrentline), Chr(9), 0)
        RaiseEvent LineChange(nCurrentline, sCurrentLine)
        Call LineExecute(sCurrentLine)
        If blStep Then nCurrentline = nCurrentline + 1: Exit Sub
        
    Next nCurrentline
    
    If blDebug And Not IsEmpty(ReturnVar) Then
        PrintDebug ("Returned variables:")
        Dim i As Long
        For i = LBound(ReturnVar) To UBound(ReturnVar)
            PrintDebug (i & ": " & ReturnVar(i))
        Next
    End If
    
    PrintDebug ("Execution complete.")
End Sub

Public Function FindAndExecuteSub(strSub As String) As Variant
Dim sCurrentLine As String
Dim lngSubLen As Long
Dim i As Long

    'loop through the current script and locate the section that is our given
    'sub, execute it, and return a variable
    
    If LenB(Script) = 0 Then Call NotifyError("No script has been loaded!", 0): Exit Function
    SplitScript = Split(Script, vbCrLf)
    nCurrentline = LBound(SplitScript)
    nNumCurrentLines = UBound(SplitScript)
    If Not Left(strSub, 1) = "@" Then strSub = "@" & strSub
    strSub = LCase$(strSub)
    lngSubLen = Len(strSub)
    blReturn = False
    
    For nCurrentline = nCurrentline To UBound(SplitScript)
    
        sCurrentLine = TrimChar$(SplitScript(nCurrentline), Chr(9), 0)
        If Left(LCase$(sCurrentLine), lngSubLen) = strSub Then 'TODO: fix this
            For i = nCurrentline + 1 To UBound(SplitScript)
                If blReturn = True Then Exit For
                sCurrentLine = TrimChar$(SplitScript(i), Chr(9), 0)
                If Left(LCase$(sCurrentLine), lngSubLen + 4) = strSub & " end" Then blReturn = True: Exit For
                RaiseEvent LineChange(nCurrentline, sCurrentLine)
                Call LineExecute(sCurrentLine)
            Next i
        End If
       
       If blReturn Then FindAndExecuteSub = ReturnVar: Exit For
    Next nCurrentline
End Function

Public Sub LineExecute(ByRef sCurrentLine As String, Optional blStep As Boolean = False)
    Dim sTemp As String
    Dim lngTemp As Long
    
        ' Get rid of comments
        ScrExeiTemp = InStrRev(sCurrentLine, "#")
        If ScrExeiTemp >= 1 Then
            lngTemp = InStrRev(sCurrentLine, Chr(34)) 'Do not include quotes AFTER the # symbol. It only checks for the last one used.
            If (lngTemp = 0) Or (lngTemp < ScrExeiTemp) Or ScrExeiTemp = 1 Then ' And (InStrRev(sCurrentLine, Chr(34), Len(sCurrentLine)) < ScrExeiTemp) Then
                sCurrentLine = TrimChar(Left(sCurrentLine, ScrExeiTemp - 1), Chr(9), 0) 'Right(sCurrentLine, Len(sCurrentLine) - ScrExeiTemp) 'Mid$(sCurrentLine, 1, ScrExeiTemp - 1)
            End If
        End If
        
        If LenB(sCurrentLine) = 0 Then GoTo SkipCurrentDo:
               
        ' Check for end of a block
        If ScrExeSkip And LCase(sCurrentLine) = "end" Then
            ScrExeSkip = False
            GoTo SkipCurrentDo:
        ElseIf ScrExeSkip = True Then
            GoTo SkipCurrentDo:
        End If
        
        ' Skipping if stuff...
        If IfSkip = True Then
            Select Case LCase(sCurrentLine)
                Case "end"
                    IfSkip = False
                Case "else"
                    IfSkip = False
            End Select
            GoTo SkipCurrentDo:
        Else
            If LCase(sCurrentLine) = "else" Then
                ScrExeSkip = True
                GoTo SkipCurrentDo:
            ElseIf LCase(sCurrentLine) = "end" Then
                GoTo SkipCurrentDo:
            End If
        End If
        
        ' Looping block
        If LCase(sCurrentLine) = "end" And InsideLoop.Count > 0 Then
            If Eval(Mid$(InsideLoop(InsideLoop.Count), InStr(1, InsideLoop(InsideLoop.Count), ":") + 1)) Then
                nCurrentline = CInt(Mid$(InsideLoop(InsideLoop.Count), 1, InStr(1, InsideLoop(InsideLoop.Count), ":") - 1)) - 1
            Else
                InsideLoop.Remove InsideLoop.Count
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Chck if its one of those crazy things
        ScrExeiTemp = InStr(1, sCurrentLine, " ")
        If ScrExeiTemp <> 0 Then
            Select Case LCase(Mid$(sCurrentLine, 1, ScrExeiTemp - 1))
                Case "if" ' Need I explain?
                    If Not Eval(ProcessFunctions(Mid$(sCurrentLine, ScrExeiTemp + 1))) Then
                        IfSkip = True
                    End If
                    
                    GoTo SkipCurrentDo:
                
                Case "return"
                ReDim ReturnVar(0)
                    blReturn = True
                    
                    Dim strArray() As String
                    strArray = GetReturnArgs(sCurrentLine)   'hacked up an argument parser, trims off the comma and spits everything out..
                        For lngTemp = LBound(strArray) To UBound(strArray)  'parse all of them, in case there are functions hidden inside
                            ScrExesTemp = ProcessFunctions(strArray(lngTemp))
                            If LenB(ScrExesTemp) > 0 Then
                                If VarExists(ScrExesTemp) Then ScrExesTemp = GetVar(ScrExesTemp)
                                If LenB(ReturnVar(0)) = 0 Then
                                    ReturnVar(0) = ScrExesTemp
                                Else
                                    ReDim Preserve ReturnVar(UBound(ReturnVar) + 1)
                                    ReturnVar(UBound(ReturnVar)) = ScrExesTemp
                                End If
                            End If
                        Next
                    GoTo SkipCurrentDo:
                
                Case "goto" ' Goto a label
                    'StopLabel = True
                    GotoScriptLine Trim(Mid$(sCurrentLine, ScrExeiTemp + 1))
                    GoTo SkipCurrentDo:
                
                Case "textblock"
                    'TODO add to list
                    'set a block of text to a var
                    ScrExesTemp = Trim$(Mid$(sCurrentLine, ScrExeiTemp))
                    sTemp = vbNullString
                    ScrExesTemp2 = vbNullString
                    lngTemp = nCurrentline + 1
                    sTemp = SplitScript(lngTemp)
                    Do While LCase$(TrimChar(sTemp, vbTab)) <> "end textblock"
                        If LenB(ScrExesTemp2) > 0 Then ScrExesTemp2 = ScrExesTemp2 & vbCrLf
                        ScrExesTemp2 = ScrExesTemp2 & sTemp
                        If lngTemp >= UBound(SplitScript) Then
                            Call NotifyError("No 'block end' for " & ScrExesTemp, lngTemp): ExitScript
                            Exit Do
                        Else
                            lngTemp = lngTemp + 1
                            sTemp = SplitScript(lngTemp)
                        End If
                    Loop
                
                    AddVar ScrExesTemp, ScrExesTemp2
                    nCurrentline = lngTemp
                    GoTo SkipCurrentDo:

                Case "var", "dim" ' Create variable
                    ScrExearTemp2 = Split(Trim(Mid$(sCurrentLine, ScrExeiTemp + 1)), ",")
                    For ScrExei = LBound(ScrExearTemp2) To UBound(ScrExearTemp2)
                        If VarExists(ScrExearTemp2(ScrExei)) Then
                            Call NotifyError("Variable has already been defined: " & ScrExearTemp2(ScrExei), nCurrentline)
                        Else
                            AddVar ScrExearTemp2(ScrExei), ""
                        End If
                        Next ScrExei
                        
                    GoTo SkipCurrentDo:
                
                Case "revar", "redim"
                    If Mid(LCase$(sCurrentLine), 7, 8) = "preserve" Then
                        'get the ubound of the array and add or subtract any that are not included
                        'need to create a RemoveVar function to simplify this
                        lngTemp = InStr(15, sCurrentLine, "[")
                        ScrExesTemp = Mid(sCurrentLine, 16, (lngTemp - 16))
                        sTemp = Mid(sCurrentLine, lngTemp + 1, InStrRev(sCurrentLine, "]") - lngTemp - 1)
                        If NumberOfFunctions(sTemp) > -1 Then
                            sTemp = DoFunction(sTemp)
                        End If
                        
                        If Not IsNumeric(sTemp) Then RaiseEvent ScriptError("Unable to redim var " & ScrExesTemp & ": value isn't numerical", nCurrentline): Exit Sub
                        
                        If VarExists("Ubound_" & ScrExesTemp) Then
                            ScrExeiTemp2 = GetVar("Ubound_" & ScrExesTemp) + 1
                            If ScrExeiTemp2 > sTemp Then
                                For ScrExeiTemp = ScrExeiTemp2 To sTemp
                                    AddVar ScrExesTemp & "[" & ScrExeiTemp & "]", ""
                                Next ScrExeiTemp
                            ElseIf ScrExeiTemp < sTemp Then
                                For ScrExeiTemp = sTemp To ScrExeiTemp2 Step -1
                                    RemoveVar ScrExesTemp & "[" & ScrExeiTemp & "]"
                                Next ScrExeiTemp
                            End If
                        Else
                            'the array doesn't exist.
                            'do a loop through all of them
                            For ScrExeiTemp = 0 To sTemp
                                AddVar ScrExesTemp & "[" & ScrExeiTemp & "]", ""
                            Next ScrExeiTemp
                            AddVar "Ubound_" & ScrExesTemp, sTemp
                            AddVar ScrExesTemp & "[0]", GetVar(ScrExesTemp)
                            AddVar ScrExesTemp, sTemp
                        End If
                        
                    ElseIf InStr(1, sCurrentLine, "[") > 0 Then
                        sTemp = Trim$(Mid(sCurrentLine, InStr(1, sCurrentLine, " ") + 1))
                        lngTemp = CLng(GetVar("Ubound_" & sTemp))
                        For ScrExeiTemp = 0 To lngTemp
                            RemoveVar sTemp & "[" & ScrExeiTemp & "]"
                        Next
                        ChangeVar sTemp, ""
                    
                    Else
                        lngTemp = InStr(7, sCurrentLine, "[")
                        ScrExesTemp = Mid(sCurrentLine, 7, (lngTemp - 7))
                        sTemp = Mid(sCurrentLine, lngTemp + 1, InStrRev(sCurrentLine, "]") - lngTemp - 1)
                        If NumberOfFunctions(sTemp) > -1 Then
                            'ScrExesTemp
                            sTemp = DoFunction(sTemp)
                        End If
                                            
                        If Not IsNumeric(sTemp) Then RaiseEvent ScriptError("Unable to redim var " & ScrExesTemp & ": value isn't numerical", nCurrentline): Exit Sub
                        AddVar ScrExesTemp, sTemp
                        AddVar "Ubound_" & ScrExesTemp, sTemp
                        For ScrExeiTemp = 0 To sTemp
                            AddVar ScrExesTemp & "[" & ScrExeiTemp & "]", ""
                        Next ScrExeiTemp
                                            
                    End If
                    GoTo SkipCurrentDo:
                    
                Case "while" ' Loop
                    If Eval(Mid$(sCurrentLine, ScrExeiTemp + 1)) Then
                        InsideLoop.add nCurrentline + 1 & ":" & Trim(Mid$(sCurrentLine, ScrExeiTemp + 1))
                    Else
                        ScrExeSkip = True
                    End If
                    GoTo SkipCurrentDo:
            End Select
        End If
        
        For ScrExei = 1 To Len(sCurrentLine)
            Select Case Mid$(sCurrentLine, ScrExei, 1)
            
                Case "@"
                    'it's a sub/function
                    Debug.Print "ho?"
                    'todo: Need to move this up and check if we're about to enter inside of a
                    'function and whether or not we are supposed to be in the function we've found..
                    GoTo SkipCurrentDo:
            
                Case "+"
                    If Mid$(sCurrentLine, (ScrExei + 1), 1) = "=" Then
                        ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "+") - 1))
                        ScrExesTemp2 = GetVar(ScrExesTemp)
                        If ScrExesTemp2 = "" Then ScrExesTemp2 = 0
                        ChangeVar ScrExesTemp, Equation(Chr(34) & GetVar(ScrExesTemp) & Chr(34) & "+" & Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                        GoTo SkipCurrentDo:
                    End If
                    
                Case "-"
                    If Mid$(sCurrentLine, (ScrExei + 1), 1) = "=" Then
                        ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "-") - 1))
                        ScrExesTemp2 = GetVar(ScrExesTemp)
                        If ScrExesTemp2 = "" Then ScrExesTemp2 = 0
                        ChangeVar ScrExesTemp, Equation(Chr(34) & GetVar(ScrExesTemp) & Chr(34) & "-" & Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                        GoTo SkipCurrentDo:
                    End If
                    
                Case "*"
                    If Mid$(sCurrentLine, (ScrExei + 1), 1) = "*" Then
                        ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "+") - 1))
                        ScrExesTemp2 = GetVar(ScrExesTemp)
                        If ScrExesTemp2 = "" Then ScrExesTemp2 = 0
                        ChangeVar ScrExesTemp, Equation(Chr(34) & GetVar(ScrExesTemp) & Chr(34) & "*" & Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                        GoTo SkipCurrentDo:
                    End If
                    
                Case "/"
                    If Mid$(sCurrentLine, (ScrExei + 1), 1) = "=" Then
                        ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "/") - 1))
                        ScrExesTemp2 = GetVar(ScrExesTemp)
                        If ScrExesTemp2 = "" Then ScrExesTemp2 = 0
                        ChangeVar ScrExesTemp, Equation(Chr(34) & GetVar(ScrExesTemp) & Chr(34) & "/" & Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                        GoTo SkipCurrentDo:
                    End If
                    
                Case "\"
                    If Mid$(sCurrentLine, (ScrExei + 1), 1) = "=" Then
                        ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "\") - 1))
                        ScrExesTemp2 = GetVar(ScrExesTemp)
                        If ScrExesTemp2 = "" Then ScrExesTemp2 = 0
                        ChangeVar ScrExesTemp, Equation(Chr(34) & GetVar(ScrExesTemp) & Chr(34) & "\" & Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                        GoTo SkipCurrentDo:
                    End If
                    
                Case "^"
                    If Mid$(sCurrentLine, (ScrExei + 1), 1) = "=" Then
                        ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "^") - 1))
                        ScrExesTemp2 = GetVar(ScrExesTemp)
                        If ScrExesTemp2 = "" Then ScrExesTemp2 = 0
                        ChangeVar ScrExesTemp, Equation(Chr(34) & GetVar(ScrExesTemp) & Chr(34) & "^" & Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                        GoTo SkipCurrentDo:
                    End If
                
                Case "=" ' Assignment
                    sTemp = Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1)
                    If NumberOfFunctions(Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1)) > -1 Then sTemp = DoFunction(sTemp)

                    ChangeVar Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "=") - 1)), Equation(sTemp)

                    GoTo SkipCurrentDo:
                Case "~" ' Assignment from Function
                    ScrExesTemp = Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "~") - 1))
                    ChangeVar Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "~") - 1)), DoFunction(Mid$(sCurrentLine, InStr(1, sCurrentLine, "~") + 1), ScrExesTemp)
                    GoTo SkipCurrentDo:
                Case "(" ' Procedure Call
                    DoFunction sCurrentLine
                    GoTo SkipCurrentDo:
            End Select
        Next ScrExei
        
        Call NotifyError("Line not processed: " & sCurrentLine, (ScrExei + 1))
        
SkipCurrentDo:
    If blStep Then nCurrentline = nCurrentline + 1: Exit Sub
        
End Sub
    
Private Function ExecFunction(sFunction As String, ArgList() As String, Optional CallingVar As String = "") As Variant
    On Error GoTo ErrorCatch:
    Dim arrTemp() As String
    Dim i As Long
    Dim blHandled As Boolean
    Dim strFile As String
    Dim intTemp As Integer
    Dim strTemp As String
    Dim blTemp As Boolean
    If Left(sFunction, 1) = "~" Then sFunction = Right(sFunction, Len(sFunction) - 1)
    
        PrintDebug "Executing Function: " & sFunction & "(" & Join(ArgList, ", ") & ")", 1
    
    Select Case LCase$(sFunction)
        Case "msgbox"
        blHandled = True
            'ExecFunction = MsgBox(ArgList(0), ArgList(2), ArgList(1))
            Select Case UBound(ArgList)
                Case Is = 0
                    ExecFunction = MsgBox(ArgList(0), vbOKOnly, "Message")
                Case Is = 1
                    ExecFunction = MsgBox(ArgList(0), vbOKOnly, ArgList(1))
                Case Is = 2
                    ExecFunction = MsgBox(ArgList(0), ArgList(2), ArgList(1))
            End Select
            
        Case "inputbox"
        blHandled = True
            Select Case UBound(ArgList)
                Case Is = 0
                    ExecFunction = InputBox(ArgList(0))
                Case Is = 1
                    ExecFunction = InputBox(ArgList(0), ArgList(1))
                Case Is = 2
                    ExecFunction = InputBox(ArgList(0), ArgList(1), ArgList(2))
            End Select
            
        Case "app"
        blHandled = True
            Select Case LCase(ArgList(0))
            
                Case Is = "major"
                    ExecFunction = App.Major
                Case Is = "minor"
                    ExecFunction = App.Minor
                Case Is = "revision"
                    ExecFunction = App.Revision
                Case Is = "exename"
                    ExecFunction = App.EXEName
                Case Is = "path"
                    ExecFunction = App.Path
                Case Is = "title"
                    ExecFunction = App.Title
                Case Is = "threadid"
                    ExecFunction = App.ThreadID
            End Select
                       
        Case "stop"
            blHandled = True
            ExitScript
            
        Case "shutdown"
            End
            
        ' String Manipulating Junk
        Case "mid"
            blHandled = True
            ExecFunction = Mid$(ArgList(0), ArgList(1), ArgList(2))
        Case "lcase"
            blHandled = True
            ExecFunction = LCase(ArgList(0))
        Case "ucase"
            blHandled = True
            ExecFunction = UCase(ArgList(0))
        Case "instr"
            blHandled = True
            ExecFunction = InStr(ArgList(0), ArgList(1), ArgList(2))
        Case "instrrev"
            blHandled = True
            ExecFunction = InStrRev(ArgList(0), ArgList(1))
        Case "strcomp"
            blHandled = True
            ExecFunction = StrComp(ArgList(0), ArgList(1))
        Case "safestr", "formatstr"
            blHandled = True
            ExecFunction = FormatStrForVar(ArgList(0))
        Case "len"
            blHandled = True
            ExecFunction = Len(ArgList(0))
        Case "left"
            blHandled = True
            ExecFunction = Left(ArgList(0), ArgList(1))
        Case "right"
            blHandled = True
            ExecFunction = Right(ArgList(0), ArgList(1))
        Case "chr"
            blHandled = True
            ExecFunction = Chr(ArgList(0))
            'Exit Function
        Case "asc"
            blHandled = True
            ExecFunction = Asc(ArgList(0))
        Case "reverse"
            blHandled = True
            ExecFunction = StrReverse(ArgList(0))
        Case "replace"
            blHandled = True
            ExecFunction = Replace(ArgList(0), ArgList(1), ArgList(2))
        Case "trimchar"
            blHandled = True
            Select Case UBound(ArgList)
                Case Is = 1
                    ExecFunction = TrimChar(ArgList(0), Mid(ArgList(1), 1, 1))
                Case Is = 2
                    ExecFunction = TrimChar(ArgList(0), ArgList(1), CInt(ArgList(2)))
            End Select
            Exit Function
        Case "varexists", "variableexists"
            blHandled = True
            If LenB(CallingVar) > 0 Then ArgList(0) = CallingVar
            ExecFunction = VarExists(ArgList(0))
            
        Case "split"
            blHandled = True
            If LenB(CallingVar) > 0 Then
                strTemp = CallingVar
                arrTemp = Split(ArgList(0), ArgList(1))
            Else
                strTemp = ArgList(0)
                arrTemp = Split(ArgList(1), ArgList(2))
            End If
                For i = 0 To UBound(arrTemp)
                    AddVar strTemp & "[" & i & "]", arrTemp(i)
                Next i
                'Either call the reference varname itself, or add Ubound_ to make it easier to understand in script
                AddVar "Ubound_" & strTemp, UBound(arrTemp)
                AddVar strTemp, UBound(arrTemp)
                ExecFunction = UBound(arrTemp)
            
        Case "join"
            blHandled = True
            If Not VarExists("Ubound_" & ArgList(0)) Then _
                RaiseEvent ScriptError("Variable is not an array: " & ArgList(0), nCurrentline): Exit Function
            
            ReDim arrTemp(GetVar("Ubound_" & ArgList(0)))
            For i = LBound(arrTemp) To UBound(arrTemp)
                arrTemp(i) = GetVar(ArgList(0) & "[" & i & "]")
            Next i
            Select Case UBound(ArgList)
                Case Is = 0
                    ExecFunction = Join(arrTemp, "")
                Case Is = 1
                    ExecFunction = Join(arrTemp, ArgList(1))
            End Select
        
        Case "replacestr"
            blHandled = True
            strFile = LoadFile(ArgList(0))
            intTemp = ArgList(ArgList(5))
            strTemp = Replace(strFile, ArgList(1), ArgList(2), ArgList(3), ArgList(4))
            WriteFile strFile, ArgList(0), True
            ExecFunction = True
        
        Case "ubound"
            blHandled = True
            ExecFunction = GetVar("Ubound_" & ArgList(0))
        
        ' Math stuff
        Case "cos"
            blHandled = True
            ExecFunction = Cos(ArgList(0))
        Case "tan"
            blHandled = True
            ExecFunction = Tan(ArgList(0))
        Case "log"
            blHandled = True
            ExecFunction = Log(ArgList(0))
        Case "rnd"
            blHandled = True
            ExecFunction = Rnd()
        
        Case "randomnumber"
            blHandled = True
            ExecFunction = Int((CLng(ArgList(0)) * Rnd) + 1)
            Exit Function
            
        Case "randomize"
            blHandled = True
            Randomize
            ExecFunction = True
        Case "int"
            blHandled = True
            ExecFunction = Int(ArgList(0))
            
        Case "val"
            blHandled = True
            ExecFunction = Val(ArgList(0))
        
        Case "gettickcount"
            ExecFunction = GetTickCount()
        
        Case "execute"
            blHandled = True
            If AllowExecute Then ExecFunction = False: Exit Function
            If UBound(ArgList) = 1 Then
                ExecFunction = ExecFile(ArgList(0), ArgList(1))
            ElseIf UBound(ArgList) = 2 Then
                ExecFunction = ExecFile(ArgList(0), ArgList(1), , ArgList(2))
            ElseIf UBound(ArgList) = 3 Then
                ExecFunction = ExecFile(ArgList(0), ArgList(1), , ArgList(2), CLng(ArgList(3)))
            End If
            
            ExecFunction = True
        
        Case "shell"
            blHandled = True
            If AllowExecute Then ExecFunction = False: Exit Function
            Shell ArgList(0)
            ExecFunction = True
        
        Case "shellwait"
            'TODO add to list
            Dim idProg As Long
            blHandled = True
            If AllowExecute = False Then ExecFunction = False: Exit Function
            If UBound(ArgList) = 0 Then
                idProg = Shell(ArgList(0))
            Else
                idProg = Shell(ArgList(0) & " " & ArgList(1))
            End If
            ExecFunction = fWait(idProg)
        
        Case "newline", "crlf"
            'todo add to list
            blHandled = True
            ExecFunction = vbNewLine
            
        Case "clipboard"
            blHandled = True
            Select Case ArgList(0)
                Case Is = "clear"
                    Clipboard.Clear
                    ExecFunction = True
                Case Is = "get", "gettext"
                    ExecFunction = Clipboard.GetText(ArgList(1))
                Case Is = "set", "settext"
                    Clipboard.SetText ArgList(1)
                    ExecFunction = True
            End Select
        
        Case "time"
            blHandled = True
            ExecFunction = Time
            
        Case "now"
            blHandled = True
            ExecFunction = Now
        
        Case "doevents"
            blHandled = True
            DoEvents
            ExecFunction = True
            
        Case "pause"
            blHandled = True
            Pause CLng(ArgList(0))
            ExecFunction = True
                        
        Case "print"
            blHandled = True
            PrintDebug ArgList(0), 0
            RaiseEvent DebugOut(ArgList(0))
            ExecFunction = True
            Exit Function
            
        Case "debug"
            blHandled = True
            Select Case LCase$(ArgList(0))
            
                Case Is = "on", "true"
                    blDebug = True
                Case Is = "off", "false"
                    blDebug = False
                Case Is = "show"
                    ShowDebug
                Case Is = "hide"
                    blDebug = False
                    myDebug.Visible = False
                    Unload myDebug
                Case Is = 0
                    DebugLevel = 0
                Case Is = 1
                    DebugLevel = 1
                Case Is = 2
                    DebugLevel = 2
                Case Is = 3
                    DebugLevel = 3
                Case Is = 4
                    DebugLevel = 3
                Case Is = "clear"
                    If Not myDebug Is Nothing Then Call myDebug.Clear
            End Select
            ExecFunction = True
            Exit Function
            
        Case "showerrors", "errors"
            blHandled = True
            Select Case LCase$(ArgList(0))
            
                Case Is = "on", "true", 1
                    blShowErrors = True
                Case Is = "off", "false", 0
                    blShowErrors = False
            End Select
            
            PrintDebug "Show Errors: " & blDebug, 0
            Exit Function
            
        'Case "getname", "username"
        '    blHandled = True
        '    ExecFunction = "Bradley"
        '
        'Case "getgamepath"
        '    blHandled = True
        '    ExecFunction = App.Path & "\test\"
        
        Case "getscreenwidth", "getx"
            blHandled = True
            ExecFunction = Screen.Width / Screen.TwipsPerPixelX
            Exit Function
            
        Case "getscreenheight", "gety"
            blHandled = True
            ExecFunction = Screen.Height / Screen.TwipsPerPixelY
            Exit Function
            
        Case "getos"
            blHandled = True
            ExecFunction = GetOSVer
            
        'Case "getip"
        '    blHandled = True
        '    ExecFunction = "127.0.0.1"
        '    Exit Function
            
        Case "is64bit"
            blHandled = True
            If IsHost64Bit Then ExecFunction = True
                
        Case "filesize", "getsize", "filelen"
            blHandled = True
            ExecFunction = FileLen(ArgList(0))
            Exit Function
            
        Case "fileexists", "exists"
            blHandled = True
            ExecFunction = FileExists(ArgList(0))
            ExecFunction = True
            
        Case "folderexists", "direxists"
            blHandled = True
            ExecFunction = DirExists(ArgList(0))
            ExecFunction = True
            
        Case "filecopy", "copy"
            blHandled = True
            FileCopy ArgList(0), ArgList(1)
            ExecFunction = True
            
        Case "deletefolder", "rmdir"
            blHandled = True
             RmDir ArgList(0)
             ExecFunction = True
             
        Case "deletefile", "kill"
            blHandled = True
            Kill ArgList(0)
            ExecFunction = True
            
        Case "createfolder", "mkdir"
            blHandled = True
            MkDir ArgList(0)
            ExecFunction = True
            
        Case "renamefolder", "renfolder"
            blHandled = True
            Name ArgList(0) As ArgList(1)
            ExecFunction = True
            
        Case "renamefile", "renfile"
            blHandled = True
            Name ArgList(0) As ArgList(1)
            ExecFunction = True
            
        Case "createfile", "newfile"
        blHandled = True
            WriteFile ArgList(0), ArgList(1), True
            ExecFunction = True
            
        Case "appendfile"
            blHandled = True
            WriteFile ArgList(0), ArgList(1), False
            ExecFunction = True
        
        Case "filetovar"
            blHandled = True
            If LenB(CallingVar) > 0 Then
                AddVar CallingVar, LoadFile(ArgList(0))
            Else
                AddVar ArgList(0), LoadFile(ArgList(1))
            End If
            ExecFunction = True
            Exit Function
        
        Case "copyvar"
            blHandled = True
            If LenB(CallingVar) > 0 Then
                'todo
                AddVar CallingVar, GetVar(ArgList(0))
            Else
                AddVar ArgList(0), GetVar(ArgList(1))
            End If
            ExecFunction = True
            Exit Function
        
        Case "downloadfile"
            blHandled = True
            If AllowDownloads Then
                ExecFunction = DownloadFile(ArgList(0), ArgList(1))
                ExecFunction = True
            End If
            
        Case "loadfile", "readfile"
            blHandled = True
            If LenB(CallingVar) > 0 Then
                AddFile CallingVar, LoadFile(ArgList(0))
            Else
                AddFile ArgList(0), LoadFile(ArgList(1))
            End If
            ExecFunction = True
            Exit Function
                      
        Case "isfileloaded"
            blHandled = True
            ExecFunction = IsFileLoadedByName(ArgList(0))

        Case "findposinfile"
            blHandled = True
            'find a str in file and return its position
            If UBound(ArgList) = 1 Then
                ExecFunction = FindFilePos(ArgList(0), ArgList(1), 1)
            Else
                ExecFunction = FindFilePos(ArgList(0), ArgList(1), CLng(ArgList(2)))
            End If
            'Exit Function
                        
        Case "findline", "findlinebystr"
            blHandled = True
            ExecFunction = FindFileLine(ArgList(0), ArgList(1), CLng(ArgList(2)))
            'Exit Function
            
        Case "replaceline", "replacelineinfile"
            blHandled = True
            SetFileLine ArgList(0), ArgList(1), CLng(ArgList(2))
            ExecFunction = True
            
        Case "getfileline", "getline"
            blHandled = True
            ExecFunction = GetFileLine(ArgList(0), CLng(ArgList(1)))
        
        Case "addfileline", "addline"
            blHandled = True
            AddFileLine ArgList(0), ArgList(1), CLng(ArgList(2))
            ExecFunction = True
            
        Case "remfileline", "removeline"
            blHandled = True
            RemoveFileLine ArgList(0), CLng(ArgList(1))
            ExecFunction = True
            
        Case "insertfilepos", "insertfile"
            blHandled = True
            InsertFilePos ArgList(0), ArgList(1), CLng(ArgList(2))
            ExecFunction = True
            
        Case "savefilevar", "savefile"
            blHandled = True
            strTemp = GetFile(ArgList(0))
            WriteFile strTemp, ArgList(1), True
            ExecFunction = True
            
        Case "writeini"
            blHandled = True
            WriteIni ArgList(0), ArgList(1), ArgList(2), ArgList(3)
            ExecFunction = True
            
        Case "readini"
            blHandled = True
            ExecFunction = ReadIni(ArgList(0), ArgList(1), ArgList(2))
            Exit Function
            
        Case "getdirectories", "dirlist"
            blHandled = True
            If LenB(CallingVar) = 0 Then
                If UBound(ArgList) >= 1 Then
                    CallingVar = ArgList(0)
                    strFile = ArgList(1)
                    If UBound(ArgList) = 2 Then blTemp = CBool(ArgList(2)) Else blTemp = True
                End If
            Else
                strFile = ArgList(0)
                If UBound(ArgList) = 1 Then blTemp = CBool(ArgList(1)) Else blTemp = True
            End If
            
            If UBound(ArgList) = 2 Then blTemp = CBool(ArgList(2)) Else blTemp = True
            
            strTemp = Dir(strFile, vbDirectory)
            i = 0
                        
            If LenB(strFile) = 0 Then strFile = App.Path
            If Right(strFile, 1) <> "\" Then strFile = strFile & "\"
            
            Do While strTemp > ""
                If GetAttr(strTemp) And vbDirectory Then
                    If strTemp <> "." And strTemp <> ".." Then
                    If blTemp Then
                        AddVar CallingVar & "[" & i & "]", strFile & strTemp & "\"
                    Else
                        AddVar CallingVar & "[" & i & "]", strTemp
                    End If
                        i = i + 1
                    End If
                End If
                strTemp = Dir()
            Loop
            
            i = i - 1
            AddVar "Ubound_" & CallingVar, i
            AddVar CallingVar, i
            ExecFunction = True
            
        Case "getfilelist", "filelist"
            blHandled = True
            If LenB(CallingVar) = 0 Then
                If UBound(ArgList) >= 1 Then
                    CallingVar = ArgList(0)
                    strFile = ArgList(1)
                    If UBound(ArgList) = 2 Then blTemp = CBool(ArgList(2)) Else blTemp = True
                End If
            Else
                strFile = ArgList(0)
                If UBound(ArgList) = 1 Then blTemp = CBool(ArgList(1)) Else blTemp = True
            End If
            
            
            
            strTemp = Dir(strFile)
            i = 0
            
            If LenB(strFile) = 0 Then strFile = App.Path
            If Right(strFile, 1) <> "\" Then strFile = strFile & "\"
            
            Do While strTemp > ""
                    If blTemp Then
                        AddVar CallingVar & "[" & i & "]", strFile & strTemp
                    Else
                        AddVar CallingVar & "[" & i & "]", strTemp
                    End If
                    i = i + 1
                strTemp = Dir()
            Loop
            
            i = i - 1
            AddVar "Ubound_" & CallingVar, i
            AddVar CallingVar, i
            ExecFunction = True
            
        Case "chdir", "workingdir"
            blHandled = True
            ChDir ArgList(0)
            ExecFunction = True
        'Case Else
            'call NotifyError("Unknown Function: " & sFunction, nCurrentLine)

    End Select
    
    If Not blHandled Then NotifyError "Unhandled Function: " & sFunction, nCurrentline: Exit Function
    ExecFunction = Chr(34) & ExecFunction & Chr(34)
    
    Exit Function
    
ErrorCatch:
    Select Case err.Number
        Case 9
            Call NotifyError("Argument not optional in call to procedure: " & sFunction, nCurrentline)
    End Select
    ExecFunction = False
    'Resume Next
End Function

Private Sub AddFile(fName As String, fContents As String)
'Private FileNames As New Collection
'Private FilePaths As New Collection
'Private FileContents As New Collection

    Dim i As Long
    For i = FileNames.Count To 1 Step -1
        If FileNames(i) = fName Then
            PrintDebug "AddFile: " & fName & " from " & GetFile(fName) & " to: " & fContents, 4

            FileNames.Remove i
            FileContents.Remove i
            
            FileNames.add fName
            FileContents.add fContents
            Exit Sub
        End If
    Next i
    
    FileNames.add fName
    FileContents.add fContents
    
    PrintDebug "AddFile: " & fName & " didn't exist, created it. (there's no way to dim file vars currently)", 1
End Sub

Private Function IsFileLoadedByName(fName As String) As Boolean
    Dim i As Long
    For i = FileNames.Count To 1 Step -1
        If FileNames(i) = fName Then
            PrintDebug "FileLoadedByName (" & fName & ") = True", 3
            IsFileLoadedByName = True
            Exit Function
        End If
    Next i
    PrintDebug "FileLoadedByName (" & fName & ") = False", 1
End Function

Private Function GetFile(fName As String) As String
    Dim i As Long
    Dim strTemp As String
    Dim strArrVar As String
    
    For i = 1 To FileNames.Count
        If FileNames(i) = fName Then
            GetFile = FileContents(i)
            PrintDebug "GetFile: " & fName & ": " & FileContents(i), 3
            Exit Function
        End If
    Next i

End Function

Private Function FindFilePos(fName As String, strFind As String, lngStartPos As Long) As Long
    Dim strFile As String
    Dim lngPos As Long
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function

    lngPos = InStr(lngStartPos, strFile, strFind)
    
    FindFilePos = lngPos
        
End Function

Private Function InsertFilePos(fName As String, strInsert As String, lngInsertPos As Long)

    Dim strFile As String
    Dim strTemp As String
    Dim strTemp2
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function
    strTemp = Mid(strFile, 1, lngInsertPos)
    strTemp2 = Mid(strFile, lngInsertPos)
    strFile = strTemp & strInsert & strTemp2
    AddFile fName, strFile
        
End Function

Private Function FindFileLine(fName As String, strFind, lngStartLine As Long) As Long
    Dim strTemp() As String
    Dim strFile As String
    Dim i As Long
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function
    strTemp = Split(strFile, vbCrLf)
        
        For i = IIf(lngStartLine > 0, lngStartLine, LBound(strTemp)) To UBound(strTemp)
            If InStr(1, strTemp(i), strFind) > 0 Then FindFileLine = i: Exit Function
        Next i
        
End Function

Private Function GetFileLine(fName As String, fLine As Long) As String
    Dim strTemp() As String
    Dim strFile As String
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function
    strTemp = Split(strFile, vbCrLf)
    GetFileLine = strTemp(fLine)
End Function

Private Function SetFileLine(fName As String, fString As String, fLine As Long) As Boolean

    Dim strTemp() As String
    Dim strFile As String
    
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function
    strTemp = Split(strFile, vbCrLf)
    
    strTemp(fLine) = fString
    
    strFile = Join(strTemp, vbCrLf)
    
    AddFile fName, strFile
End Function

Private Function RemoveFileLine(fName As String, fLine As Long)

    Dim strTemp() As String
    Dim strFile As String
    Dim i As Long
    
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function
    strTemp = Split(strFile, vbCrLf)
    
    For i = LBound(strTemp) To UBound(strTemp)
    
        If i <> fLine Then strFile = strFile & strTemp(i) & vbCrLf

    Next i

    AddFile fName, strFile

End Function

Private Function AddFileLine(fName As String, fString As String, fAddBefore As Long)

    Dim strTemp() As String
    Dim strFile As String
    Dim i As Long
    
    strFile = GetFile(fName)
    If LenB(strFile) = 0 Then Exit Function
    strTemp = Split(strFile, vbCrLf)
    strFile = vbNullString
    
    For i = LBound(strTemp) To UBound(strTemp)
    
        If i <> fAddBefore Then
            strFile = strFile & strTemp(i) & vbCrLf
        Else
            strFile = strFile & fString & vbCrLf & strTemp(i) & vbCrLf
        End If

    Next i

    AddFile fName, strFile
End Function


Private Function RawFileInsert(fPath As String, lngStart As Long, strData As String) As Boolean
On Error GoTo RawInsErr
    Dim FHandle As Long
    FHandle = FreeFile
        
    Open fPath For Binary Access Read Write As #FHandle
        
        InsertData FHandle, strData, lngStart

    Close #FHandle
    
    RawFileInsert = True
    Exit Function

RawInsErr:
RaiseEvent ScriptError("RawInsWrite Error: " & err.Number & ": " & err.Description, nCurrentline)
End Function

Private Function RawFileDelete(fPath As String, lngStart As Long, lngLength As Long)
On Error GoTo RawFileDelErr
    Dim FHandle As Long
    FHandle = FreeFile
        
    Open fPath For Binary Access Read Write As #FHandle
        
        DeleteData FHandle, lngStart, lngLength

    Close #FHandle
    
    RawFileDelete = True
    Exit Function

RawFileDelErr:
RaiseEvent ScriptError("RawFileDelete Error: " & err.Number & ": " & err.Description, nCurrentline)
End Function

Private Function RawFileRead(fPath As String, lngStart As Long, lngLength As Long) As String
    Dim FHandle As Long
    Dim Buffer As String
    If lngLength > 0 Then
        Buffer = Space$(lngLength)
    Else
        Buffer = Space$(FileLen(fPath) - lngStart)
    End If
    FHandle = FreeFile
    If lngStart = 0 Then lngStart = 1
    
    Open fPath For Binary Access Read As #FHandle
        
        Get #FHandle, lngStart, Buffer
    
    Close #FHandle
    
    RawFileRead = Buffer
    
End Function

Private Function RawFileWrite(fPath As String, strData As String, Optional lngStart As Long = -1) As Boolean
On Error GoTo RawWriteErr
    Dim FHandle As Long
    FHandle = FreeFile
        
    
        
    Open fPath For Binary Access Write As #FHandle
    
        If lngStart = -1 Then lngStart = LOF(FHandle)
        Put #FHandle, lngStart, strData
    
    Close #FHandle
    
    RawFileWrite = True
    Exit Function

RawWriteErr:
RaiseEvent ScriptError("RawFileWrite Error: " & err.Number & ": " & err.Description, nCurrentline)
End Function


Public Sub AddVar(TheVar As String, NewVal As Variant)
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            PrintDebug "AddVar " & TheVar & ": [" & NewVal & "] (Already existed)", 3
            RemoveVar TheVar
            
            Variables.add NewVal
            VarNames.add TheVar
            Exit Sub
        End If
    Next i
    
    'We didn't find a matching variable.. so let's add one
    Variables.add NewVal
    VarNames.add TheVar
    PrintDebug "AddVar " & TheVar & ": [" & NewVal & "]", 3
    
End Sub

Private Sub RemoveVar(TheVar As String)
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            PrintDebug "RemoveVar: " & TheVar, 2
            VarNames.Remove i
            Variables.Remove i
            Exit Sub
        End If
    Next i
End Sub

Private Function VarExists(VarName As String) As Boolean
    Dim i As Long
    If InStr(1, VarName, " ") Then VarExists = False: Exit Function
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = VarName Then
        PrintDebug "VarExists (" & VarName & ") = True", 3
            VarExists = True
            Exit Function
        End If
    Next i
    PrintDebug "VarExists (" & VarName & ") = False", 1
End Function

Private Sub ChangeVar(TheVar As String, NewVal As Variant)
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            PrintDebug "ChangeVar: " & TheVar & " from " & Variables(i) & " to: " & NewVal, 2

            Variables.Remove i
            VarNames.Remove i
            
            Variables.add NewVal
            VarNames.add TheVar
            Exit Sub
        End If
    Next i
    
    PrintDebug "ChangeVar: " & TheVar & " doesn't exist!", 1
    
End Sub

Private Function GetVar(TheVar As String) As Variant
    Dim i As Long
    Dim strTemp As String
    Dim strArrVar As String
    'If Len(TheVar) < 3 Then GetVar = TheVar: Exit Function
    'PrintDebug "GetVar: " & TheVar
    strTemp = Replace(TheVar, Chr(34), "")
    If strTemp = "True" Or strTemp = "False" Then GetVar = TheVar: Exit Function
    
    i = InStr(1, TheVar, "[")
    If i >= 1 Then
        strArrVar = Left(TheVar, i - 1)
        strTemp = Mid(TheVar, i + 1, InStrRev(TheVar, "]") - i - 1)
        strTemp = GetVar(strTemp)
        If IsNumeric(strTemp) Then
            TheVar = strArrVar & "[" & strTemp & "]"
        End If
    Else
        If InStr(1, TheVar, " ") > 0 Then GetVar = TheVar: Exit Function
        If IsNumeric(strTemp) Then GetVar = strTemp: Exit Function
    End If
    
    For i = 1 To VarNames.Count
        If VarNames(i) = TheVar Then
            GetVar = Variables(i)
            PrintDebug "GetVar: " & TheVar & ": " & Variables(i), 2
            Exit Function
        End If
    Next i
    
    GetVar = TheVar
    
    'If InStr(1, TheVar, Chr(34)) > 0 Then GetVar = TrimChar(TheVar, Chr(34), 1): Exit Function
    
End Function

Private Function FormatStrForVar(strString As String) As String
    Dim blTrimLeft As Boolean
    Dim blTrimRight As Boolean
    Dim i As Long
    Dim strSplit() As String
    
    strSplit = Split(strString, vbCrLf)
    
    For i = LBound(strSplit) To UBound(strSplit)
    
        If Left(strSplit(i), 1) = Chr(34) And Right(strSplit(i), 1) = Chr(34) Then 'if first and last char are quotes, we need to do something different
            blTrimLeft = True
            blTrimRight = True
        ElseIf Right(strSplit(i), 1) = Chr(34) Then
            blTrimRight = True
        ElseIf Left(strSplit(i), 1) = Chr(34) Then
            blTrimLeft = True
        Else
            FormatStrForVar = Replace(strSplit(i), Chr(34), Chr(34) & " & " & Chr(34) & Chr(34) & Chr(34) & " & " & Chr(34))
        End If
        
        If blTrimLeft Then
            strSplit(i) = Right(strSplit(i), 4): blTrimLeft = False
        End If
        
        If blTrimRight Then
            strSplit(i) = Left(strSplit(i), Len(strSplit(i)) - 4): blTrimRight = False
        End If
        
        strSplit(i) = Chr(34) & strSplit(i) & Chr(34)
    Next i

    FormatStrForVar = Join(strSplit, vbCrLf)

End Function

Private Function Eval(ByVal sFormula As String) As Boolean
    Dim i As Long, iWait As Integer
    Dim LeftVal As String, RightVal As String, Operator As String
    Dim sTemp As String
    Dim blLeftNot As Boolean
    Dim blRightNot As Boolean
    Dim blEval As Boolean
    
    If LenB(sFormula) = 0 Then Exit Function
    
    PrintDebug "Eval: " & sFormula, 2
    
    ' Get the left value
    For i = 1 To Len(sFormula)
        sTemp = Mid$(sFormula, i, 1)
        Select Case sTemp
            Case "("
                iWait = iWait + 1
            Case ")"
                iWait = iWait - 1
                If iWait = 0 Then
                    LeftVal = Mid$(sFormula, 1, i)
                    sFormula = Trim(Mid$(sFormula, i + 1))
                    Exit For
                End If
            Case Chr(34)
                i = InStr(i + 1, sFormula, Chr(34))
                If iWait = 0 Then
                    LeftVal = Mid$(sFormula, 1, i)
                    sFormula = Trim(Mid$(sFormula, i + 1))
                    Exit For
                End If
            Case Else
                If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
                    If iWait = 0 Then
                        LeftVal = Trim(Mid$(sFormula, 1, i - 1))
                        sFormula = Trim(Mid$(sFormula, i))
                        Exit For
                    End If
                End If
        End Select
    Next i
    
    ' Get the comparison operator
    If LenB(sFormula) > 0 Then
    sTemp = Mid$(sFormula, 2, 1)
        If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
            Operator = Left(sFormula, 2)
            sFormula = Mid$(sFormula, 3)
        Else
            Operator = Left(sFormula, 1)
            sFormula = Mid$(sFormula, 2)
        End If
    End If
    
    ' Set the rest to the right side operator
    RightVal = sFormula
    sFormula = ""
    
    If Left(LeftVal, 1) = "!" Then
        blLeftNot = True
        LeftVal = Right(LeftVal, Len(LeftVal) - 1)
    End If
    If Left(RightVal, 1) = "!" Then
        blRightNot = True
        RightVal = Right(RightVal, Len(RightVal) - 1)
    End If
    
    ' Solve these sides
    'LeftVal = Trim$(Equation(LeftVal))
    If NumberOfFunctions(LeftVal) > -1 Then
        LeftVal = DoFunction(Trim$(LeftVal))
    End If
    
    LeftVal = Trim$(Equation(LeftVal))

    If LenB(RightVal) > 0 Then
        If NumberOfFunctions(RightVal) > -1 Then
            RightVal = DoFunction(Trim$(RightVal))
        End If
        
        RightVal = Trim$(Equation(RightVal))

    Else
        If LenB(Operator) = 0 Then
            Operator = "="
            RightVal = "True"
        End If
    End If
    
    
    sTemp = TrimChar(LeftVal, Chr(34), 0)
    If blLeftNot Then
        blEval = sTemp
        LeftVal = Not blEval
    End If
    
    sTemp = TrimChar(RightVal, Chr(34), 0)
    If blRightNot Then
        blEval = sTemp
        RightVal = Not blEval
    End If
        
    If LenB(Operator) > 0 Then
        Select Case Operator
            Case ">"
                If Val(LeftVal) > Val(RightVal) Then Eval = True
            Case "<"
                If Val(LeftVal) < Val(RightVal) Then Eval = True
            Case "<>"
                If Val(LeftVal) <> Val(RightVal) Then Eval = True
            Case ">="
                If Val(LeftVal) >= Val(RightVal) Then Eval = True
            Case "<="
                If Val(LeftVal) <= Val(RightVal) Then Eval = True
            Case "="
                If LeftVal = RightVal Then Eval = True
            Case "=="
                If LCase(LeftVal) = LCase(RightVal) Then Eval = True
        End Select
    End If
    
    If LenB(Eval) = 0 Then Exit Function
    PrintDebug "Eval Solved: " & Eval, 3
End Function

Private Function Equation(ByVal sEquation As String) As Variant
    ' For solving full equations like: iTemp + ("3" * ("5" / "4"))
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String
    Dim hasCalc As Boolean
    ' First lets change anything inside ()'s to a constant value
    If Len(sEquation) <= 1 Then Equation = sEquation: Exit Function
    Do
        For i = 1 To Len(sEquation)
            sTemp = Mid$(sEquation, i, 1)
            If sTemp = Chr(34) Then
                If Not (i + 1) > Len(sEquation) Then i = InStr(i + 1, sEquation, Chr(34))
            ElseIf sTemp = "(" Then
                iTemp = i
            ElseIf sTemp = ")" Then
                iTemp2 = i
                Exit For
            'ElseIf sTemp = "+" Or sTemp = "-" Or sTemp = "*" Or sTemp = "/" Or sTemp = "\" Or sTemp = "^" Then
            '    hasCalc = True
            End If
            DoEvents
        Next i
        If iTemp <> 0 Then
            sEquation = Mid$(sEquation, 1, iTemp - 1) & " " & Chr(34) & Solve(Mid$(sEquation, iTemp + 1, iTemp2 - (iTemp + 1))) & Chr(34) & " " & Mid$(sEquation, iTemp2 + 1)
            iTemp = 0: iTemp2 = 0
        Else
            Exit Do
        End If
    Loop
    
    ' Now that all the ()'s are constant values, lets finish it up
    'If hasCalc Then Equation = Solve(sEquation) Else Equation = sEquation
    Equation = Solve(sEquation)
    If LenB(Equation) = 0 Then Equation = sEquation: Exit Function
    If Equation = sEquation Then Exit Function
    PrintDebug "Equation: " & Equation, 3
End Function

Private Function Solve(sEquation As String) As Variant
    ' For solving equations without ()'s like: itemp + "3" / "2"
    Dim i As Integer, iTemp As Integer
    Dim sTemp As String, sTemp2 As String
    Dim WaitVal As Variant
    Dim WaitOp As String
    Dim WaitVar

    For i = 1 To Len(sEquation)
        sTemp = Mid$(sEquation, i, 1)
        If sTemp = Chr(34) Then
            If Mid$(sEquation, i, 3) = Chr(34) & Chr(34) & Chr(34) Then
                i = i + 3
                WaitVal = Chr(34)
            Else
                iTemp = InStr(i + 1, sEquation, Chr(34))
                WaitVal = Mid$(sEquation, i + 1, iTemp - (i + 1))
                i = iTemp
            End If
            If WaitOp <> "" Then
                Select Case WaitOp
                    Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                    Case "-"
                        Solve = Val(Solve) - Val(WaitVal)
                    Case "/"
                        Solve = Val(Solve) / Val(WaitVal)
                    Case "\"
                        Solve = Val(Solve) \ Val(WaitVal)
                    Case "^"
                        Solve = Val(Solve) ^ Val(WaitVal)
                    Case "*"
                        Solve = Val(Solve) * Val(WaitVal)
                    Case "&" ' String Concentration
                        Solve = Solve & WaitVal
                End Select
                WaitOp = ""
            Else
                Solve = WaitVal
            End If
        ElseIf sTemp = " " Then
            ' Make sure it skips this
        ElseIf sTemp = "+" Or sTemp = "-" Or sTemp = "/" Or sTemp = "\" Or sTemp = "^" Or sTemp = "&" Or sTemp = "*" Then
            If WaitVar <> 0 Then
                WaitVal = GetVar(Mid$(sEquation, WaitVar, i - (WaitVar + 1)))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
                
                WaitVar = 0
            End If
            
            WaitOp = sTemp
        Else
            If WaitVar = 0 Then WaitVar = i
            If i >= Len(sEquation) Then
                WaitVal = GetVar(Mid$(sEquation, WaitVar, i))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                            Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
            End If
        End If
    Next i
    
    If LenB(Solve) = 0 Then Exit Function
    If Solve = sEquation Then Exit Function
    PrintDebug "Solve: " & Solve, 3
    
End Function

Public Function DoFunction(sCall As String, Optional CallingVar As String = "") As Variant

On Error GoTo err

    'Test string:   debug(~reverse(~rnd("10")))
    'and:           debug("(" & ~reverse(~rnd("10") & ")")))
    'or :           debug("(" & ~reverse(~rnd("10")) & ")"))
    
    Dim ArgList() As String
    'Dim CmdList() As String
    Dim strTemp As String
    Dim lngTemp As String
    Dim sName As String
    Dim i As Long
    
    sCall = Trim$(sCall)
    
    If LenB(sCall) = 0 Then Exit Function

    If Not Left(sCall, 1) = "~" Then sCall = "~" & sCall
    lngTemp = InStr(1, sCall, "(")
    'strFunction = Left(sCall, lngTemp - 1)

    sName = Trim$(Mid$(sCall, 1, InStr(1, sCall, "(") - 1))
    'sCall = Mid(sCall, Len(sName) + 2, Len(sCall) - Len(sName) - 2)
    sCall = Replace(sCall, sName, "", 1, 1)
    sCall = Trim$(sCall)
    
    PrintDebug "DoFunction: " & sName & " sCall: " & sCall, 1
    
    If NumberOfFunctions(sCall) > -1 Then
        strTemp = ProcessFunctions(sCall)
    Else
        strTemp = sCall
    End If
    
    If LenB(strTemp) > 0 Then
        If InStr(1, strTemp, "(") = 0 Then strTemp = "(" & strTemp & ")"
        ArgList = Split(Mid$(strTemp, InStr(1, strTemp, "(") + 1, InStrRev(strTemp, ")") - (InStr(1, strTemp, "(") + 1)), ",")  'split each group by the comma
        
        For i = LBound(ArgList) To UBound(ArgList) 'loop through each comma-separated group
            strTemp = ArgList(i)
            ArgList(i) = Equation(Trim(ArgList(i)))
            If strTemp = ArgList(i) Then ArgList(i) = GetVar(strTemp)
        Next i
    End If
    'return string
    If UBound(ArgList) = 0 And VarExists(strTemp) Then CallingVar = strTemp
        strTemp = ExecFunction(sName, ArgList, CallingVar)
    
    DoFunction = strTemp
    If LenB(strTemp) = 0 Then Exit Function
    PrintDebug "DoFunction solved: " & DoFunction, 3
    
Exit Function

err:
    Call NotifyError(err.Number & ": " & err.Description, nCurrentline)
End Function

Private Function ProcessFunctions(strString As String) As Variant
    Dim ArgList() As String
    Dim arrFunctions() As String
    Dim i As Long, ii As Long
    Dim strTemp As String, sFunc As String, sName As String
    Dim lngFuncs As Long
    Dim lngNext As Long
    lngFuncs = NumberOfFunctions(strString)
    lngNext = 1
    
    PrintDebug "ProcessFunctions: " & strString, 2
    
    If lngFuncs = -1 Then
        If Right(strString, 2) = "))" Then strString = Left(strString, Len(strString) - 1)
        ProcessFunctions = strString
        Exit Function
    End If
    If lngFuncs = 1 Then
        'If Right(strString, 2) = "))" Then strString = Left(strString, Len(strString) - 1)
        arrFunctions = GetNextFunction(strString) 'This also process functions within functions
        strString = Replace(strString, arrFunctions(0), arrFunctions(1))
        'lngNext = InStr(1, strString, arrFunctions(1)) + Len(arrFunctions(1))
        PrintDebug "ProcessFunctions: Solved Function [ " & arrFunctions(0) & " ] for: [ " & arrFunctions(1) & " ]", 3
        ProcessFunctions = strString
        Exit Function
    End If
    
    If lngFuncs > 0 Then lngFuncs = lngFuncs - 1
    ReDim arrFunctions(lngFuncs)
        
        
    If UBound(arrFunctions) >= 0 Then
    
        arrFunctions = GetNextFunction(strString) 'This also process functions within functions
        strString = Replace(strString, arrFunctions(0), arrFunctions(1), 1, 1)
        lngNext = InStr(1, strString, arrFunctions(1)) + Len(arrFunctions(1))
        PrintDebug "ProcessFunctions: Solved Function [ " & arrFunctions(0) & " ] for: [ " & arrFunctions(1) & " ]", 3
        
        For i = 1 To lngFuncs
            'get and do the functions
            'strTemp = GetNextFunction(strString) 'This also process functions within functions
            'ArgList = Split(Mid$(strTemp, InStr(1, strTemp, "(") + 1, InStrRev(strTemp, ")") - (InStr(1, strTemp, "(") + 1)), ",")  'split each group by the comma
            'sName = Trim$(Mid$(strTemp, 1, InStr(1, strTemp, "(") - 1))
            '
            'For ii = LBound(ArgList) To UBound(ArgList) 'loop through each comma-separated group
            '    strTemp = ArgList(ii)
            '    ArgList(ii) = Equation(Trim(ArgList(i)))
            '    If strTemp = ArgList(i) Then ArgList(i) = GetVar(strTemp)
            'Next ii
            '
            'ProcessFunctions = ExecFunction(sName, ArgList)
            
            arrFunctions = GetNextFunction(Mid(strString, lngNext, Len(strString) - lngNext + 1)) 'This also process functions within functions
            strString = Replace(strString, arrFunctions(0), arrFunctions(1), 1, 1)
            lngNext = InStr(1, strString, arrFunctions(1))
            PrintDebug "ProcessFunctions: Solved Function [ " & arrFunctions(0) & " ] for: [ " & arrFunctions(1) & " ]", 3
        Next

    End If
    
    ProcessFunctions = strString
    
End Function

Private Function GetNextFunction(strString As String) As Variant
    Dim lngLeftBracketCount As Long
    Dim lngRightBracketCount As Long
    Dim blIsQuoteOpen As Boolean
    Dim blIsFunctionOpen As Boolean
    Dim i As Long
    Dim strReturn(1) As String
    Dim lngFuncStart As Long
    Dim blNot As Boolean
    Dim sTemp As String
    For i = 1 To Len(strString)
    
        Select Case Mid(strString, i, 1)
            Case Is = Chr(34)
                blIsQuoteOpen = Not blIsQuoteOpen
        
            Case Is = "("
                If Not blIsQuoteOpen And blIsFunctionOpen Then
                    lngLeftBracketCount = lngLeftBracketCount + 1
                End If
            Case Is = ")"
                If Not blIsQuoteOpen And blIsFunctionOpen Then
                    lngRightBracketCount = lngRightBracketCount + 1
                    If lngRightBracketCount = lngLeftBracketCount Then
                        'function has ended
                        strReturn(1) = DoFunction(Mid(strString, lngFuncStart, i - lngFuncStart + 1))
                        'strReturn(0) = Mid(strString, lngFuncStart, i - lngFuncStart + 1)
                        
                        If lngFuncStart > 1 Then
                            If Mid(strString, lngFuncStart - 1, 1) = "!" Then
                                strReturn(0) = Mid(strString, lngFuncStart - 1, i - lngFuncStart + 2)
                                sTemp = TrimChar(strReturn(1), Chr(34), 1)
                                If sTemp = "True" Or sTemp = "False" Then
                                    blNot = strReturn(1)
                                    strReturn(1) = Chr(34) & (Not blNot) & Chr(34)
                                End If
                            Else
                                strReturn(0) = Mid(strString, lngFuncStart, i - lngFuncStart + 1)
                            End If
                        Else
                            strReturn(0) = Mid(strString, lngFuncStart, i - lngFuncStart + 1)
                        End If
                        
                        GetNextFunction = strReturn()
                        Exit Function
                    End If
                End If

            Case Is = "~"
            If Not blIsFunctionOpen Then
                lngFuncStart = i
                blIsFunctionOpen = True
                lngRightBracketCount = 0
                lngLeftBracketCount = 0
            End If
        End Select
    Next i

End Function

Private Function GetBracketContents(strString As String) As String
    Dim lngLeftBracketCount As Long
    Dim lngRightBracketCount As Long
    Dim lngFirstPos As Long
    Dim lngLastPos As Long
    Dim blIsQuoteOpen As Boolean
    Dim i As Long
    'extract the text between two brackets in a string and return its contents
        
        For i = 1 To Len(strString)
            
            Select Case Mid(strString, i, 1)
            
                Case Is = Chr(34)
                    blIsQuoteOpen = Not blIsQuoteOpen
            
                Case Is = "("
                    If Not blIsQuoteOpen Then
                        lngLeftBracketCount = lngLeftBracketCount + 1
                        If lngFirstPos = 0 Then lngFirstPos = i
                    End If
                    
                Case Is = ")"
                    If Not blIsQuoteOpen Then
                        lngRightBracketCount = lngRightBracketCount + 1
                        lngLastPos = i
                    End If
                    
                Case Is = "~"
                    If Not blIsQuoteOpen Then
                        'there's a function!
                    End If
                    
            End Select
        
            If (lngRightBracketCount > 0) And (lngLeftBracketCount > 0) Then
                
                If lngLeftBracketCount = lngRightBracketCount Then
                    'We've found the close of this bracket.
                    'let's extract the data and return it.
                    
                    GetBracketContents = Mid$(strString, lngFirstPos + 1, lngLastPos - lngFirstPos - 1)
                    
                End If
                
            End If
        
        Next i
        
End Function

Private Function GetReturnArgs(strString As String) As Variant
    Dim lngLastPos As Long
    Dim blIsQuoteOpen As Boolean
    Dim strReturn() As String
    Dim i As Long
    ReDim strReturn(0) As String
    'extract the text between two brackets in a string and return its contents
        
        strString = Mid(strString, Len("return ") + 1)
        lngLastPos = 1
        
        For i = 1 To Len(strString)
            
            Select Case Mid(strString, i, 1)
            
                Case Is = Chr(34)
                    blIsQuoteOpen = Not blIsQuoteOpen
            
                Case Is = ","
                    If Not blIsQuoteOpen Then
                        If LenB(strReturn(UBound(strReturn))) > 0 Then ReDim Preserve strReturn(UBound(strReturn) + 1) As String
                        strReturn(UBound(strReturn)) = Mid(strString, lngLastPos, i - 1)
                        lngLastPos = i + 1 'account for the comma
                    End If
                    
            End Select
        
        Next i
        
        If i > lngLastPos Then
            If LenB(strReturn(UBound(strReturn))) > 0 Then ReDim Preserve strReturn(UBound(strReturn) + 1) As String
            strReturn(UBound(strReturn)) = Trim$(Mid(strString, lngLastPos, i - 1))
        End If
        
        GetReturnArgs = strReturn
        
End Function

Private Function NumberOfFunctions(strString As String) As Long
    Dim lngLeftBracketCount As Long
    Dim lngRightBracketCount As Long
    Dim blIsQuoteOpen As Boolean
    Dim blIsFunctionOpen As Boolean
    Dim i As Long
    
    NumberOfFunctions = -1
    
    For i = 1 To Len(strString)
    
        Select Case Mid(strString, i, 1)
            Case Is = Chr(34)
                blIsQuoteOpen = Not blIsQuoteOpen
        
            Case Is = "("
                If Not blIsQuoteOpen And blIsFunctionOpen Then
                    lngLeftBracketCount = lngLeftBracketCount + 1
                End If
            Case Is = ")"
                If Not blIsQuoteOpen And blIsFunctionOpen Then
                    lngRightBracketCount = lngRightBracketCount + 1
                    If lngRightBracketCount = lngLeftBracketCount Then
                        'function has ended
                        If NumberOfFunctions = -1 Then NumberOfFunctions = 0
                        NumberOfFunctions = NumberOfFunctions + 1
                        blIsFunctionOpen = False
                    End If
                End If

            Case Is = "~"
            If Not blIsFunctionOpen Then
                blIsFunctionOpen = True
                lngRightBracketCount = 0
                lngLeftBracketCount = 0
            End If
        End Select
    Next i


'PrintDebug "NumberOfFunctions: " & NumberOfFunctions & " found in [ " & strString & " ]"

End Function

Private Function IsCalc(char As String) As Boolean
    Dim sChr As Integer
    
    sChr = Asc(char)
    
    Select Case sChr
    
    Case 42, 43, 45, 47, 92, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 40, 41, 32 '*+-/0-9() and " "
        IsCalc = True
    
    End Select
        
End Function

Private Sub Pause(TimeToWait As Long)
Dim i As Long
    For i = 1 To TimeToWait Step 10
        Sleep 10: DoEvents
    Next i
End Sub

Public Sub CleanupVars()
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        VarNames.Remove i
        Variables.Remove i
    Next i
    For i = FileNames.Count To 1 Step -1
        FileNames.Remove i
        FileContents.Remove i
    Next i
    IfSkip = False
    ScrExeSkip = False
    ScrExeiTemp = 0
    ScrExeiTemp2 = 0
    ScrExei = 0
    ScrExesTemp = ""
    ScrExesTemp2 = ""
    ReDim ScrExearTemp2(0)
    Set InsideLoop = New Collection
    ReDim ReturnVar(0)
    'PrintDebug "CleanupVars complete. Everything but Debug ;-)", 1
End Sub

Public Sub ExitScript()
On Error Resume Next
    CleanupVars
    StopExecution = True
    nCurrentline = UBound(SplitScript)
    
End Sub

'String to trim, character to trim, number of characters to trim
Private Function TrimChar(StringToTrim As String, sChar As String, Optional iNum As Integer = 1) As String
Dim lPos As Long
Dim lRemoved As Long

StringToTrim = Trim$(StringToTrim)

Do While Left(StringToTrim, 1) = sChar
    StringToTrim = Right(StringToTrim, Len(StringToTrim) - 1): lRemoved = lRemoved + 1
    If iNum > 0 And lRemoved = iNum Then Exit Do
Loop
lRemoved = 0

Do While Right(StringToTrim, 1) = sChar
    StringToTrim = Left(StringToTrim, Len(StringToTrim) - 1): lRemoved = lRemoved + 1
    If iNum > 0 And lRemoved = iNum Then Exit Do
Loop
'replace the first occurrence of sChar in front and reverse before encountering any chars other than a space (but trim that, too)

TrimChar = StringToTrim

End Function

Private Function GotoScriptLine(Txt As String)
    Dim i As Long

    For i = LBound(SplitScript) To UBound(SplitScript)
        If Left(LCase$(Trim$(SplitScript(i))), Len(Txt)) = LCase$(Txt) Then
            nCurrentline = i
            PrintDebug "GotoScriptLine: found [" & Txt & "] at line " & i + 1, 2
            Exit Function
        End If
    Next
End Function

Private Function IsPathAllowed(Path As String) As Boolean

    'this should be made into its own function and called every time
    'we need to get an EXEPath.. It's used in a few places.
    If Left(Path, 1) = "\" Then
        Path = FullPathFromLocal(Path)
    End If

    If InStr(1, LCase$(Path), LCase$(App.Path)) > 0 Then IsPathAllowed = True

    PrintDebug "IsPathAllowed: [ " & Path & " ] : [ " & IsPathAllowed & " ]", 1

End Function

Private Sub NotifyError(Msg As String, Line As Long)

        PrintDebug ("Error on Line " & Line & ": " & Msg)
        If blShowErrors Then RaiseEvent ScriptError(Msg & vbNewLine, Line)
    
End Sub

Private Sub PrintDebug(Msg As String, Optional Level As Integer = 0)
    If Not blDebug Then Exit Sub
    
    If Level <= DebugLevel Then
        If Not myDebug Is Nothing Then
            Call myDebug.PrintDebug(Msg)
        Else
            RaiseEvent DebugOut(Msg)
        End If
    End If
    
End Sub

Public Sub ShowDebug()
    If Not blDebug Then blDebug = True
    If myDebug Is Nothing Then Set myDebug = New frmScriptDebug
    If Not myParent Is Nothing Then Set myDebug.myForm = myParent: myDebug.tmrSnap.Enabled = True
    Set myDebug.theScript = Me
    myDebug.Visible = True
End Sub

Private Sub Class_Initialize()
     Set InsideLoop = New Collection
     ReDim ReturnVar(0)
End Sub

Public Sub GetAllVars(ByRef Names() As String, ByRef VAR() As String)
On Error Resume Next
    Dim i As Long
    Dim strTemp As String
    If VarNames.Count = 0 Then Exit Sub
    
    ReDim Names(VarNames.Count) As String
    ReDim VAR(VarNames.Count) As String
        
    For i = VarNames.Count To 1 Step -1
        Names(i) = VarNames(i)
        VAR(i) = Variables(i)
    Next i

End Sub

Public Sub GetAllFileVars(ByRef Names() As String, ByRef VAR() As String)
On Error Resume Next
    Dim i As Long
    Dim strTemp As String
    If VarNames.Count = 0 Then Exit Sub
    
    ReDim Names(FileNames.Count) As String
    ReDim VAR(FileContents.Count) As String
        
    For i = FileNames.Count To 1 Step -1
        Names(i) = VarNames(i)
        VAR(i) = Variables(i)
    Next i

End Sub

Private Function LoadFile(dFile As String) As String

    Dim ff As Integer

    On Error Resume Next

    ff = FreeFile
    Open dFile For Binary As #ff
        LoadFile = Space(LOF(ff))
        Get #ff, , LoadFile
    Close #ff

End Function

Private Sub WriteFile(strOutput As String, Optional strFile As String, Optional Overwrite As Boolean = False)

    Dim ff As Integer

    On Error Resume Next

    ff = FreeFile
    If Not Overwrite Then
        Open strFile For Binary Access Read Write As #ff
        Seek #ff, LOF(ff)
        Put #ff, , strOutput
    Else
        Open strFile For Output As #ff
        Print #ff, strOutput
    End If
        
        
    Close #ff

End Sub

Private Function WriteIni(INIpath As String, PutKey As String, PutVariable As String, PutValue As String)
Dim Temp As String
Dim LcaseTemp As String
Dim ReadKey As String
Dim ReadVariable As String
Dim LOKEY As Integer
Dim HIKEY As Integer
Dim KeyLen As Integer
Dim VAR As Integer
Dim VARENDOFLINE As Integer
Dim NF As Integer
Dim X As Integer

AssignVariables:
    NF = FreeFile
    ReadKey = vbCrLf & "[" & LCase$(PutKey) & "]" & Chr$(13)
    KeyLen = Len(ReadKey)
    ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
        
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    Temp = Input$(LOF(NF), NF)
    Temp = vbCrLf & Temp & "[]"
    Close NF
    LcaseTemp = LCase$(Temp)
    
LogicMenu:
    LOKEY = InStr(LcaseTemp, ReadKey)
    If LOKEY = 0 Then GoTo AddKey:
    HIKEY = InStr(LOKEY + KeyLen, LcaseTemp, "[")
    VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
    If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
    GoTo RenewVariable:
    
AddKey:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Temp & vbCrLf & vbCrLf & "[" & PutKey & "]" & vbCrLf & PutVariable & "=" & PutValue
        GoTo TrimFinalString:
        
AddVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Left$(Temp, LOKEY + KeyLen) & PutVariable & "=" & PutValue & vbCrLf & Mid$(Temp, LOKEY + KeyLen + 1)
        GoTo TrimFinalString:
        
RenewVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        VARENDOFLINE = InStr(VAR, Temp, Chr$(13))
        Temp = Left$(Temp, VAR) & PutVariable & "=" & PutValue & Mid$(Temp, VARENDOFLINE)
        GoTo TrimFinalString:

TrimFinalString:
        Temp = Mid$(Temp, 2)
        Do Until InStr(Temp, vbCrLf & vbCrLf & vbCrLf) = 0
        Temp = Replace(Temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        Loop
    
        Do Until Right$(Temp, 1) > Chr$(13)
        Temp = Left$(Temp, Len(Temp) - 1)
        Loop
    
        Do Until Left$(Temp, 1) > Chr$(13)
        Temp = Mid$(Temp, 2)
        Loop
    
OutputAmendedINIFile:
        Open INIpath For Output As NF
        Print #NF, Temp
        Close NF
    
End Function

Private Function ReadIni(INIpath As String, Key As String, Variable As String) As String
Dim NF As Integer
Dim Temp As String
Dim LcaseTemp As String
Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIni = ""
        Key = "[" & LCase$(Key) & "]"
        Variable = LCase$(Variable)
    
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    While Not EOF(NF)
    Line Input #NF, Temp
    LcaseTemp = LCase$(Temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = Key Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIni = Mid$(Temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
End Function

Private Function FileExists(sFileName As String) As Boolean
    If Len(sFileName) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir(sFileName)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Private Sub InsertData(ByVal FileNum As Integer, DataToInsert As String, ByVal InsertPos As Long)
    '
    '  Made by Michael Ciurescu (CVMichael from vbforums.com)
    '  Original thread: [url]http://www.vbforums.com/showthread.php?t=433537[/url]
    '
 
    Const cBuffSize As Long = 262144 ' 256 KBytes
    
    Dim Buffer() As Byte, BuffPos As Long
    
    ' Shift all data to the right
    
    If LOF(FileNum) - InsertPos < cBuffSize Then
        ' we can do it in one copy, we don't need a loop
        
        ' resize the buffer so we don't copy too much (more than file size)
        ReDim Buffer(LOF(FileNum) - InsertPos - 1)
        
        ' copy and paste the data to the new location
        Get FileNum, InsertPos, Buffer
        Put FileNum, InsertPos + Len(DataToInsert), Buffer
    Else
        ' we start from the end of the file
        BuffPos = LOF(FileNum)
        
        ' we HAVE to shift from right to left,
        ' otherwise we override important data
        
        Do Until BuffPos <= InsertPos
            ' substract the buffer size from current position
            BuffPos = BuffPos - cBuffSize
            
            If BuffPos < InsertPos Then
                ' we reached the last copy
                ' data passed the InsertPos position, so we have to
                ' resize the buffer so it does not go over InsertPos
                
                ReDim Buffer(cBuffSize - (InsertPos - BuffPos) - 1)
                BuffPos = InsertPos
            Else
                ReDim Buffer(cBuffSize - 1)
            End If
            
            ' copy and paste the data to the new location
            Get FileNum, BuffPos, Buffer
            Put FileNum, BuffPos + Len(DataToInsert), Buffer
        Loop
    End If
    
    ' Insert the actual data
    Put FileNum, InsertPos, DataToInsert
End Sub
 
Private Sub DeleteData(ByVal FileName As String, DeletePos As Long, DeleteLength As Long)
    '
    '  Made by Michael Ciurescu (CVMichael from vbforums.com)
    '  Original thread: [url]http://www.vbforums.com/showthread.php?t=433537[/url]
    '
    Const cBuffSize As Long = 262144 ' 256 KBytes
    
    Dim SA As SECURITY_ATTRIBUTES
    Dim FHandle As Long
    Dim FileLen As Double
    Dim Buffer() As Byte, BuffPtr As Long
    
    Dim BytesToRead As Long, BytesRead As Long
    Dim ReadPos As Double, WritePos As Double
    
    ' using API position 0 is the first byte, using VB functions position 1 is first byte
    ' so decrement by one to use the same standard...
    DeletePos = DeletePos - 1
    
    ' open the file
    FHandle = CreateFile(FileName, GENERIC_READ_WRITE, 0, SA, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    ' get file size
    FileLen = FileSizeDouble(FHandle)
    
    ' alocate memory
    ReDim Buffer(cBuffSize - 1)
    
    ' get memory pointer
    BuffPtr = VarPtr(Buffer(0))
    
    ' calculate read & write positions
    WritePos = DeletePos
    ReadPos = WritePos + DeleteLength
    
    ' shift the data to left
    
    Do Until ReadPos >= FileLen
        ' calculate how much data to read/write
        BytesToRead = dblMIN(cBuffSize, FileLen - ReadPos)
        
        ' copy and paste the data to the new location
        SeekPosDouble FHandle, ReadPos
        ReadFileA FHandle, ByVal BuffPtr, BytesToRead, BytesRead, ByVal 0&
        
        SeekPosDouble FHandle, WritePos
        WriteFileA FHandle, ByVal BuffPtr, BytesRead, BytesRead, ByVal 0&
        
        WritePos = WritePos + BytesRead
        ReadPos = WritePos + DeleteLength
    Loop
    
    If WritePos < FileLen Then
        ' Seek to where we need to truncate the file
        SeekPosDouble FHandle, WritePos
        
        ' truncate the file
        SetEndOfFile FHandle
    End If
    
    Erase Buffer
    If FHandle <> 0 Then CloseHandle FHandle
End Sub
 
Private Function dblMIN(ByVal V1 As Double, ByVal V2 As Double) As Double
    If V1 < V2 Then
        dblMIN = V1
    Else
        dblMIN = V2
    End If
End Function
 
Private Function SeekPosDouble(ByVal FHandle As Long, ByVal NewPos As Double) As Boolean
    Dim SizeLow As Long, SizeHigh As Long
    
    SizeLow = DoubleToLongs(NewPos, SizeHigh)
    
    SeekPosDouble = SeekPos(FHandle, SizeLow, SizeHigh)
End Function
 
Private Function SeekPos(ByVal FHandle As Long, ByVal NewPos As Long, _
                                   Optional ByVal PosHigh As Long = 0) As Boolean
    Dim Ret As Long, dwError As Long
    
    Ret = SetFilePointer(FHandle, NewPos, PosHigh, FILE_BEGIN)
    
    If Ret = -1 Then
        dwError = GetLastError
        If dwError = NO_ERROR Then SeekPos = True
    Else
        SeekPos = True
    End If
End Function
 
Private Function FileSizeDouble(ByVal FHandle As Long) As Double
    Dim SizeLow As Long, SizeHigh As Long
    
    If FHandle <> 0 Then SizeLow = GetFileSize(FHandle, SizeHigh)
    
    FileSizeDouble = CDbl(SizeHigh) * (2 ^ 32) + LongToDouble(SizeLow)
End Function
 
Private Function LongToDouble(ByVal Lng As Long) As Double
    If Lng And &H80000000 = 0 Then
        LongToDouble = CDbl(Lng)
    Else
        LongToDouble = (Lng Xor &H80000000) + (2 ^ 31)
    End If
End Function
 
Private Function DoubleToLongs(ByVal Dbl As Double, ByRef SizeHigh As Long) As Long
    Dim SizeLowDbl As Double
    
    SizeHigh = Fix(Dbl / 4294967296#)
    SizeLowDbl = Dbl - SizeHigh * 4294967296#
    
    If SizeLowDbl > 2147483647 Then
        DoubleToLongs = CLng(SizeLowDbl - 2147483648#) Xor &H80000000
    Else
        DoubleToLongs = SizeLowDbl
    End If
End Function


