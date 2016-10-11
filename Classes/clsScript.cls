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

Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)



Public Script As String
Private SplitScript() As String
Private VarNames As New Collection
Private Variables As New Collection
Private nCurrentLine As Long

Public Event LineChange(nLine As Long, sContents As String)
Public Event ScriptError(Msg As String, Line As Long)
Public Event DebugOut(Msg As String)
Public Event CommandOut(Command As String, ArgList() As String)

Private IfSkip As Boolean
Private ScrExeSkip As Boolean
Private ScrExeiTemp As Integer, ScrExeiTemp2 As Integer, ScrExei As Integer
Private ScrExesTemp As String, ScrExesTemp2 As String
Private ScrExearTemp2() As String
Private InsideLoop As Collection
Private blDebug As Boolean
Private blShowErrors As Boolean

Private myDebug As Form
'Public myForm As Form
'Public theForm As Form

Public Sub ScriptExecute(Optional blOneLine As Boolean = False)
    Dim sCurrentLine As String
    Dim sTemp As String
    Dim lngTemp As Long
    
    If LenB(Script) = 0 Then Call NotifyError("No script has been loaded!", 0): Exit Sub
    
    SplitScript = Split(Script, vbCrLf)
    
    If Not blOneLine Then nCurrentLine = LBound(SplitScript)
    
    For nCurrentLine = nCurrentLine To UBound(SplitScript)
        sCurrentLine = TrimChar$(SplitScript(nCurrentLine), Chr(9), 0)
        RaiseEvent LineChange(nCurrentLine, sCurrentLine)
        
        ' Get rid of comments
        ScrExeiTemp = InStrRev(sCurrentLine, "#")
        If ScrExeiTemp >= 1 Then
            lngTemp = InStrRev(sCurrentLine, Chr(34))
            If (lngTemp = 0) Or (lngTemp < ScrExeiTemp) Or ScrExeiTemp = 1 Then ' And (InStrRev(sCurrentLine, Chr(34), Len(sCurrentLine)) < ScrExeiTemp) Then
                sCurrentLine = TrimChar(Left(sCurrentLine, ScrExeiTemp - 1), Chr(9), 0) 'Right(sCurrentLine, Len(sCurrentLine) - ScrExeiTemp) 'Mid$(sCurrentLine, 1, ScrExeiTemp - 1)
            End If
        End If
        
        If LenB(sCurrentLine) = 0 Then GoTo SkipCurrentDo:
        
        ' Check for a label
        'If StopLabel Then
        '    StopLabel = False
        '    GoTo SkipCurrentDo:
        'End If
        
        ' Check for end of a block
        If ScrExeSkip = True Then
            If LCase(sCurrentLine) = "end" Then
                ScrExeSkip = False
            End If
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
            End If
        End If
        
        ' Looping block
        If LCase(sCurrentLine) = "end" And InsideLoop.Count > 0 Then
            If Eval(Mid$(InsideLoop(InsideLoop.Count), InStr(1, InsideLoop(InsideLoop.Count), ":") + 1)) Then
                nCurrentLine = CInt(Mid$(InsideLoop(InsideLoop.Count), 1, InStr(1, InsideLoop(InsideLoop.Count), ":") - 1)) - 1
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
                Case "goto" ' Goto a label
                    'StopLabel = True
                    GotoScriptLine Trim(Mid$(sCurrentLine, ScrExeiTemp + 1))
                    GoTo SkipCurrentDo:
                    
                Case "var", "dim" ' Create variable
                    ScrExearTemp2 = Split(Trim(Mid$(sCurrentLine, ScrExeiTemp + 1)), ",")
                    For ScrExei = LBound(ScrExearTemp2) To UBound(ScrExearTemp2)
                        If VarExists(ScrExearTemp2(ScrExei)) Then
                            Call NotifyError("Variable has already been defined: " & ScrExearTemp2(ScrExei), nCurrentLine)
                        Else
                            AddVar ScrExearTemp2(ScrExei), ""
                        End If
                        Next ScrExei
                        
                    GoTo SkipCurrentDo:
                
                Case "revar", "redim"
                    If Mid(LCase$(sCurrentLine), 5, 8) = "preserve" Then
                        'preserve the array, only add new ones or remove old ones outside of the bounds
                        sTemp = DoFunction(Right(sCurrentLine, 13))
                        
                    Else
                        
                    End If
                    
                    'TODO: allow preserve subcommand. Clear the var. Check if the var is an array.
                    
                Case "while" ' Loop
                    If Eval(Mid$(sCurrentLine, ScrExeiTemp + 1)) Then
                        InsideLoop.Add nCurrentLine + 1 & ":" & Trim(Mid$(sCurrentLine, ScrExeiTemp + 1))
                    Else
                        ScrExeSkip = True
                    End If
                    GoTo SkipCurrentDo:
            End Select
        End If
        
        For ScrExei = 1 To Len(sCurrentLine)
            Select Case Mid$(sCurrentLine, ScrExei, 1)
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
                    ChangeVar Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "=") - 1)), Equation(Mid$(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                    GoTo SkipCurrentDo:
                Case "~" ' Assignment from Function
                    ChangeVar Trim(Mid$(sCurrentLine, 1, InStr(1, sCurrentLine, "~") - 1)), DoFunction(Mid$(sCurrentLine, InStr(1, sCurrentLine, "~") + 1))
                    GoTo SkipCurrentDo:
                Case "(" ' Procedure Call
                    DoFunction sCurrentLine
                    GoTo SkipCurrentDo:
            End Select
        Next ScrExei
        
        Call NotifyError("Line not processed: " & sCurrentLine, (ScrExei + 1))
        
SkipCurrentDo:
    If blOneLine Then nCurrentLine = nCurrentLine + 1: Exit Sub
    Next nCurrentLine
    
    ' Its the clean up crew!

End Sub
    
Private Function ExecFunction(sFunction As String, ArgList() As String) As Variant
    On Error GoTo ErrorCatch:
    Dim arrTemp() As String
    Dim i As Long
    If Left(sFunction, 1) = "~" Then sFunction = Right(sFunction, Len(sFunction) - 1)
    
        PrintDebug "Function: " & ExecFunction & "(" & Join(ArgList, ", ")
    
    Select Case LCase$(sFunction)
        Case "msgbox"
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
            Select Case UBound(ArgList)
                Case Is = 0
                    ExecFunction = InputBox(ArgList(0))
                Case Is = 1
                    ExecFunction = InputBox(ArgList(0), ArgList(1))
                Case Is = 2
                    ExecFunction = InputBox(ArgList(0), ArgList(1), ArgList(2))
            End Select
            
        Case "app"
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
            
        'Case "myform"
        '
        '    Select Case ArgList(0)
        '
        '        Case Is = "find"
        '            'set the form
        '            'search by form name or tag?
        '
        '        Case Is = "callbyname"
        '
        '        Case Is = "caption"
        '
        '        Case Is = "visible"
        '
        '        Case Is = "windowstate"
        '            VbCallType.VbMethod
        '
        '    End Select
        '
        '    theForm.Caption = ArgList(0)
        '
        'Dim cCont As Control
        '
        'For Each cCont In theForm.Controls
        '    Debug.Print cCont.Name & " - " & TypeName(cCont)
        'Next

            
        Case "stop"
            ExitScript
            
        ' String Manipulating Junk
        Case "mid"
            ExecFunction = Mid$(ArgList(0), ArgList(1), ArgList(2))
        Case "lcase"
            ExecFunction = LCase(ArgList(0))
        Case "ucase"
            ExecFunction = UCase(ArgList(0))
        Case "instr"
            ExecFunction = InStr(ArgList(0), ArgList(1), ArgList(2))
        Case "instrrev"
            ExecFunction = InStrRev(ArgList(0), ArgList(1), ArgList(2))
        Case "strcomp"
            ExecFunction = StrComp(ArgList(0), ArgList(1))
        Case "safestr", "makestrsafe", "formatstr"
            ExecFunction = FormatStrForVar(ArgList(0))
        Case "len"
            ExecFunction = Len(ArgList(0))
        Case "left"
            ExecFunction = Left(ArgList(0), ArgList(1))
        Case "right"
            ExecFunction = Right(ArgList(0), ArgList(1))
        Case "chr"
            ExecFunction = Chr(ArgList(0))
            'Exit Function
        Case "asc"
            ExecFunction = Asc(ArgList(0))
        Case "reverse"
            ExecFunction = StrReverse(ArgList(0))
        Case "replace"
            ExecFunction = Replace(ArgList(0), ArgList(1), ArgList(2))
        Case "trimchar"
            Select Case UBound(ArgList)
                Case Is = 1
                    ExecFunction = TrimChar(ArgList(0), CInt(ArgList(1)))
                Case Is = 2
                    ExecFunction = TrimChar(ArgList(0), ArgList(1), CInt(ArgList(2)))
            End Select
        Case "varexists", "variableexists"
            ExecFunction = VarExists(ArgList(0))
            
        Case "split"
            arrTemp = Split(ArgList(1), ArgList(2))
            For i = 0 To UBound(arrTemp)
                AddVar ArgList(0) & "[" & i & "]", arrTemp(i)
            Next i
            
            'Either call the reference varname itself, or add Ubound_ to make it easier to understand in script
            AddVar "Ubound_" & ArgList(0), UBound(arrTemp)
            AddVar ArgList(0), UBound(arrTemp)
            ExecFunction = UBound(arrTemp)
        
        Case "join"
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
                    
        Case "ubound"
            ExecFunction = GetVar("Ubound_" & ArgList(0))
        
        ' Math stuff
        Case "cos"
            ExecFunction = Cos(ArgList(0))
        Case "tan"
            ExecFunction = Tan(ArgList(0))
        Case "log"
            ExecFunction = Log(ArgList(0))
        Case "rnd"
            ExecFunction = Rnd(ArgList(0))
        Case "randomize"
            Randomize
        Case "int"
            ExecFunction = Int(ArgList(0))
        Case "val"
            ExecFunction = Val(ArgList(0))
        
        ' lol wut?
        'Case "printer_print"
        '    Printer.Print argList(0)
        'Case "finish"
        '    Printer.EndDoc
        
        ' Other Stuff
        'Case "shell"
        '    Shell argList(0)
        
        Case "clipboard"
            Select Case ArgList(0)
                Case Is = "clear"
                    Clipboard.Clear
                Case Is = "get", "gettext"
                    ExecFunction = Clipboard.GetText(ArgList(1))
                Case Is = "set", "settext"
                    Clipboard.SetText ArgList(0)
            End Select
        
        Case "time"
            ExecFunction = Time
            
        Case "now"
            ExecFunction = Now
        
        Case "doevents"
            DoEvents
            
        Case "pause"
            Pause CLng(ArgList(0))
            
        Case "print"
            PrintDebug ArgList(0)
            RaiseEvent DebugOut(ArgList(0))
            
        Case "debug"
            Select Case LCase$(ArgList(0))
            
                Case Is = "on", "true", 1
                    blDebug = True
                Case Is = "off", "false", 0
                    blDebug = False
                Case Is = "show"
                    If myDebug Is Nothing Then Set myDebug = New frmDebug
                    myDebug.Visible = True
                Case Is = "hide"
                    myDebug.Visible = False
                    Unload myDebug
            End Select
            
            PrintDebug "Debug: " & blDebug
            
        Case "showerrors", "errors"
            Select Case LCase$(ArgList(0))
            
                Case Is = "on", "true", 1
                    blShowErrors = True
                Case Is = "off", "false", 0
                    blShowErrors = False
            End Select
            
            PrintDebug "Show Errors: " & blDebug
            
        Case "getname", "username"
            ExecFunction = "Bradley"
            
        Case "getgamepath"
            ExecFunction = App.Path & "\test\"
        
        Case "getscreenwidth", "getx"
            ExecFunction = "1920"
        
        Case "getscreenheight", "gety"
            ExecFunction = "1080"
        
        Case "getos"
            'ExecFunction = getosver
        
        Case "getip"
            ExecFunction = "127.0.0.1"
            
        Case "is64bit"
            'If is64bit Then ExecFunction = True
            ExecFunction = True
        
        Case "downloadfile"
            'ExecFunction = downloadfile(ArgList(0), ArgList(1))
            
        Case "findstr", "findstrinfile"
            'find a str in file and return its position
            
        Case "findline", "findstrbyline"
        
        Case "replacestr", "replacestrinfile"
            'replace a specific string of characters in a file.
            'can replace all or one. This will be a tricky one to use.
            
        Case "replaceline", "replacelineinfile"
            'replace line by line number
            
        Case "fileexists", "exists"
            'ExecFunction = FileExists(ArgList(0))
        Case "folderexists", "direxists"
            'ExecFunction = DirExists(ArgList(0))
        Case "filecopy", "copy"
            FileCopy ArgList(0), ArgList(1)
        Case "deletefolder", "rmdir"
             RmDir ArgList(0)
        Case "deletefile", "kill"
            Kill ArgList(0)
        Case "createfolder", "mkdir"
            MkDir ArgList(0)
        Case "renamefolder", "renfolder"
        
        Case "renamefile", "renfile"
        
        Case "createfile", "newfile"
        
        Case "appendfile"
        
        Case "loadfile", "readfile"
            'need to append a special string to the beginning and end
            'so that the text inbetween does not get modified!
            'AddVar ArgList(0), LoadFile(ArgList(1))
            'ExecFunction = True
            
            Exit Function
        
        Case "chdir", "workingdir"
             ChDir ArgList(0)
            
        Case Else
            'call NotifyError("Unknown Function: " & sFunction, nCurrentLine)

    End Select
    
    ExecFunction = Chr(34) & ExecFunction & Chr(34)
    
    Exit Function
    
ErrorCatch:
    Select Case err.Number
        Case 9
            Call NotifyError("Argument not optional in call to procedure: " & sFunction, nCurrentLine)
    End Select
    Resume Next
End Function

Private Sub AddVar(TheVar As String, NewVal As Variant)
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            Variables.Remove i
            VarNames.Remove i
            
            Variables.Add NewVal
            VarNames.Add TheVar
            Exit Sub
        End If
    Next i
    
    'We didn't find a matching variable.. so let's add one
    Variables.Add NewVal
    VarNames.Add TheVar
    
End Sub

Private Function VarExists(VarName As String) As Boolean
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = VarName Then
        PrintDebug "VarExists (" & VarName & ") = True"
            VarExists = True
            Exit Function
        End If
    Next i
    PrintDebug "VarExists (" & VarName & ") = False"
End Function

Private Sub ChangeVar(TheVar As String, NewVal As Variant)
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            PrintDebug "ChangeVar: " & TheVar & "from " & GetVar(TheVar) & " to: " & NewVal

            Variables.Remove i
            VarNames.Remove i
            
            Variables.Add NewVal
            VarNames.Add TheVar
            Exit Sub
        End If
    Next i
    
    PrintDebug "ChangeVar: " & TheVar & " doesn't exist!"
    
End Sub

Private Function GetVar(TheVar As String) As Variant
    Dim i As Long
    Dim strTemp As String
    Dim strArrVar As String
    'If Len(TheVar) < 3 Then GetVar = TheVar: Exit Function
    PrintDebug "GetVar: " & TheVar
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
            PrintDebug "GetVar: " & TheVar & ": " & Variables(i)
            Exit Function
        End If
    Next i
    
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
    
    PrintDebug "Eval: " & sFormula
    
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
        
    'Todo add left/right true verification if empty string
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
    
    PrintDebug "Eval Solved: " & Eval
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
    If LenB(Equation) = 0 Then Equation = sEquation
    PrintDebug "Equation: " & Equation
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
    
    PrintDebug "Solve: " & Solve
    
End Function

Private Function DoFunction(sCall As String) As Variant

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
    
    PrintDebug "DoFunction: " & sName & " sCall: " & sCall
    
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
    DoFunction = ExecFunction(sName, ArgList)
    
    PrintDebug "DoFunction solved: " & DoFunction
    
Exit Function

err:
    Call NotifyError(err.Number & ": " & err.Description, nCurrentLine)
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
    
    PrintDebug "ProcessFunctions: " & strString
    
    If lngFuncs = -1 Then
        If Right(strString, 2) = "))" Then strString = Left(strString, Len(strString) - 1)
        ProcessFunctions = strString
        Exit Function
    End If
    If lngFuncs = 0 Then
        Debug.Assert "huh?"
    End If
    If lngFuncs = 1 Then
        'If Right(strString, 2) = "))" Then strString = Left(strString, Len(strString) - 1)
        arrFunctions = GetNextFunction(strString) 'This also process functions within functions
        strString = Replace(strString, arrFunctions(0), arrFunctions(1))
        'lngNext = InStr(1, strString, arrFunctions(1)) + Len(arrFunctions(1))
        PrintDebug "ProcessFunctions: Solved Function [ " & arrFunctions(0) & " ] for: [ " & arrFunctions(1) & " ]"
        ProcessFunctions = strString
        Exit Function
    End If
    
    If lngFuncs > 0 Then lngFuncs = lngFuncs - 1
    ReDim arrFunctions(lngFuncs)
        
        
    If UBound(arrFunctions) >= 0 Then
    
        arrFunctions = GetNextFunction(strString) 'This also process functions within functions
        strString = Replace(strString, arrFunctions(0), arrFunctions(1), 1, 1)
        lngNext = InStr(1, strString, arrFunctions(1)) + Len(arrFunctions(1))
        PrintDebug "ProcessFunctions: Solved Function [ " & arrFunctions(0) & " ] for: [ " & arrFunctions(1) & " ]"
        
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
            PrintDebug "ProcessFunctions: Solved Function [ " & arrFunctions(0) & " ] for: [ " & arrFunctions(1) & " ]"
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

Public Function GetBracketContents(strString As String) As String
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

PrintDebug "NumberOfFunctions: " & NumberOfFunctions & " found in [ " & strString & " ]"

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

Private Sub CleanupVars()
    Dim i As Long
    For i = VarNames.Count To 1 Step -1
        VarNames.Remove i
        Variables.Remove i
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
    PrintDebug "CleanupVars complete. Everything but Debug ;-)"
End Sub

Private Sub ExitScript()
    CleanupVars
    nCurrentLine = UBound(SplitScript)
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
            nCurrentLine = i
            PrintDebug "GotoScriptLine: found [" & Txt & "] at line " & i + 1
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

    PrintDebug "IsPathAllowed: [ " & Path & " ] : [ " & IsPathAllowed & " ]"

End Function

Public Function CallByNameEx(Obj As Object, _
    ProcName As String, CallType As VbCallType, _
    Optional vArgsArray As Variant)
    Dim oTLI As Object
    Dim ProcID As Long
    Dim numArgs As Long
    Dim i As Long
    Dim v()
    
    On Error GoTo Handler
    
    Set oTLI = CreateObject("TLI.TLIApplication")
    ProcID = oTLI.InvokeID(Obj, ProcName)
    
    If IsMissing(vArgsArray) Then
        CallByNameEx = oTLI.InvokeHook( _
            Obj, ProcID, CallType)
    End If
    
    If IsArray(vArgsArray) Then
        numArgs = UBound(vArgsArray)
        ReDim v(numArgs)
        For i = 0 To numArgs
            v(i) = vArgsArray(numArgs - i)
        Next i
        CallByNameEx = oTLI.InvokeHookArray( _
            Obj, ProcID, CallType, v)
    End If
Exit Function

Private Sub NotifyError(Msg As String, Line As Long)

    If blDebug Then
        PrintDebug (Line & ": " & Msg)
    End If
    
    
End Sub

Private Sub PrintDebug(Msg As String)
    If Not blDebug Then Exit Sub
    
    If Not myDebug Is Nothing Then
        Call myDebug.PrintDebug(Msg & vbNewLine)
    Else
        RaiseEvent DebugOut(Msg & vbNewLine)
    End If
    
End Sub

Private Sub Class_Initialize()
     Set InsideLoop = New Collection
End Sub
