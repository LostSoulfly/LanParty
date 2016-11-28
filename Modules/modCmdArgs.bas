Attribute VB_Name = "modCmdArgs"
Option Explicit

'=======================================
'     ============================
' GetCommandArgs - © Nik Keso 2009
'----------------------------------
'The function returns an array with the
'     command line arguments,
'contained in the command$, like cmd.exe


Public Function GetCommandArgs() As String()
Dim CountQ As Integer 'chr(34) counter
Dim OpenQ As Boolean ' left open string indicator (ex:"c:\bbb ccc.bat ) OpenQ=true, (ex:"c:\bbb ccc.bat" ) OpenQ=false
Dim ArgIndex As Integer
Dim tmpSTR As String
Dim strIndx As Integer
Dim TmpArr() As String
Dim comSTR As String
    GetCommandArgs = Split("", " ")
    TmpArr = Split("", " ")
    comSTR = Trim$(Command$) 'remove front and back spaces
    If Len(comSTR) = 0 Then Exit Function
    CountQ = UBound(Split(comSTR, """"))
    If CountQ Mod 2 = 1 Then Exit Function 'like cmd.exe , command$ must contain even number of chr(34)=(")
    strIndx = 1

    Do
        If Mid$(comSTR, strIndx, 1) = """" Then OpenQ = Not OpenQ
        If Mid$(comSTR, strIndx, 1) = " " And OpenQ = False Then
            If tmpSTR <> "" Then 'don't include the spaces between args as args!!!!!
                ReDim Preserve TmpArr(ArgIndex)
                TmpArr(ArgIndex) = tmpSTR
                ArgIndex = ArgIndex + 1
            End If
            tmpSTR = ""
        Else
            tmpSTR = tmpSTR & Mid$(comSTR, strIndx, 1)
        End If
        strIndx = strIndx + 1
    Loop Until strIndx = Len(comSTR) + 1
    
    ReDim Preserve TmpArr(ArgIndex)
    TmpArr(ArgIndex) = tmpSTR
    GetCommandArgs = TmpArr
End Function


Public Function GetArgs(myString As String) As String()
Dim CountQ As Integer 'chr(34) counter
Dim OpenQ As Boolean ' left open string indicator (ex:"c:\bbb ccc.bat ) OpenQ=true, (ex:"c:\bbb ccc.bat" ) OpenQ=false
Dim ArgIndex As Integer
Dim tmpSTR As String
Dim strIndx As Integer
Dim TmpArr() As String
Dim comSTR As String
    GetArgs = Split("", " ")
    TmpArr = Split("", " ")
    comSTR = Trim$(myString) 'remove front and back spaces
    If Len(comSTR) = 0 Then Exit Function
    CountQ = UBound(Split(comSTR, """"))
    If CountQ Mod 2 = 1 Then Exit Function
    strIndx = 1
    'todo fix this. It needs to ignore spaces inside of quotes
    ' and only process them if they're comma separated!
    Do
        If Mid$(comSTR, strIndx, 1) = """" Then OpenQ = Not OpenQ
        If Mid$(comSTR, strIndx, 1) = Chr(34) And OpenQ = False Then
            If tmpSTR <> "" Then 'don't include the spaces between args as args!!!!!
                ReDim Preserve TmpArr(ArgIndex)
                If Right(tmpSTR, 2) = Chr(34) & "," Then tmpSTR = Left(tmpSTR, Len(tmpSTR) - 1)
                TmpArr(ArgIndex) = tmpSTR
                ArgIndex = ArgIndex + 1
            End If
            tmpSTR = ""
        Else
            tmpSTR = tmpSTR & Mid$(comSTR, strIndx, 1)
        End If
        strIndx = strIndx + 1
    Loop Until strIndx = Len(comSTR) + 1
    
    ReDim Preserve TmpArr(ArgIndex)
    If Right(tmpSTR, 2) = Chr(34) & "," Then tmpSTR = Left(tmpSTR, Len(tmpSTR) - 1)
    TmpArr(ArgIndex) = tmpSTR
    GetArgs = TmpArr
End Function
        

Public Function AddToArgs(strSwitch As String, strCommand As String) As String
Dim lngPosition

lngPosition = InStr(1, strCommand, " -a ")
If lngPosition = 0 Then lngPosition = InStr(1, strCommand, " -a2 ")

If lngPosition > 0 Then
    AddToArgs = Mid(strCommand, 1, lngPosition) & strSwitch & Mid(strCommand, lngPosition, Len(strCommand) - lngPosition + 1)
    Exit Function
End If

AddToArgs = strCommand & " " & strSwitch
End Function
