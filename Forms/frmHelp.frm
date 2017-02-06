VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Help"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExample 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   5295
   End
   Begin VB.TextBox txtDescription 
      Height          =   1365
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   5295
   End
   Begin VB.TextBox txtArgs 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   5295
   End
   Begin VB.ListBox lstHelp 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colName As Collection
Private colArgs As Collection
Private colDesc As Collection
Private colExam As Collection

Private Sub AddHelpItem(sName As String, sArgs As String, sDesc As String, Optional sExam As String)

    colName.Add sName
    colArgs.Add sArgs
    colDesc.Add sDesc
    colExam.Add sExam

End Sub

Private Sub Form_Load()

    Set colName = New Collection
    Set colArgs = New Collection
    Set colDesc = New Collection
    Set colExam = New Collection

    lstHelp.FontSize = 14

    AddHelpItem "msgbox", "Message, Style(int), Caption", "Show a messagebox", "~msgbox(""Hello world"", ""8"", ""Message!"")"
    AddHelpItem "msgbox", "Message, Caption", "Show a messagebox", "~msgbox(""Hello world"", ""Message!"")"
    AddHelpItem "inputbox", "Message, Caption, Default Text", "Show an box that returns a user's input", "~inputbox(""What's your name?"", ""Hello"", ""Bob"")"
    AddHelpItem "app", "SWITCH: major, minor, revision, exename, pathtitle, threadid", "Returns the requested VB6 App variable."
    AddHelpItem "stop", "", "Stops processing the current script."
    AddHelpItem "shutdown", "", "Ends the program completely."
    AddHelpItem "mid", "Text, Start(int), Length(int)", "Returns text from a string at a certain point, extending the number of characters supplied", "~mid(""H E L L O"", ""3"", ""1"")"
    AddHelpItem "lcase", "Text", "Returns the lowercase of a supplied string."
    AddHelpItem "ucase", "Text", "Returns the uppercase of a supplied string."
    AddHelpItem "instr", "Start(int), Where, What", "Returns the position of WHAT in string WHERE, after Start(int)."
    AddHelpItem "instrrev", "Where, What", "Opposite of instr, but always started from end."
    AddHelpItem "strcomp", "String1, String2", "Return result of comparing two strings."
    AddHelpItem "safestr, formatstr", "String/Variable", "Returns a formatted string for use in LSL."
    AddHelpItem "len", "String/Variable", "Returns the length of the supplied string or variable."
    AddHelpItem "left", "String/Variable, Length(int)", "Returns text from the left side of the string extending Length(int)."
    AddHelpItem "right", "String/Variable, Length(int)", "Returns text from the right side of the string extending Length(int)."
    addehlpitem "textblock", "", "End with 'end textblock'"
    AddHelpItem "chr", "", ""
    AddHelpItem "asc", "", ""
    AddHelpItem "reverse", "", ""
    AddHelpItem "replace", "", ""
    AddHelpItem "trimchar", "", ""
    AddHelpItem "varexists", "", "Returns True or False depending on whether or not a variable exists."
    AddHelpItem "split", "", ""
    AddHelpItem "join", "", ""
    AddHelpItem "ubound", "", ""
    AddHelpItem "rnd", "", ""
    AddHelpItem "randomnum", "", ""
    AddHelpItem "randomize", "", ""
    AddHelpItem "int", "", ""
    AddHelpItem "val", "", ""
    AddHelpItem "execute", "", ""
    AddHelpItem "shell", "", ""
    AddHelpItem "clipboard", "SWITCH: clear, get/gettext, set/settext: Text", ""
    AddHelpItem "time", "", "Returns the time: HH:MM:SS PM/AM"
    AddHelpItem "now", "", "Returns the current date: MM/DD/YYYY HH:MM:SS PM/AM"
    AddHelpItem "doevents", "", "Forces the interpreter to call VB's DoEvents."
    AddHelpItem "pause", "", "Causes the script to pause for xx milliseconds. Recommended to use multiples of 10."
    AddHelpItem "print", "Text", "Prints text to the debug console, with a debug level of 0."
    AddHelpItem "debug", "SWITCH: on/true, off/false, show, hide, 0, 1, 2, 3, clear", "Configure debug output level or clear debug screen."
    AddHelpItem "showerrors, errors", "SWITCH: on/true/1, off/false/0", "Enable or disable reporting of errors"
    AddHelpItem "getname, username", "SWITCH: 1/2", "Returns the user's 1: chat name or 2: Windows username."
    AddHelpItem "getgamepath", "", "Returns the current game's path."
    AddHelpItem "getscreenwidth, getx", "", "Returns the current screen width."
    AddHelpItem "getscreenheight, gety", "", "Returns the current screen height."
    AddHelpItem "getos", "", "Returns a string with the OS version information."
    AddHelpItem "getip", "", "Returns a string with the current IP address."
    AddHelpItem "is64bit", "", "Returns True if on a 64bit OS."
    AddHelpItem "filesize, getsize, filelen", "File Path", "Returns the size of a file, if it exists."
    AddHelpItem "fileexists, exists", "File Path", "Returns True/False whether file exists or not."
    AddHelpItem "folderexists, direxists", "Dir Path", "Returns True/False whether directory exists or not."
    AddHelpItem "filecopy, copy", "Source Path, Dest Path", "Copy a file from Source to Destination location."
    AddHelpItem "deletefolder, rmdir", "Dir Path", "Deletes a folder or directory."
    AddHelpItem "deletefile, kill", "File Path", "Deletes a file from disk."
    AddHelpItem "createfolder, mkdir", "Dir Name", "Create a folder."
    AddHelpItem "renamefolder, renfolder", "Source Path, Dest Path", "Rename a folder."
    AddHelpItem "renamefile, renfile", "Source Path, Dest Path", "Rename a file."
    AddHelpItem "createfile, newfile", "File Path, Data", "Create a new file with supplied Data."
    AddHelpItem "appendfile", "", ""
    AddHelpItem "isfileloaded", "Variable Name", ""
    AddHelpItem "loadfile, readfile", "File Path, FileVar Name", ""
    AddHelpItem "filetovar", "File Path, Variable Name", "Load to a Variable instead of a FileVar."
    AddHelpItem "findpos, findposinfile", "", ""
    AddHelpItem "getline, getfileline", "", ""
    AddHelpItem "replaceline ", "", ""
    AddHelpItem "replacestrinfile", "", ""
    AddHelpItem "findline, findlinebystr", "", ""
    AddHelpItem "downloadfile", "", ""
    AddHelpItem "chdir, workingdir", "", ""
    AddHelpItem "getfilelist, filelist", "", ""
    AddHelpItem "getdirectories, dirlist", "", ""
    AddHelpItem "readini", "", ""
    AddHelpItem "writeini", "", ""
    
    PopulateList
    
End Sub

Private Sub PopulateList()
Dim i As Long
lstHelp.Clear
For i = 1 To colName.Count

    lstHelp.AddItem colName(i)

Next i

End Sub

Private Sub lstHelp_Click()
    On Error Resume Next
    
    txtArgs.Text = colName(lstHelp.ListIndex + 1) & "(" & colArgs(lstHelp.ListIndex + 1) & ")"
    txtDescription.Text = colDesc(lstHelp.ListIndex + 1)
    txtExample.Text = colExam(lstHelp.ListIndex + 1)
    
End Sub
