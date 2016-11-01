Attribute VB_Name = "modScript"

Public Function RunScriptSub(sScript As String, sSub As String) As Variant

If Not Settings.AllowScripts Then Exit Function
Dim cScript As New clsScript

    cScript.Script = sScript
    
    'Expose some variables (this is one-way, the script can't change them)
    cScript.AddVar "UserName", Settings.UserName
    cScript.AddVar "UniqueIP", Settings.UniqueID
    cScript.AddVar "GameEXE", Game(GameIndex).GameEXE
    cScript.AddVar "GameInstalled", Game(GameIndex).Installed
    cScript.AddVar "GameInstallFirst", Game(GameIndex).InstallFirst
    cScript.AddVar "GameInstallerPath", Game(GameIndex).InstallerPath
    cScript.AddVar "GameCmdArgs", Game(GameIndex).CommandArgs
    cScript.AddVar "GameEXEPath", Game(GameIndex).EXEPath
    cScript.AddVar "GameUID", Game(GameIndex).GameUID
    cScript.AddVar "GameName", Game(GameIndex).Name
    cScript.AddVar "GameType", Game(GameIndex).GameType
    cScript.AddVar "CurrentIP", Settings.CurrentIP
    cScript.AddVar "LanAdmin", Settings.LanAdmin
    cScript.AddVar "AllowDownload", Settings.ScriptDownload
    cScript.AddVar "AllowExecute", Settings.ScriptExecute
    
    RunScriptSub = cScript.FindAndExecuteSub(sSub)
    Set cScript = Nothing
    
End Function
