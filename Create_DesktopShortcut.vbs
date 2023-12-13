
if  WScript.Arguments.count = 0 then
  
	msgbox "ファイルがドロップされていません。処理を終了します。"
	WScript.Quit
 
end if

Set GetPathArray = WScript.Arguments

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
Set mywsh = CreateObject("WScript.Shell")


sendfolder = mywsh.SpecialFolders("Desktop") 
powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
Const OPT = "-ExecutionPolicy Bypass -WindowStyle Hidden -File "

For Each pt in GetPathArray

'ファイル設定
workfile = fso.GetFileName(pt)
linkfile = fso.GetBaseName(pt) & ".lnk"
workfolder = fso.GetFile(pt).ParentFolder

' ショートカット作成
Set shortcut = ws.CreateShortcut(sendfolder & "\" & linkfile)
With shortcut
    .TargetPath = powershell
    .Arguments = OPT & workfolder & "\" & workfile
    .WorkingDirectory = workfolder
    .WindowStyle = 7 
    .Save
End With

Next

set fso = nothing
