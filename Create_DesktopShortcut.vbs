
if  WScript.Arguments.count = 0 then
  
	msgbox "�t�@�C�����h���b�v����Ă��܂���B�������I�����܂��B"
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

'�t�@�C���ݒ�
workfile = fso.GetFileName(pt)
linkfile = fso.GetBaseName(pt) & ".lnk"
workfolder = fso.GetFile(pt).ParentFolder

' �V���[�g�J�b�g�쐬
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
