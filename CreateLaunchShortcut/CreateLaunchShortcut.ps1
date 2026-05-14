    if($Args.Count -eq 0){
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($null,"ファイルがドロップされていません。処理を終了します。","メッセージ",[System.Windows.Forms.MessageBoxButtons]::OK)
        return
    }

    foreach ($arg in $Args){
        $myfile = Get-Item -LiteralPath $arg
        $workfile = $myfile.Name
        $linkfile = $myfile.BaseName + ".lnk"
        $workfolder = $myfile.DirectoryName
        $opt = "-ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "
        $WsShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WsShell.CreateShortcut((Join-Path $workfolder $linkfile))
        $Shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Shortcut.Arguments = $opt + '"'+ $workfile + '"'
        $Shortcut.WorkingDirectory = ""
        $Shortcut.WindowStyle = 7
        $Shortcut.Save()
    }