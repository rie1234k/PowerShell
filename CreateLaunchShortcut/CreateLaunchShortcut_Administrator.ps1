    if($Args.Count -eq 0){
        $f = Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($f,"ファイルがドロップされていません。処理を終了します。","メッセージ",[System.Windows.Forms.MessageBoxButtons]::OK)
        exit 
    }

    foreach ($arg in $Args){  
        $myfile = Get-Item -LiteralPath $arg
        $linkfile = $myfile.BaseName + ".lnk"
        $workfolder = $myfile.DirectoryName
        $opt = "-ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "
        $WsShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WsShell.CreateShortcut((Join-Path $workfolder $linkfile))
        $Shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Shortcut.Arguments = $opt + '"'+ $myfile.FullName + '"'
        $Shortcut.WorkingDirectory = $workfolder
        $Shortcut.WindowStyle = 7
        $Shortcut.Save()

        #管理者権限で実行にチェックを入れる
        $offset = 0x15 # 管理者スイッチオフセット
        $Path = (Join-Path $workfolder $linkfile)
        $byteReader = [System.IO.File]::ReadAllBytes($Path)
        $byteReader[$offset] = 0x20  # 0x00 無効
        [System.IO.File]::WriteAllBytes($Path, $byteReader)

    }