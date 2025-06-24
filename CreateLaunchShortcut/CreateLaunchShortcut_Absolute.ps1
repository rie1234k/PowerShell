
    if($Args.Count -eq 0){
        $f = Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($f,"ファイルがドロップされていません。処理を終了します。","メッセージ",[System.Windows.Forms.MessageBoxButtons]::OK)
        exit 
    }

    $Args | foreach{  
        $myfile = Get-Item -LiteralPath $_
        $workfile = $myfile.Name
        $linkfile = $myfile.BaseName + "_Run.lnk"
        $workfolder = $myfile.DirectoryName
        $opt = "-ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "
        $WsShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WsShell.CreateShortcut($workfolder + "\" + $linkfile)
        $Shortcut.TargetPath = "powershell"
        $Shortcut.Arguments = $opt + '"'+ $workfolder + "\" + $workfile + '"'
        $Shortcut.WorkingDirectory = $workfolder
        $Shortcut.WindowStyle = 7
        $Shortcut.Save()
    }
    
