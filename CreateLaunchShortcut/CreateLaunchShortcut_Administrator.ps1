    if($Args.Count -eq 0){
        $f = Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($f,"�t�@�C�����h���b�v����Ă��܂���B�������I�����܂��B","���b�Z�[�W",[System.Windows.Forms.MessageBoxButtons]::OK)
        exit 
    }

    $Args | foreach{  
        $myfile = Get-Item -LiteralPath $_        
        $linkfile = $myfile.BaseName + "_Run.lnk"
        $workfolder = $myfile.DirectoryName
        $opt = "-ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "
        $WsShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WsShell.CreateShortcut((Join-Path $workfolder $linkfile))
        $Shortcut.TargetPath = "powershell"
        $Shortcut.Arguments = $opt + '"'+ $myfile + '"'
        $Shortcut.WorkingDirectory = $workfolder
        $Shortcut.WindowStyle = 7
        $Shortcut.Save()

        #�Ǘ��҂Ƃ��Ď��s�I�v�V������L���ɂ���
        $offset = 0x15
        $Path = (Join-Path $workfolder $linkfile)
        $byteReader = [System.IO.File]::ReadAllBytes($Path)
        $byteReader[$offset] = 0x20
        [System.IO.File]::WriteAllBytes($Path, $byteReader)

    }

    
