	$folderpath = Split-Path  $PSScriptRoot -Parent # 起動したいファイルあるのフォルダパス（現在はこのスクリプトの1つ上の階層で設定）
	$filename = "StatingFile.xlsx"

	$file = $folderpath + "\" + $filename　

    
    # 起動済みかどうか確認する
    foreach ($title in (Get-Process Excel).MainWindowTitle)
    {
        if ($title -like "*"+$filename+"*")
        {

            # アセンブリの読み込み
            Add-Type -Assembly System.Windows.Forms

            # メッセージボックスの表示
            [System.Windows.Forms.MessageBox]::Show("既に開いているため、開けません。", "メッセージ")

            # 既に起動済みのファイルをアクティブにする
            add-type -assembly microsoft.visualbasic
            [microsoft.visualbasic.interaction]::AppActivate($filename)

            $dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
            Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32

            [Win32.NativeMethods]::ShowWindowAsync( (Get-Process | Where-Object{$_.MainWindowTitle -like '*'+$filename+'*'}).MainWindowHandle, 3) | Out-Null　# ウィンドウの最大化

            exit
        }
    }
        
 	    $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true

        add-type -assembly microsoft.visualbasic
        [microsoft.visualbasic.interaction]::AppActivate((Get-Process | Where-Object { $_.Mainwindowtitle -eq "EXCEL" }).id)

        $dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
        Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32
        [Win32.NativeMethods]::ShowWindowAsync($excel.HWND, 3) | Out-Null

        $book = $excel.Workbooks.Open($file,3,$true)

        $book = $null
        $excel = $null
        [GC]::Collect()