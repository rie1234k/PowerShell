	$folderpath = Split-Path  $PSScriptRoot　# 起動したいファイルあるのフォルダパス（現在はこのスクリプトの1つ上の階層で設定）
	$filename = "StatingFile.xlsx" # 起動したいファイル名
    $bufName = "MessageFile.xlsx"　# 起動時に開く一時ファイルを用意し、起動したいファイルと同じ場所に置いておく

	$file = $folderpath + "\" + $filename
    $buffile =$folderpath + "\"+ $bufname
    $savefile = [System.Environment]::GetFolderPath("mydocument") +"\"+ $bufName　
        
        Copy-Item -Path $buffile -Destination $savefile -Force #一時ファイルをマイドキュメントに保存

        Invoke-Item $savefile # 一時ファイルを関連付けられたアプリケーションで開く

        Start-Sleep -s 5 # 起動が遅くて、うまくいかない場合があるため、5秒待つ

 	    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") # 一時ファイルを開いたエクセルをアクティブにする

        $book = $excel.Workbooks.Open($file,3,$true) # 起動したいファイルを読み取り専用で開く　引数は左から、FileName、UpdateLinks[3：リンクを更新する]、ReadOnly[$true:読み取り専用で開く]

        add-type -assembly microsoft.visualbasic
        [microsoft.visualbasic.interaction]::AppActivate($filename)　# 起動したファイルをアクティブにする

        # ShowWindowAsyncを使用する準備
        $dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
        Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32

        [Win32.NativeMethods]::ShowWindowAsync($excel.HWND, 3) | Out-Null　# 画面を最大化

        # 開いた一時ファイルを取得
        foreach ($bk in $excel.WorkBooks)
    {
        if ($bk.Name -eq $bufName)
        {
            $targetbook = $bk
        }
    }
        
        $SaveChanges = $False # 変更を保存しない。
        $targetbook.Close($SaveChanges)　

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetbook)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($bk)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        
        Remove-Item $savefile　# 一時ファイルを削除
