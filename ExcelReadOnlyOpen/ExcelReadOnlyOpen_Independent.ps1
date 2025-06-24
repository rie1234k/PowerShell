	$folderpath = Split-Path  $PSScriptRoot -Parent # �N���������t�@�C������̃t�H���_�p�X�i���݂͂��̃X�N���v�g��1��̊K�w�Őݒ�j
	$filename = "StatingFile.xlsx"

	$file = $folderpath + "\" + $filename�@

    
    # �N���ς݂��ǂ����m�F����
    foreach ($title in (Get-Process Excel).MainWindowTitle)
    {
        if ($title -like "*"+$filename+"*")
        {

            # �A�Z���u���̓ǂݍ���
            Add-Type -Assembly System.Windows.Forms

            # ���b�Z�[�W�{�b�N�X�̕\��
            [System.Windows.Forms.MessageBox]::Show("���ɊJ���Ă��邽�߁A�J���܂���B", "���b�Z�[�W")

            # ���ɋN���ς݂̃t�@�C�����A�N�e�B�u�ɂ���
            add-type -assembly microsoft.visualbasic
            [microsoft.visualbasic.interaction]::AppActivate($filename)

            $dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
            Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32

            [Win32.NativeMethods]::ShowWindowAsync( (Get-Process | Where-Object{$_.MainWindowTitle -like '*'+$filename+'*'}).MainWindowHandle, 3) | Out-Null�@# �E�B���h�E�̍ő剻

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