	$folderpath = Split-Path  $PSScriptRoot�@# �N���������t�@�C������̃t�H���_�p�X�i���݂͂��̃X�N���v�g��1��̊K�w�Őݒ�j
	$filename = "StatingFile.xlsx" # �N���������t�@�C����
    $bufName = "MessageFile.xlsx"�@# �N�����ɊJ���ꎞ�t�@�C����p�ӂ��A�N���������t�@�C���Ɠ����ꏊ�ɒu���Ă���

	$file = $folderpath + "\" + $filename
    $buffile =$folderpath + "\"+ $bufname
    $savefile = [System.Environment]::GetFolderPath("mydocument") +"\"+ $bufName�@
        
        Copy-Item -Path $buffile -Destination $savefile -Force #�ꎞ�t�@�C�����}�C�h�L�������g�ɕۑ�

        Invoke-Item $savefile # �ꎞ�t�@�C�����֘A�t����ꂽ�A�v���P�[�V�����ŊJ��

        Start-Sleep -s 5 # �N�����x���āA���܂������Ȃ��ꍇ�����邽�߁A5�b�҂�

 	    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") # �ꎞ�t�@�C�����J�����G�N�Z�����A�N�e�B�u�ɂ���

        $book = $excel.Workbooks.Open($file,3,$true) # �N���������t�@�C����ǂݎ���p�ŊJ���@�����͍�����AFileName�AUpdateLinks[3�F�����N���X�V����]�AReadOnly[$true:�ǂݎ���p�ŊJ��]

        add-type -assembly microsoft.visualbasic
        [microsoft.visualbasic.interaction]::AppActivate($filename)�@# �N�������t�@�C�����A�N�e�B�u�ɂ���

        # ShowWindowAsync���g�p���鏀��
        $dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
        Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32

        [Win32.NativeMethods]::ShowWindowAsync($excel.HWND, 3) | Out-Null�@# ��ʂ��ő剻

        # �J�����ꎞ�t�@�C�����擾
        foreach ($bk in $excel.WorkBooks)
    {
        if ($bk.Name -eq $bufName)
        {
            $targetbook = $bk
        }
    }
        
        $SaveChanges = $False # �ύX��ۑ����Ȃ��B
        $targetbook.Close($SaveChanges)�@

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetbook)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($bk)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        
        Remove-Item $savefile�@# �ꎞ�t�@�C�����폜
