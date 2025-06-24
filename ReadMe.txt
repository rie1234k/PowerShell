PowerShell関連 各ファイルの説明

【MessageFile.xlsx】
ExcelReadOnlyOpen_Coexist.ps1スクリプト用のサンプルファイル

【StatingFile.xlsx】
ExcelReadOnlyOpen_Coexist.ps1スクリプト用のサンプルファイル



《ExcelReadOnlyOpen》フォルダ
【ExcelReadOnlyOpen_Coexist.ps1】
エクセルを読み取り専用、ウィンドウ最大化で開く。
既にエクセルを開いている場合、そのエクセルに追加してファイルが開かれる。

【ExcelReadOnlyOpen_Independent.ps1】
エクセルを読み取り専用、ウィンドウ最大化で開く。
既にエクセルを開いていても、そのエクセルとは別に、独立して新しいエクセルを開く形となっている。



《FolderDeleteTool》フォルダ
【FolderDeleteTool.ps1】
指定したフォルダを削除するツール
進捗状況表示付き



《CreateLaunchShortcut》フォルダ
【CreateLaunchShortcut.ps1】
PowerShellスクリプト(ps1ファイル)をドラッグアンドドロップして、起動用ショートカットを作成する。
複数ファイルのドラッグアンドドロップにも対応。
相対パス仕様。

【CreateLaunchShortcut_Absolute.ps1】
PowerShellスクリプト(ps1ファイル)をドラッグアンドドロップして、起動用ショートカットを作成する。
複数ファイルのドラッグアンドドロップにも対応。
絶対パス仕様。

【CreateLaunchShortcut.lnk】
「CreateLaunchShortcut.ps1」の起動用ショートカット。
このショートカットにPowerShellスクリプト(ps1ファイル)をドラッグアンドドロップすると、
起動用ショートカットが作成される。

【CreateLaunchShortcut.bat】
このバッチファイルに、PowerShellスクリプト(ps1ファイル)をドラッグアンドドロップすると、
CreateLaunchShortcut.ps1にファイルを渡して実行する。
CreateLaunchShortcut.ps1の起動用ショートカットを作成するためのバッチファイル。