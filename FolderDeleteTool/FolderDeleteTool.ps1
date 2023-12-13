

# 調査対象フォルダの初期化
$SearchFolder = ""

# 空の配列を作成
$arr = @()

# アセンブリ読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# フォーム作成
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(400,130)
$Form.Text = "フォルダ削除ツール"
# ラベル作成
$LabelFilePath = New-Object System.Windows.Forms.Label
$LabelFilePath.Location = New-Object System.Drawing.Point(20,10)
$LabelFilePath.Size = New-Object System.Drawing.Size(300,20)
$LabelFilePath.Text = "削除したいフォルダのパスを入力してください"
$Form.Controls.Add($LabelFilePath)
# 入力用テキストボックス
$TextBoxFilePath = New-Object System.Windows.Forms.TextBox
$TextBoxFilePath.Location = New-Object System.Drawing.Point(20,30)
$TextBoxFilePath.Size = New-Object System.Drawing.Size(300,20)
$Form.Controls.Add($TextBoxFilePath)
# 参照ボタン
$ButtonFilePath = New-Object System.Windows.Forms.Button
$ButtonFilePath.Location = New-Object System.Drawing.Point(320,30)
$ButtonFilePath.Size = New-Object System.Drawing.Size(40,20)
$ButtonFilePath.Text = "参照"
$Form.Controls.Add($ButtonFilePath)
# OKボタン
$ButtonOK = New-Object System.Windows.Forms.Button
$ButtonOK.Location = New-Object System.Drawing.Point(230,60)
$ButtonOK.Size = New-Object System.Drawing.Size(60,20)
$ButtonOK.Text = "OK"
$Form.Controls.Add($ButtonOK)
# Cancelボタン
$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Location = New-Object System.Drawing.Point(300,60)
$ButtonCancel.Size = New-Object System.Drawing.Size(60,20)
$ButtonCancel.Text = "キャンセル"
$ButtonCancel.DialogResult = "Cancel"
$Form.Controls.Add($ButtonCancel)
# 参照ボタンをクリック時の動作
$ButtonFilePath.add_click{
#ダイアログを表示しファイルを選択する
$Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
if($Dialog.ShowDialog() -eq "OK"){
$TextBoxFilePath.Text = $Dialog.SelectedPath
}
}
# OKボタンをクリック時の動作
$ButtonOK.add_click{
#ファイルパスが入力されていないときは背景を黄色にする
if($TextBoxFilePath.text -eq ""){
$TextBoxFilePath.BackColor = "yellow"
}else{
$Form.DialogResult = "OK"
}
}

#フォームを表示し処理が完了したらファイルパスを返す
$FormResult = $Form.ShowDialog()
if($FormResult -eq "OK"){

    $SearchFolder = $TextBoxFilePath.text

    $arr = (Get-ChildItem -Recurse $SearchFolder| ? { ! $_.PSIsContainer }).FullName

    echo "フォルダ削除を開始します..."

    #全体数
    $TotalNumber = (Get-ChildItem -Recurse $SearchFolder | ? { ! $_.PSIsContainer } | Measure-Object).Count

    #処理完了数。初期値は0とする。
    $counter = 0

    #進捗状況をプログレスバーだけでなく数字でも表すために、分母として全体数を文字列として定義（取得）した変数を用意する。
    $denominator = "/"+[string]$TotalNumber

    #ここからプログラム処理。ここではWhile文を用いて「counterの値がTotalNumberより小さい間は処理を繰り返す」としている
    while($counter -lt $TotalNumber) {
    
        #1処理完了毎にcounterをカウントアップさせる。
        $counter ++;

        #処理完了数 / 全体数を進捗状況とするので、予め変数として規定しておく
        $per = ($counter / $TotalNumber * 100)
    
        #ここが進捗状況（プログレスバー）の設定行。 -activityはプログレスバーの表記名。
        #-statusは処理中の処理状況を数値でも表したい場合に設定する。ここでは、counter / 分母（denominator）としている。
        #-percentCompleteがプログレスバーを作る部分。「処理完了数 / 全体数」である変数$perを指定している。
        Write-Progress -activity "進捗状況" -status $counter$denominator -percentComplete $per

        #ここから処理したい内容を記載していく
        if($TotalNumber -eq 1){
            
            Remove-Item -Path $arr -Force #削除するファイルが1件の場合、$arrは配列になっていない　 添え字で指定すると文字列の指定した文字になってしまう

        }else{
            
            Remove-Item -Path $arr[$counter - 1] -Force
        
        }
    }

    Remove-Item -path $SearchFolder -Recurse
    
    $wsobj = new-object -comobject wscript.shell
    $result = $wsobj.popup("フォルダ削除が完了しました。")

}
