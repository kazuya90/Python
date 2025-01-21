# Excelアプリケーションを起動
$excel = New-Object -ComObject Excel.Application
# バックグラウンドで実行
# $excel.Visible = $false

try {
    # フォルダ内のすべての.xlsxファイルを処理
    # 入力を求める
    $folderPath = Read-Host "フォルダのパスを入力してください"
    $files = Get-ChildItem -LiteralPath $folderPath -Filter *.xlsx

    foreach ($file in $files) {
        # ファイル名を表示
        Write-Host "Processing file: $($file.FullName)"
        
        $workbook = $excel.Workbooks.Open($file.FullName)
        
        # 変更を保存
        $workbook.Save()
        
        # ワークブックを閉じる
        $workbook.Close()
        
        # COMオブジェクトの解放
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
} catch {
    Write-Host "An error occurred: $_"
} finally {
    # Excelアプリケーションを終了
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel application closed."
}