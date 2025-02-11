# Excelアプリケーションを起動
$excel = New-Object -ComObject Excel.Application
# バックグラウンドで実行
# $excel.Visible = $false

# 引数にフォルダのパスを指定して実行

try {
    # 引数にあるすべてのファイルへ処理
    foreach ($file in $args) {
        Write-Host "Processing file: $($file)"
        
        $workbook = $excel.Workbooks.Open($file)
    
        # 変更を保存
        $workbook.Save()
    
        # ワークブックを閉じる
        $workbook.Close()
    
        # COMオブジェクトの解放
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
}
    catch {
    Write-Host "An error occurred: $_"
} finally {
    # Excelアプリケーションを終了
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel application closed."
}