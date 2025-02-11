# Excel�A�v���P�[�V�������N��
$excel = New-Object -ComObject Excel.Application
# �o�b�N�O���E���h�Ŏ��s
# $excel.Visible = $false

# �����Ƀt�H���_�̃p�X���w�肵�Ď��s

try {
    # �����ɂ��邷�ׂẴt�@�C���֏���
    foreach ($file in $args) {
        Write-Host "Processing file: $($file)"
        
        $workbook = $excel.Workbooks.Open($file)
    
        # �ύX��ۑ�
        $workbook.Save()
    
        # ���[�N�u�b�N�����
        $workbook.Close()
    
        # COM�I�u�W�F�N�g�̉��
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
}
    catch {
    Write-Host "An error occurred: $_"
} finally {
    # Excel�A�v���P�[�V�������I��
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel application closed."
}