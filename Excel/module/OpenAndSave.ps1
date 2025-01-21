# Excel�A�v���P�[�V�������N��
$excel = New-Object -ComObject Excel.Application
# �o�b�N�O���E���h�Ŏ��s
# $excel.Visible = $false

try {
    # �t�H���_���̂��ׂĂ�.xlsx�t�@�C��������
    # ���͂����߂�
    $folderPath = Read-Host "�t�H���_�̃p�X����͂��Ă�������"
    $files = Get-ChildItem -LiteralPath $folderPath -Filter *.xlsx

    foreach ($file in $files) {
        # �t�@�C������\��
        Write-Host "Processing file: $($file.FullName)"
        
        $workbook = $excel.Workbooks.Open($file.FullName)
        
        # �ύX��ۑ�
        $workbook.Save()
        
        # ���[�N�u�b�N�����
        $workbook.Close()
        
        # COM�I�u�W�F�N�g�̉��
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
} catch {
    Write-Host "An error occurred: $_"
} finally {
    # Excel�A�v���P�[�V�������I��
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel application closed."
}