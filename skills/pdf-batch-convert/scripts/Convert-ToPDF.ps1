param([string[]]$InputFiles, [switch]$NoUI)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
if (-not $InputFiles) {
    if ($NoUI) { throw '변환할 파일 경로가 필요합니다.' }
    $picker = New-Object System.Windows.Forms.OpenFileDialog
    $picker.Title = 'PDF로 변환할 문서 선택'
    $picker.Filter = '한글 및 Word|*.hwp;*.hwpx;*.doc;*.docx'
    $picker.Multiselect = $true
    if ($picker.ShowDialog() -ne 'OK') { $picker.Dispose(); exit }
    $InputFiles = $picker.FileNames
    $picker.Dispose()
}
function Release-Com($obj) {
    if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
    }
}
$results = New-Object 'System.Collections.Generic.List[object]'
foreach ($source in ($InputFiles | Select-Object -Unique)) {
    $app = $null; $doc = $null; $temp = $null; $target = ''; $ext = ''
    try {
        $file = Get-Item -LiteralPath $source
        $ext = $file.Extension.ToLowerInvariant()
        if ($file.PSIsContainer -or $ext -notin '.hwp','.hwpx','.doc','.docx') { throw '지원하지 않는 형식입니다.' }
        if ($file.Name.StartsWith('~$')) { throw '임시 파일은 제외하세요.' }
        $folder = Join-Path $file.DirectoryName 'PDF 결과'
        [void][IO.Directory]::CreateDirectory($folder)
        $target = Join-Path $folder ($file.Name + '.pdf')
        if (Test-Path -LiteralPath $target) {
            $results.Add([pscustomobject]@{File=$source;Status='건너뜀';PDF=$target;Detail='기존 PDF 있음'})
            continue
        }
        $temp = Join-Path $folder ('.converting-' + [guid]::NewGuid().ToString('N') + '.pdf')
        Write-Host "변환 중: $source"
        if ($ext -in '.doc','.docx') {
            $app = New-Object -ComObject Word.Application
            $app.Visible = $false
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = 0
            $doc = $app.Documents.Open($file.FullName, $false, $true, $false)
            $doc.ExportAsFixedFormat($temp, 17)
        } else {
            $app = New-Object -ComObject HWPFrame.HwpObject
            if (-not $app.Open($file.FullName, '', '')) { throw '한글 문서 열기 실패' }
            if (-not $app.SaveAs($temp, 'PDF', '')) { throw '한글 PDF 저장 실패' }
        }
        $stream = [IO.File]::OpenRead($temp)
        try {
            $header = New-Object byte[] 5
            if ($stream.Read($header,0,5) -ne 5 -or [Text.Encoding]::ASCII.GetString($header) -ne '%PDF-') { throw 'PDF 형식 확인 실패' }
        } finally { $stream.Dispose() }
        [IO.File]::Move($temp, $target)
        $results.Add([pscustomobject]@{File=$source;Status='성공';PDF=$target;Detail=''})
    } catch {
        $results.Add([pscustomobject]@{File=$source;Status='실패';PDF=$target;Detail=$_.Exception.Message})
    } finally {
        if ($doc) { try { $doc.Close(0) } catch {}; Release-Com $doc }
        if ($app) {
            try {
                if ($ext -in '.hwp','.hwpx') { $app.Clear(1); $app.Quit() }
                else { $app.Quit(0) }
            } catch {}
            Release-Com $app
        }
        if ($temp -and (Test-Path -LiteralPath $temp)) { Remove-Item -LiteralPath $temp -ErrorAction SilentlyContinue }
    }
}
$logs = Join-Path $PSScriptRoot '변환 기록'
[void][IO.Directory]::CreateDirectory($logs)
$log = Join-Path $logs ((Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N') + '.csv')
$results | Export-Csv -LiteralPath $log -NoTypeInformation -Encoding UTF8
$ok = @($results | Where-Object Status -eq '성공').Count
$skip = @($results | Where-Object Status -eq '건너뜀').Count
$fail = @($results | Where-Object Status -eq '실패').Count
$summary = "성공 $ok / 건너뜀 $skip / 실패 $fail`r`nPDF: 원본 폴더의 PDF 결과`r`n기록: $log"
$results | Format-Table -AutoSize
Write-Host $summary
if (-not $NoUI) { [void][System.Windows.Forms.MessageBox]::Show($summary, 'PDF 변환 결과') }
if ($fail -gt 0) { exit 1 }
