$scriptPath = Join-Path $env:LOCALAPPDATA "open_word_handler.ps1"

$scriptCode = @'
param([string]$rawUri)
if (-not $rawUri) { exit }

Add-Type -AssemblyName System.Windows.Forms

# 1. 회사 이름 디코딩 및 특수문자 제거
$clean = $rawUri -replace "^myword://", "" -replace "/$", ""
$company = [System.Uri]::UnescapeDataString($clean).Trim()
foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
    $company = $company.Replace($ch.ToString(), "")
}
if ([string]::IsNullOrWhiteSpace($company)) { exit }

# 2. 실제 존재하는 '자소서 모음' 폴더 경로 탐색
$candidates = @(
    "$env:USERPROFILE\OneDrive\바탕 화면\자소서 모음",
    "$env:USERPROFILE\OneDrive\Desktop\자소서 모음",
    "$env:USERPROFILE\바탕 화면\자소서 모음",
    "$env:USERPROFILE\Desktop\자소서 모음"
)

$targetDir = ""
foreach ($dir in $candidates) {
    if (Test-Path -LiteralPath $dir) {
        $targetDir = $dir
        break
    }
}

if (-not $targetDir) {
    $realDesktop = [System.Environment]::GetFolderPath('Desktop')
    $targetDir = Join-Path $realDesktop "자소서 모음"
    if (-not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }
}

$fileName = "${company}_자소서.docx"
$filePath = Join-Path $targetDir $fileName

# 3. 파일 존재 여부 확인 및 분기 처리
try {
    if (Test-Path -LiteralPath $filePath) {
        # 2번째 클릭부터: 기존 파일 Word로 즉시 열기
        Start-Process -FilePath $filePath
    } else {
        # 첫 클릭: 파일만 생성하고 안내창 출력 (탐색기 폴더 창 열림 기능 제거)
        $null = New-Item -ItemType File -Path $filePath -Force

        [System.Windows.Forms.MessageBox]::Show(
            "파일이 성공적으로 생성되었습니다!`n`n파일명: $fileName`n저장위치: $targetDir`n`n다음 클릭부터는 Word로 바로 열립니다.",
            "생성 완료",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "오류 내용: " + $_.Exception.Message,
        "오류",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}
'@

[System.IO.File]::WriteAllText($scriptPath, $scriptCode, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "[수정 완료] 자소서 모음 폴더 열림 기능이 제거되었습니다!" -ForegroundColor Green
Write-Host "이제 첫 클릭 시 폴더 창 없이 알림 메시지만 나타납니다." -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Green