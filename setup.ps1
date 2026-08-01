# fsearch를 Windows PATH에 추가하는 스크립트
# 관리자 권한으로 실행해야 합니다.

param(
    [switch]$AllUsers = $false
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$fsearchPath = Join-Path $scriptPath "fsearch.bat"

if (-not (Test-Path $fsearchPath)) {
    Write-Error "fsearch.bat를 찾을 수 없습니다: $fsearchPath"
    exit 1
}

# 현재 PATH 가져오기
if ($AllUsers) {
    $currentPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $scope = "Machine"
    Write-Host "🔧 시스템 전체 PATH에 추가합니다 (모든 사용자)" -ForegroundColor Cyan
} else {
    $currentPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $scope = "User"
    Write-Host "🔧 사용자 PATH에 추가합니다" -ForegroundColor Cyan
}

# 이미 추가되어 있는지 확인
if ($currentPath -like "*$scriptPath*") {
    Write-Host "✅ 이미 PATH에 등록되어 있습니다!" -ForegroundColor Green
    exit 0
}

# PATH에 추가
$newPath = $currentPath + ";" + $scriptPath
[System.Environment]::SetEnvironmentVariable("Path", $newPath, $scope)

Write-Host "✅ PATH에 등록되었습니다!" -ForegroundColor Green
Write-Host "📁 경로: $scriptPath" -ForegroundColor Yellow
Write-Host ""
Write-Host "이제 어디서나 다음 명령으로 사용 가능합니다:" -ForegroundColor Green
Write-Host "  fsearch '검색어'" -ForegroundColor Cyan
Write-Host ""
Write-Host "새 PowerShell/CMD 창을 열어서 사용하세요." -ForegroundColor Yellow
