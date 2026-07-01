# Compress prebuilt exe into a zip under 100MB for GitHub push
# Usage: run from repository root

$sourceDir = Join-Path $PSScriptRoot "..\prebuilt\windows-x64" | Resolve-Path -Relative -ErrorAction SilentlyContinue
if (-not $sourceDir) {
	$sourceDir = "prebuilt/windows-x64"
}
$exePath = Join-Path $sourceDir 'ExcelTrainingMonitor.exe'
$zipPath = Join-Path $sourceDir 'ExcelTrainingMonitor.zip'

if (-not (Test-Path $exePath)) {
	Write-Host "No exe found at $exePath. Nothing to compress." -ForegroundColor Yellow
	exit 0
}

Write-Host "Found exe: $exePath"

# Remove existing zip if present
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

# Create zip using Compress-Archive (optimal compression)
try {
	Compress-Archive -LiteralPath $exePath -DestinationPath $zipPath -CompressionLevel Optimal -Force
} catch {
	Write-Error "Compress-Archive failed: $_"
	exit 1
}

$zipSize = (Get-Item $zipPath).Length
Write-Host "Created zip: $zipPath ($([math]::Round($zipSize/1MB,2)) MB)"

if ($zipSize -le 100MB) {
	Write-Host "Zip is under 100 MB — good to push." -ForegroundColor Green
	exit 0
}

Write-Warning "Zip is larger than 100 MB. Attempting executable compression with UPX (if available) and recompress."

# Try UPX if available to shrink the exe and recompress
$upx = "upx"
$upxFound = $false
try {
	$null = & $upx --version 2>$null
	$upxFound = $LASTEXITCODE -eq 0 -or $true
} catch {
	$upxFound = $false
}

if ($upxFound) {
	Write-Host "UPX found in PATH. Running upx -9 on exe..."
	$backupExe = "$exePath.orig"
	Copy-Item $exePath $backupExe -Force
	try {
		& $upx -9 --best --no-backup $exePath
		$upxExit = $LASTEXITCODE
	} catch {
		Write-Warning "UPX run failed: $_"
		$upxExit = 1
	}

	if ($upxExit -eq 0) {
		Write-Host "UPX compression finished, recreating zip..."
		Remove-Item $zipPath -Force
		Compress-Archive -LiteralPath $exePath -DestinationPath $zipPath -CompressionLevel Optimal -Force
		$zipSize = (Get-Item $zipPath).Length
		Write-Host "Recreated zip: $zipPath ($([math]::Round($zipSize/1MB,2)) MB)"
		if ($zipSize -le 100MB) {
			Write-Host "Zip is under 100 MB after UPX compression — good to push." -ForegroundColor Green
			exit 0
		}
		Write-Warning "Still larger than 100 MB after UPX. Restoring original exe." 
		if (Test-Path $backupExe) { Move-Item -Force $backupExe $exePath }
	} else {
		Write-Warning "UPX failed. Restoring original exe." 
		if (Test-Path $backupExe) { Move-Item -Force $backupExe $exePath }
	}
} else {
	Write-Warning "UPX not found in PATH. Install UPX (https://upx.github.io/) to attempt further exe compression, or reduce build size manually."
}

Write-Error "Unable to compress prebuilt to under 100 MB. Zip size: $([math]::Round($zipSize/1MB,2)) MB.\nOptions:\n - Install UPX and retry (tools may auto-compress)\n - Use Git LFS\n - Reduce build size (strip symbols, build Release, remove embedded resources)"
exit 2
