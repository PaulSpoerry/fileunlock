# List of target files
$targetFiles = @(
    "C:\source\repos\DSSOrders\node_modules\@swc\core-win32-x64-msvc\swc.win32-x64-msvc.node",
    "C:\source\repos\DSSOrders\node_modules\some-other\locked.file"
)

# Fixed path to handle.exe
$handlePath = "C:\source\repos\fileunlock\handle\handle.exe"

if (-not (Test-Path $handlePath)) {
    Write-Error "handle.exe not found at $handlePath"
    exit 1
}

# Accept EULA if needed
& $handlePath > $null

foreach ($targetFile in $targetFiles) {
    Write-Host "`nChecking file: $targetFile"

    # Run handle.exe and get results
    $output = & $handlePath -accepteula $targetFile 2>&1
    $lockedLines = $output | Where-Object { $_ -match 'pid: (\d+)' -and $_ -match '([0-9A-F]+): File' }

    if (-not $lockedLines) {
        Write-Host "  No process is locking this file."
        continue
    }

    foreach ($line in $lockedLines) {
        if ($line -match 'pid: (\d+).*?([0-9A-F]+): File') {
            $pid = $matches[1]
            $handleID = $matches[2]

            Write-Host "  Found handle $handleID on PID $pid. Attempting to close..."

            try {
                $closeOutput = & $handlePath -accepteula -c $handleID -p $pid 2>&1

                if ($closeOutput -match "Handle closed") {
                    Write-Host "    [✔] Successfully closed handle $handleID on PID $pid"
                } else {
                    Write-Warning "    [!] Failed to close handle $handleID — attempting to kill PID $pid"
                    try {
                        Stop-Process -Id $pid -Force -ErrorAction Stop
                        Write-Host "    [✖] Process $pid killed."
                    } catch {
                        Write-Warning '    [X] Failed to kill process ' + $pid + ': ' + $PSItem.Exception.Message
                    }
                }
            } catch {
                Write-Warning '    [X] Error while processing handle ' + $handleID + ' on PID ' + $pid + ': ' + $PSItem.Exception.Message
            }
        }
    }
}
