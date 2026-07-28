# ==============================================================================
# WIN 11 FILTER, AUTOPRINT & ARCHIVE PIPELINE (v1.2.1 PRODUCTION REFINED)
# REPOSITORY: https://github.com/skuple/autoprint
# ==============================================================================

[CmdletBinding()]
param (
    [switch]$Headless,
    [switch]$Stop,
    [string]$WatchFolder = ""
)

# ------------------------------------------------------------------------------
# DYNAMIC EXECUTABLE PATH RESOLVER
# ------------------------------------------------------------------------------
function Resolve-ExecutablePath ([string]$exeName, [array]$fallbackPaths) {
    $cmd = Get-Command $exeName -ErrorAction SilentlyContinue
    if ($cmd -and (Test-Path $cmd.Source)) { return $cmd.Source }

    $regKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$exeName",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$exeName",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$exeName"
    )
    foreach ($key in $regKeys) {
        if (Test-Path $key) {
            $regPath = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'(default)'
            if ($regPath -and (Test-Path $regPath)) { return $regPath }
        }
    }

    foreach ($path in $fallbackPaths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) { return $path }
    }
    return $null
}

# ------------------------------------------------------------------------------
# INTERACTIVE USER PROFILE RESOLVER
# ------------------------------------------------------------------------------
function Get-RealUserProfilePath {
    try {
        $explorerProc = Get-CimInstance -ClassName Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($explorerProc) {
            $owner = Invoke-CimMethod -InputObject $explorerProc -MethodName GetOwner -ErrorAction SilentlyContinue
            if ($owner -and $owner.User) {
                $userDir = Join-Path "C:\Users" $owner.User
                if (Test-Path $userDir) { return $userDir }
            }
        }
    } catch {}
    return $env:USERPROFILE
}

# Resolve Watch Folder
if ([string]::IsNullOrWhiteSpace($WatchFolder)) {
    $realUserDir = Get-RealUserProfilePath
    $folderToWatch = Join-Path $realUserDir "Downloads\LocalSend"
} else {
    $folderToWatch = $WatchFolder
}

$logFile    = Join-Path $folderToWatch "watcher_log.txt"
$configFile = Join-Path $folderToWatch "watcher_config.json"
$jobName    = "WindowsFolderWatcherJob"

# --- FEATURE: PROCESS TERMINATION (-Stop) ---
if ($Stop) {
    Write-Host "Stopping all watcher engines..." -ForegroundColor Yellow
    Get-Job -Name $jobName -ErrorAction SilentlyContinue | Stop-Job -ErrorAction SilentlyContinue
    Get-Job -Name $jobName -ErrorAction SilentlyContinue | Remove-Job -Force -ErrorAction SilentlyContinue
    (Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%Stable_V1.2.1.ps1%'").ProcessId | 
        Where-Object { $_ -ne $PID } | Stop-Process -Force -ErrorAction SilentlyContinue
    Write-Host "SUCCESS: Watcher processes stopped." -ForegroundColor Green
    exit
}

# ------------------------------------------------------------------------------
# CORE PIPELINE ENGINE (THREAD SAFE)
# ------------------------------------------------------------------------------
function Invoke-PipelineWorker ($path, $log, $configPath) {
    $ErrorActionPreference = "SilentlyContinue"
    $wrongFormatFolder = Join-Path $path "wrongFormat"
    $donePrintFolder   = Join-Path $path "donePrint"
    $retryPrintFolder  = Join-Path $path "retryPrint"
    $lastRetryCheck    = [datetime]::MinValue

    function Get-ElapsedText ($sw) { return "[ +$("{0:N2}" -f $sw.Elapsed.TotalSeconds)s ]" }

    # SYSTEM-SAFE PRINTER RESOLVER
    function Get-TargetPrinter {
        # 1. Try Default Printer
        $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Default = true" -ErrorAction SilentlyContinue
        if ($printer) { return $printer }

        # 2. Fallback: First Online/Connected Printer
        $fallback = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue | 
                    Where-Object { $_.WorkOffline -ne $true -and $_.PrinterStatus -ne 7 } | 
                    Select-Object -First 1
        return $fallback
    }

    # STARTUP HEALTH AUDIT
    function Invoke-PrinterStartupAudit ($logPath) {
        try {
            $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
            Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
            Add-Content -Path $logPath -Value "[ $timeStamp ] SYSTEM INITIALIZING: Running Startup Printer Health Audit..."

            $printer = Get-TargetPrinter

            if (-not $printer) {
                Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER AUDIT CRITICAL: NO ACTIVE PRINTER DETECTED ON THIS SYSTEM!"
                Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
                return $false
            }

            $portName   = $printer.PortName
            $driverName = $printer.DriverName
            $isOffline  = ($printer.WorkOffline -eq $true) -or ($printer.PrinterStatus -eq 7)
            $errState   = $printer.DetectedErrorState

            Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER AUDIT: Target Printer -> '$($printer.Name)'"
            Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER TELEMETRY: Port: $portName | Driver: $driverName | OfflineState: $isOffline"

            $stuckJobs = Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue
            $jobCount  = if ($stuckJobs) { ($stuckJobs | Measure-Object).Count } else { 0 }
            
            if ($jobCount -gt 0) {
                Add-Content -Path $logPath -Value "[ $timeStamp ] SPOOLER WARNING: Purging $jobCount stale job(s)..."
                $stuckJobs | Remove-PrintJob -ErrorAction SilentlyContinue
            } else {
                Add-Content -Path $logPath -Value "[ $timeStamp ] SPOOLER QUEUE: Clean (0 lingering jobs)."
            }

            if ($isOffline) {
                Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER HEALTH WARNING: Printer OFFLINE/UNPLUGGED."
            } else {
                Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER HEALTH PASSED: Hardware ready."
            }
            
            Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
            return $true
        } catch {
            return $false
        }
    }

    # SAFE PDF PAGE COUNTER (MEMORY GUARDED)
    function Get-PdfPageCount ($filePath) {
        try {
            [void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
            $asyncOp = [Windows.Storage.StorageFile]::GetFileFromPathAsync($filePath)
            while ($asyncOp.Status -eq "Started") { Start-Sleep -Milliseconds 20 }
            $storageFile = $asyncOp.GetResults()

            $docOp = [Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)
            while ($docOp.Status -eq "Started") { Start-Sleep -Milliseconds 20 }
            $pdfDoc = $docOp.GetResults()

            if ($pdfDoc -and $pdfDoc.PageCount) { return $pdfDoc.PageCount }
        } catch {}

        # Stream-Bounded Fallback (Reads max 5MB chunk to prevent RAM spikes)
        try {
            $stream = [System.IO.File]::OpenRead($filePath)
            $maxRead = [math]::Min($stream.Length, 5MB)
            $buffer = New-Object byte[] $maxRead
            [void]$stream.Read($buffer, 0, $maxRead)
            $stream.Close()

            $text = [System.Text.Encoding]::ASCII.GetString($buffer)
            $matches = [regex]::Matches($text, "/Type\s*/Page\b")
            if ($matches.Count -gt 0) { return $matches.Count }
        } catch {}

        return 1
    }

    function Get-PrinterHealthStatus {
        try {
            $printer = Get-TargetPrinter
            if (-not $printer) {
                return @{ IsReady = $false; PrinterName = ""; ErrorCode = -1; StateText = "NO_PRINTER_AVAILABLE" }
            }

            $errState = $printer.DetectedErrorState
            $isOffline = ($printer.WorkOffline -eq $true) -or ($printer.PrinterStatus -eq 7)

            $stateText = "NORMAL"
            if ($errState -eq 5) { $stateText = "PAPER_OUT" }
            elseif ($errState -eq 4) { $stateText = "PAPER_JAM" }
            elseif ($errState -ne 0) { $stateText = "HARDWARE_FAULT" }
            elseif ($isOffline) { $stateText = "OFFLINE" }

            return @{
                IsReady     = (-not $isOffline) -and ($errState -eq 0)
                PrinterName = $printer.Name
                ErrorCode   = $errState
                StateText   = $stateText
            }
        } catch {
            return @{ IsReady = $false; PrinterName = ""; ErrorCode = -1; StateText = "WMI_ERROR" }
        }
    }

    function Convert-DocToPdf ($docPath, $outputFolder, $soPath) {
        try {
            if (-not $soPath -or -not (Test-Path $soPath)) { return $null }
            
            $pInfo = New-Object System.Diagnostics.ProcessStartInfo
            $pInfo.FileName = $soPath
            $pInfo.Arguments = "--headless --convert-to pdf `"$docPath`" --outdir `"$outputFolder`""
            $pInfo.CreateNoWindow = $true
            $pInfo.UseShellExecute = $false

            $proc = [System.Diagnostics.Process]::Start($pInfo)
            if ($proc) {
                if (-not $proc.WaitForExit(30000)) {
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                    return $null
                }
            }

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($docPath)
            $expectedPdf = Join-Path $outputFolder "${baseName}.pdf"
            if (Test-Path $expectedPdf) { return $expectedPdf }
        } catch {}
        return $null
    }

    function Test-PdfIntegrity ($filePath) {
        $fileStream = $null
        try {
            if (-not (Test-Path $filePath)) { return @{ Valid = $false; Reason = "FILE_NOT_FOUND" } }
            $fileStream = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
            $buffer = New-Object byte[] 2048
            $bytesRead = $fileStream.Read($buffer, 0, 2048)
            if ($bytesRead -lt 10) { return @{ Valid = $false; Reason = "EMPTY_FILE" } }

            $text = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
            if ($text -notlike "*%PDF-*") { return @{ Valid = $false; Reason = "NOT_A_PDF" } }
            if ($text -like "*/Encrypt*") { return @{ Valid = $false; Reason = "ENCRYPTED" } }

            return @{ Valid = $true; Reason = "OK" }
        } catch {
            return @{ Valid = $false; Reason = "READ_ERROR" }
        } finally {
            if ($fileStream) { $fileStream.Close(); $fileStream.Dispose() }
        }
    }

    function Move-ToFolder ($currentPath, $cleanOriginalName, $targetFolder) {
        try {
            if (-not (Test-Path $targetFolder)) { New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null }
            $ext = [System.IO.Path]::GetExtension($cleanOriginalName)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($cleanOriginalName)
            
            $destPath = Join-Path $targetFolder "${base}${ext}"
            if (Test-Path $destPath) {
                $destPath = Join-Path $targetFolder "${base}_$((Get-Date).ToString('yyyyMMdd_HHmmss'))${ext}"
            }

            Move-Item -Path $currentPath -Destination $destPath -Force -ErrorAction Stop
            return $true
        } catch { return $false }
    }

    # RUN STARTUP AUDIT
    Invoke-PrinterStartupAudit -logPath $log | Out-Null

    # MAIN POLLING LOOP
    while ($true) {
        try {
            # Dynamic Binary Re-Resolution
            $currentLibre = Resolve-ExecutablePath -exeName "soffice.exe" -fallbackPaths @("C:\Program Files\LibreOffice\program\soffice.exe")
            $currentSumatra = Resolve-ExecutablePath -exeName "SumatraPDF.exe" -fallbackPaths @("C:\Program Files\SumatraPDF\SumatraPDF.exe")

            $ignoredFolders = @("wrongFormat", "donePrint", "retryPrint")
            $ignoredFiles   = @((Split-Path $log -Leaf), "watcher_config.json")

            $items = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | 
                     Where-Object { $_.Name -notin $ignoredFiles -and $_.Directory.Name -notin $ignoredFolders -and $_.Name -notlike "~$*" } | 
                     Sort-Object CreationTime

            foreach ($item in $items) {
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $currentPath = $item.FullName
                $cleanOriginalName = $item.Name
                $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
                $extension = [System.IO.Path]::GetExtension($cleanOriginalName).ToLower()

                if ($extension -notin @(".doc", ".docx", ".xls", ".xlsx", ".pdf")) {
                    Move-ToFolder $currentPath $cleanOriginalName $wrongFormatFolder | Out-Null
                    continue
                }

                $pdfToPrintPath = $currentPath

                # Office Conversion
                if ($extension -in @(".doc", ".docx", ".xls", ".xlsx")) {
                    $convertedPdf = Convert-DocToPdf -docPath $currentPath -outputFolder $path -soPath $currentLibre
                    if (-not $convertedPdf) {
                        Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                        continue
                    }
                    $pdfToPrintPath = $convertedPdf
                }

                # Integrity Test
                $integrity = Test-PdfIntegrity $pdfToPrintPath
                if (-not $integrity.Valid) {
                    Move-ToFolder $currentPath $cleanOriginalName $wrongFormatFolder | Out-Null
                    continue
                }

                # Health & Dispatch
                $health = Get-PrinterHealthStatus
                if ($health.IsReady) {
                    if ($currentSumatra) {
                        $pInfo = New-Object System.Diagnostics.ProcessStartInfo
                        $pInfo.FileName = $currentSumatra
                        $pInfo.Arguments = "-silent -print-settings `"simplex,fit`" -print-to `"$($health.PrinterName)`" `"$pdfToPrintPath`""
                        $pInfo.CreateNoWindow = $true
                        $pInfo.UseShellExecute = $false
                        $p = [System.Diagnostics.Process]::Start($pInfo)
                        if ($p) { $p.WaitForExit(20000); $p.Close() }
                    } else {
                        Start-Process -FilePath $pdfToPrintPath -Verb Print -WindowStyle Hidden -ErrorAction SilentlyContinue
                    }

                    Move-ToFolder $currentPath $cleanOriginalName $donePrintFolder | Out-Null
                    Add-Content -Path $log -Value "[ $timeStamp ] PRINT SUCCESS: $cleanOriginalName printed to $($health.PrinterName)"
                } else {
                    Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                }
            }
        } catch {}

        [System.GC]::Collect()
        Start-Sleep -Seconds 2
    }
}

# ------------------------------------------------------------------------------
# HEADLESS VS INTERACTIVE EXECUTION
# ------------------------------------------------------------------------------
if ($Headless) {
    # DIRECT THREAD EXECUTION (Prevents Start-Job RAM Leak & Overhead)
    Invoke-PipelineWorker -path $folderToWatch -log $logFile -configPath $configFile
} else {
    Write-Host "Starting Interactive Watcher..." -ForegroundColor Cyan
    Start-Job -ScriptBlock {
        param($p, $l, $c)
        Invoke-PipelineWorker -path $p -log $l -configPath $c
    } -Name $jobName -ArgumentList $folderToWatch, $logFile, $configFile | Out-Null
}