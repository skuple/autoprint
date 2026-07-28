# ==============================================================================
# WIN 11 FILTER, AUTOPRINT & ARCHIVE PIPELINE
# EXPERIMENTAL BRANCH V1.2.1-AUDITTEST: SPREADSHEET SUPPORT & STARTUP AUDIT
# ==============================================================================

[CmdletBinding()]
param (
    [switch]$Headless,
    [switch]$Stop
)

# ------------------------------------------------------------------------------
# 1. PIPELINE CONFIGURATION & PATHS
# ------------------------------------------------------------------------------
$folderToWatch   = "C:\Users\Administrator\Downloads\LocalSend"
$logFile          = "C:\Users\Administrator\Downloads\LocalSend\watcher_log.txt"
$jobName          = "WindowsFolderWatcherJob"
$configFile       = "C:\Users\Administrator\Downloads\LocalSend\watcher_config.json"
$sumatraPath      = "C:\Users\Administrator\AppData\Local\SumatraPDF\SumatraPDF.exe"
$libreOfficePath  = "C:\Program Files\LibreOffice\program\soffice.exe"

# Fallback if LibreOffice is installed in 32-bit directory
if (-not (Test-Path $libreOfficePath)) {
    $libreOfficePath = "C:\Program Files (x86)\LibreOffice\program\soffice.exe"
}

# --- FEATURE: INSTANT PROCESS TERMINATION (-Stop) ---
if ($Stop) {
    Write-Host "Stopping all running watcher processes..." -ForegroundColor Yellow
    
    # Stop background jobs
    Get-Job -Name $jobName -ErrorAction SilentlyContinue | Stop-Job -ErrorAction SilentlyContinue
    Get-Job -Name $jobName -ErrorAction SilentlyContinue | Remove-Job -Force -ErrorAction SilentlyContinue
    
    # Kill any orphaned headless process handles
    (Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%Stable_V1_DynLibre.ps1%' OR CommandLine LIKE '%automaticPrintScript.ps1%'").ProcessId | 
        Where-Object { $_ -ne $PID } | 
        Stop-Process -Force -ErrorAction SilentlyContinue

    Write-Host "SUCCESS: Watcher process engine stopped." -ForegroundColor Green
    exit
}

# Pre-run Memory Cleanup: Terminate lingering background jobs
$lingeringJobs = Get-Job -Name $jobName -ErrorAction SilentlyContinue
if ($lingeringJobs) {
    Stop-Job -Name $jobName -ErrorAction SilentlyContinue
    Remove-Job -Name $jobName -Force -ErrorAction SilentlyContinue
}

# ------------------------------------------------------------------------------
# 2. CONFIGURATION MANAGEMENT (JSON-BASED)
# ------------------------------------------------------------------------------
function Get-PipelineConfig {
    param([string]$path)
    $defaultConfig = @{
        AutoPrint            = $true
        PollingIntervalSec   = 1.5
        MaxSpoolWaitSec      = 300
        LogMaxSizeMB         = 2
        MaxLogArchivesToKeep = 5
        SelfHealingMinutes   = 10
    }

    if (Test-Path $path) {
        try {
            $json = Get-Content -Path $path -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
            if ($json) { return $json }
        } catch {}
    }

    $defaultConfig | ConvertTo-Json | Set-Content -Path $path -Force
    return [PSCustomObject]$defaultConfig
}

$globalConfig = Get-PipelineConfig -path $configFile

function Save-PipelineConfig {
    param([string]$path, $configObject)
    try {
        $configObject | ConvertTo-Json | Set-Content -Path $path -Force
    } catch {}
}

function Get-DefaultPrinterName {
    try {
        $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Default = true" -ErrorAction SilentlyContinue
        if ($printer) { return $printer.Name }
    } catch {}
    return "No Default Printer Found"
}

function Get-WatcherStatus {
    $job = Get-Job -Name $jobName -ErrorAction SilentlyContinue
    if ($job -and $job.State -eq "Running") { return "RUNNING" }
    return "STOPPED"
}

# ------------------------------------------------------------------------------
# 3. BACKGROUND WATCHER ENGINE
# ------------------------------------------------------------------------------
function Start-Watcher {
    $existingJob = Get-Job -Name $jobName -ErrorAction SilentlyContinue
    if ($existingJob) {
        if ($existingJob.State -eq "Running") {
            Write-Host "Watcher engine is already running!" -ForegroundColor Yellow
            return
        } else {
            Remove-Job -Name $jobName -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Write-Host "Starting enterprise folder gatekeeper (v1.2.1-AuditTest)..." -ForegroundColor Cyan

    $scriptBlock = {
        param($path, $log, $configPath, $sumatra, $libreOffice)
        
        $ErrorActionPreference = "SilentlyContinue"
        $wrongFormatFolder = Join-Path $path "wrongFormat"
        $donePrintFolder   = Join-Path $path "donePrint"
        $retryPrintFolder  = Join-Path $path "retryPrint"
        $lastRetryCheck    = [datetime]::MinValue

        function Get-ElapsedText ($sw) { return "[ +$("{0:N2}" -f $sw.Elapsed.TotalSeconds)s ]" }

        # --- STARTUP PRINTER HEALTH AUDIT ---
        function Invoke-PrinterStartupAudit ($logPath) {
            try {
                $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
                Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
                Add-Content -Path $logPath -Value "[ $timeStamp ] SYSTEM INITIALIZING: Running Startup Printer Health Audit..."

                $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Default = true" -ErrorAction SilentlyContinue

                if (-not $printer) {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER AUDIT CRITICAL: NO DEFAULT PRINTER CONFIGURED ON THIS MACHINE!"
                    Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
                    return $false
                }

                $portName   = $printer.PortName
                $driverName = $printer.DriverName
                $isOffline  = ($printer.WorkOffline -eq $true) -or ($printer.PrinterStatus -eq 7)
                $errState   = $printer.DetectedErrorState

                Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER AUDIT: Target Printer -> '$($printer.Name)'"
                Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER TELEMETRY: Port: $portName | Driver: $driverName | OfflineState: $isOffline"

                # Check existing queue
                $stuckJobs = Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue
                $jobCount  = if ($stuckJobs) { ($stuckJobs | Measure-Object).Count } else { 0 }
                
                if ($jobCount -gt 0) {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] SPOOLER WARNING: Found $jobCount lingering job(s) in queue. Auto-purging stale jobs..."
                    $stuckJobs | Remove-PrintJob -ErrorAction SilentlyContinue
                } else {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] SPOOLER QUEUE: Clean (0 lingering jobs)."
                }

                if ($isOffline) {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER HEALTH WARNING: Printer is OFFLINE/UNPLUGGED. Waiting for hardware..."
                } elseif ($errState -eq 5) {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER HEALTH WARNING: Paper Out detected on startup."
                } elseif ($errState -eq 4) {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER HEALTH WARNING: Paper Jam detected on startup."
                } else {
                    Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER HEALTH PASSED: Hardware ready to accept print jobs."
                }
                
                Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
                return $true
            } catch {
                $timeStamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
                Add-Content -Path $logPath -Value "[ $timeStamp ] PRINTER AUDIT ERROR: Failed to query WMI printer subsystem."
                Add-Content -Path $logPath -Value "[ $timeStamp ] =================================================="
                return $false
            }
        }

        # --- NATIVE WIN11 PDF PAGE COUNTER WITH REGEX FALLBACK ---
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

            # Fast Regex Fallback for standard PDF page objects
            try {
                $stream = [System.IO.File]::OpenRead($filePath)
                $reader = New-Object System.IO.StreamReader($stream)
                $content = $reader.ReadToEnd()
                $reader.Close()
                $stream.Close()
                $matches = [regex]::Matches($content, "/Type\s*/Page\b")
                if ($matches.Count -gt 0) { return $matches.Count }
            } catch {}

            return 1 # Fallback to 1 page if undetected
        }

        # --- PRINTER HEALTH SENSOR ---
        function Get-PrinterHealthStatus {
            try {
                $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Default = true" -ErrorAction SilentlyContinue
                if (-not $printer) {
                    return @{ IsReady = $false; PrinterName = ""; ErrorCode = -1; StateText = "NO_DEFAULT_PRINTER" }
                }

                $errState = $printer.DetectedErrorState
                $isOffline = ($printer.WorkOffline -eq $true) -or ($printer.PrinterStatus -eq 7)

                $stateText = "NORMAL"
                if ($errState -eq 5) { $stateText = "PAPER_OUT" }
                elseif ($errState -eq 4) { $stateText = "PAPER_JAM" }
                elseif ($errState -ne 0) { $stateText = "HARDWARE_FAULT" }
                elseif ($isOffline) { $stateText = "OFFLINE" }

                $isReady = (-not $isOffline) -and ($errState -eq 0)

                return @{
                    IsReady     = $isReady
                    PrinterName = $printer.Name
                    ErrorCode   = $errState
                    StateText   = $stateText
                }
            } catch {
                return @{ IsReady = $false; PrinterName = ""; ErrorCode = -1; StateText = "WMI_ERROR" }
            }
        }

        # --- LIBREOFFICE DOC/DOCX/XLS/XLSX CONVERTER (PID ISOLATED) ---
        function Convert-DocToPdf ($docPath, $outputFolder, $soPath) {
            try {
                if (-not (Test-Path $soPath)) { return $null }
                
                $pInfo = New-Object System.Diagnostics.ProcessStartInfo
                $pInfo.FileName = $soPath
                $pInfo.Arguments = "--headless --convert-to pdf `"$docPath`" --outdir `"$outputFolder`""
                $pInfo.CreateNoWindow = $true
                $pInfo.UseShellExecute = $false

                $proc = [System.Diagnostics.Process]::Start($pInfo)
                if ($proc) {
                    $targetPid = $proc.Id
                    $exited = $proc.WaitForExit(30000) # 30s process limit
                    if (-not $exited) {
                        Stop-Process -Id $targetPid -Force -ErrorAction SilentlyContinue
                        return $null
                    }
                    $proc.Close()
                }

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($docPath)
                $expectedPdf = Join-Path $outputFolder "${baseName}.pdf"

                if (Test-Path $expectedPdf) { return $expectedPdf }
            } catch {}
            return $null
        }

        # --- PDF PRE-FLIGHT INSPECTION ---
        function Test-PdfIntegrity ($filePath) {
            $fileStream = $null
            try {
                if (-not (Test-Path $filePath)) { return @{ Valid = $false; Reason = "FILE_NOT_FOUND" } }
                $fileStream = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
                $buffer = New-Object byte[] 2048
                $bytesRead = $fileStream.Read($buffer, 0, 2048)

                if ($bytesRead -lt 10) { return @{ Valid = $false; Reason = "EMPTY_OR_TRUNCATED_FILE" } }

                $text = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)

                if ($text -notlike "*%PDF-*") { return @{ Valid = $false; Reason = "CORRUPT_HEADER_NOT_PDF" } }
                if ($text -like "*/Encrypt*") { return @{ Valid = $false; Reason = "PASSWORD_PROTECTED_ENCRYPTED" } }

                return @{ Valid = $true; Reason = "OK" }
            } catch {
                return @{ Valid = $false; Reason = "FILE_READ_ERROR" }
            } finally {
                if ($fileStream) { $fileStream.Close(); $fileStream.Dispose() }
            }
        }

        # --- LOG ROTATION & RETENTION ---
        function Invoke-LogRotation ($logPath, $targetFolder, $maxSizeMB, $maxArchivesToKeep) {
            try {
                if (Test-Path $logPath) {
                    $logItem = Get-Item $logPath
                    if ($logItem.Length -gt ($maxSizeMB * 1MB)) {
                        $archiveName = "watcher_log_archive_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
                        $archivePath = Join-Path $targetFolder $archiveName
                        Move-Item -Path $logPath -Destination $archivePath -Force -ErrorAction SilentlyContinue

                        Get-ChildItem -Path $targetFolder -Filter "watcher_log_archive_*.txt" -ErrorAction SilentlyContinue | 
                            Sort-Object CreationTime -Descending | 
                            Select-Object -Skip $maxArchivesToKeep | 
                            Remove-Item -Force -ErrorAction SilentlyContinue
                    }
                }
            } catch {}
        }

        # --- FILENAME CLEANER & STATE TAGGER ---
        function Get-CleanFileName ($rawName) {
            $ext = [System.IO.Path]::GetExtension($rawName)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($rawName)
            $cleanBase = $base -replace '(_PROCESSING|_WAITING_PAPER|_PAPER_JAMMED|_HARDWARE_ERROR|_RETRY_QUEUE)$', ''
            return "${cleanBase}${ext}"
        }

        function Update-FileState ($currentPath, $cleanOriginalName, $stateTag) {
            try {
                if (-not (Test-Path $currentPath)) { return $currentPath }
                $parentDir = Split-Path $currentPath -Parent
                $ext = [System.IO.Path]::GetExtension($cleanOriginalName)
                $base = [System.IO.Path]::GetFileNameWithoutExtension($cleanOriginalName)
                
                $newName = if ($stateTag) { "${base}${stateTag}${ext}" } else { "${base}${ext}" }
                $newPath = Join-Path $parentDir $newName

                if ($currentPath -eq $newPath) { return $currentPath }

                if (Test-Path $currentPath) {
                    Rename-Item -Path $currentPath -NewName $newName -Force -ErrorAction Stop
                    return $newPath
                }
            } catch {}
            return $currentPath
        }

        # --- SAFE FILE MOVER WITH COLLISION GUARD ---
        function Move-ToFolder ($currentPath, $cleanOriginalName, $targetFolder) {
            try {
                if (-not (Test-Path $targetFolder)) {
                    New-Item -ItemType Directory -Path $targetFolder -ErrorAction SilentlyContinue | Out-Null
                }
                $ext = [System.IO.Path]::GetExtension($cleanOriginalName)
                $base = [System.IO.Path]::GetFileNameWithoutExtension($cleanOriginalName)
                
                $destPath = Join-Path $targetFolder "${base}${ext}"
                if (Test-Path $destPath) {
                    $timeTag = (Get-Date).ToString("yyyyMMdd_HHmmss")
                    $destPath = Join-Path $targetFolder "${base}_${timeTag}${ext}"
                }

                $moved = $false
                $mRetry = 0
                while (-not $moved -and $mRetry -lt 10) {
                    try {
                        Move-Item -Path $currentPath -Destination $destPath -Force -ErrorAction Stop
                        $moved = $true
                    } catch {
                        $mRetry++
                        Start-Sleep -Milliseconds 500
                    }
                }
                return $moved
            } catch {
                return $false
            }
        }

        # ----------------------------------------------------------------------
        # EXECUTE STARTUP AUDIT ONCE AT BOOT
        # ----------------------------------------------------------------------
        Invoke-PrinterStartupAudit -logPath $log | Out-Null

        # ----------------------------------------------------------------------
        # MAIN POLLING LOOP
        # ----------------------------------------------------------------------
        while ($true) {
            try {
                $cfg = Get-Content -Path $configPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
                if (-not $cfg) { $cfg = @{ AutoPrint = $true; PollingIntervalSec = 1.5; LogMaxSizeMB = 2; MaxLogArchivesToKeep = 5; SelfHealingMinutes = 10; MaxSpoolWaitSec = 300 } }

                # 1. Log Rotation Check
                Invoke-LogRotation -logPath $log -targetFolder $path -maxSizeMB $cfg.LogMaxSizeMB -maxArchivesToKeep $cfg.MaxLogArchivesToKeep

                $ignoredFolders = @("wrongFormat", "donePrint", "retryPrint")
                $ignoredFiles   = @((Split-Path $log -Leaf), "watcher_config.json", "watcher_config.xml")

                $items = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | 
                         Where-Object { 
                             $_.Name -notin $ignoredFiles -and 
                             $_.Directory.Name -notin $ignoredFolders -and
                             $_.Name -notlike "watcher_log_archive_*" -and
                             $_.Name -notlike "~$*" -and
                             $_.Name -notlike "*_CONVERTED_TEMP.pdf"
                         } | 
                         Sort-Object CreationTime

                # 2. Self-Healing retryPrint Scheduler
                if (($items.Count -eq 0) -and ((Get-Date) -gt $lastRetryCheck.AddMinutes($cfg.SelfHealingMinutes))) {
                    $lastRetryCheck = Get-Date
                    if (Test-Path $retryPrintFolder) {
                        $retryFiles = Get-ChildItem -Path $retryPrintFolder -File -ErrorAction SilentlyContinue | Sort-Object CreationTime
                        if ($retryFiles) {
                            $health = Get-PrinterHealthStatus
                            if ($health.IsReady) {
                                $fileToRequeue = $retryFiles[0]
                                $requeueDest = Join-Path $path $fileToRequeue.Name
                                if (-not (Test-Path $requeueDest)) {
                                    Move-Item -Path $fileToRequeue.FullName -Destination $requeueDest -Force -ErrorAction SilentlyContinue
                                    Add-Content -Path $log -Value "[ $((Get-Date).ToString('MM/dd/yyyy HH:mm:ss')) ] SELF-HEALING ENGINE: Re-queued '$($fileToRequeue.Name)' from retryPrint to LocalSend."
                                    
                                    $items = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | 
                                             Where-Object { $_.Name -notin $ignoredFiles -and $_.Directory.Name -notin $ignoredFolders } | 
                                             Sort-Object CreationTime
                                }
                            }
                        }
                    }
                }

                # 3. Main File Processing Loop
                foreach ($item in $items) {
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    $currentPath       = $item.FullName
                    $rawFileName       = $item.Name
                    $cleanOriginalName = Get-CleanFileName $rawFileName
                    $timeStamp         = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
                    $tempPdfToClean    = $null

                    Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) FILE DETECTED: $cleanOriginalName"

                    try {
                        # Lock Check
                        $fileReady = $false
                        $retryCount = 0
                        while (-not $fileReady -and $retryCount -lt 10) {
                            try {
                                $stream = [System.IO.File]::Open($currentPath, 'Open', 'ReadWrite', 'None')
                                $stream.Close()
                                $fileReady = $true
                            } catch {
                                $retryCount++
                                Start-Sleep -Milliseconds 500
                            }
                        }

                        if (-not (Test-Path $currentPath)) { continue }

                        $extension = [System.IO.Path]::GetExtension($cleanOriginalName).ToLower()
                        $allowedExtensions = @(".doc", ".docx", ".xls", ".xlsx", ".pdf")

                        # Extension Check
                        if ($extension -notin $allowedExtensions) {
                            Move-ToFolder $currentPath $cleanOriginalName $wrongFormatFolder | Out-Null
                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) BLOCKED & MOVED TO wrongFormat: $cleanOriginalName"
                            $sw.Stop()
                            continue
                        }

                        $pdfToPrintPath = $currentPath

                        # --- CONVERT OFFICE FILES (.DOC, .DOCX, .XLS, .XLSX) VIA LIBREOFFICE ---
                        if ($extension -in @(".doc", ".docx", ".xls", ".xlsx")) {
                            if (-not (Test-Path $libreOffice)) {
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) CONVERSION FAILED: LibreOffice not found at '$libreOffice'"
                                Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                                $sw.Stop()
                                continue
                            }

                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) CONVERTING OFFICE FILE: Passing $cleanOriginalName to LibreOffice..."
                            $convertedPdf = Convert-DocToPdf -docPath $currentPath -outputFolder $path -soPath $libreOffice

                            if (-not $convertedPdf -or -not (Test-Path $convertedPdf)) {
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) CONVERSION ERROR: LibreOffice failed to generate PDF for $cleanOriginalName"
                                Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                                $sw.Stop()
                                continue
                            }

                            # Rename temp PDF to avoid collision
                            $tempPdfToClean = Join-Path $path "($([System.IO.Path]::GetFileNameWithoutExtension($cleanOriginalName)))_CONVERTED_TEMP.pdf"
                            Move-Item -Path $convertedPdf -Destination $tempPdfToClean -Force -ErrorAction SilentlyContinue
                            $pdfToPrintPath = $tempPdfToClean
                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) CONVERSION SUCCESS: Generated temporary PDF for printing."
                        }

                        # --- PDF PRE-FLIGHT INSPECTION ---
                        $integrity = Test-PdfIntegrity $pdfToPrintPath
                        if (-not $integrity.Valid) {
                            Move-ToFolder $currentPath $cleanOriginalName $wrongFormatFolder | Out-Null
                            if ($tempPdfToClean -and (Test-Path $tempPdfToClean)) { Remove-Item -Path $tempPdfToClean -Force -ErrorAction SilentlyContinue }
                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) PRE-FLIGHT REJECTED: $cleanOriginalName Reason: $($integrity.Reason)"
                            $sw.Stop()
                            continue
                        }
                        Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) PRE-FLIGHT PASSED: PDF structure valid."

                        Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) VALID FORMAT CONFIRMED: $cleanOriginalName"

                        # --- DYNAMIC TIMEOUT & FILE METRICS ANALYSIS ---
                        $fileItem = Get-Item $pdfToPrintPath
                        $fileSizeMB = [math]::Round($fileItem.Length / 1MB, 2)
                        $pageCount = Get-PdfPageCount -filePath $pdfToPrintPath

                        # Formula: Base 60s + 30s/page + 20s/MB
                        $dynamicTimeout = 60 + ($pageCount * 30) + [math]::Round($fileSizeMB * 20)

                        # Enforce Safety Clamps: Floor 120s (2m), Ceiling 900s (15m)
                        if ($dynamicTimeout -lt 120) { $dynamicTimeout = 120 }
                        if ($dynamicTimeout -gt 900) { $dynamicTimeout = 900 }

                        $maxSpoolWait = $dynamicTimeout
                        Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) METRICS ANALYZED: $pageCount Page(s) | ${fileSizeMB} MB | Calculated Dynamic Timeout: ${maxSpoolWait}s"

                        # --- STATE TAGGING ---
                        $currentPath = Update-FileState $currentPath $cleanOriginalName "_PROCESSING"

                        # --- REFRESH PATH FOR NATIVE PDFs AFTER STATE TAGGING ---
                        if ($extension -eq ".pdf") {
                            $pdfToPrintPath = $currentPath
                        }

                        if ($cfg.AutoPrint -eq $true) {
                            
                            $health = Get-PrinterHealthStatus
                            if (-not $health.PrinterName) {
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) PRINT FAILED: No default printer configured."
                                if ($tempPdfToClean -and (Test-Path $tempPdfToClean)) { Remove-Item -Path $tempPdfToClean -Force -ErrorAction SilentlyContinue }
                                Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                                $sw.Stop()
                                continue
                            }

                            # Purge Error Print Jobs
                            try {
                                Get-PrintJob -PrinterName $health.PrinterName -ErrorAction SilentlyContinue | 
                                    Where-Object { $_.JobStatus -like "*Error*" -or $_.JobStatus -like "*UserAction*" } | 
                                    Remove-PrintJob -ErrorAction SilentlyContinue
                            } catch {}

                            # Pre-Print Readiness Retry
                            $printerReady = $false
                            $pRetry = 0
                            while (-not $printerReady -and $pRetry -lt 3) {
                                $health = Get-PrinterHealthStatus
                                if ($health.IsReady) {
                                    $printerReady = $true
                                } else {
                                    $pRetry++
                                    Start-Sleep -Seconds 2
                                }
                            }

                            if (-not $printerReady) {
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) PRINT DEFERRED: Printer unready ($($health.StateText))."
                                if ($tempPdfToClean -and (Test-Path $tempPdfToClean)) { Remove-Item -Path $tempPdfToClean -Force -ErrorAction SilentlyContinue }
                                Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                                $sw.Stop()
                                continue
                            }

                            # Process-Isolated Print Spooling via SumatraPDF
                            if (Test-Path $sumatra) {
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) DISPATCHING: Sending PDF to SumatraPDF..."

                                $pInfo = New-Object System.Diagnostics.ProcessStartInfo
                                $pInfo.FileName = $sumatra
                                $pInfo.Arguments = "-silent -print-settings `"simplex,fit`" -print-to `"$($health.PrinterName)`" `"$pdfToPrintPath`""
                                $pInfo.CreateNoWindow = $true
                                $pInfo.UseShellExecute = $false

                                $proc = [System.Diagnostics.Process]::Start($pInfo)
                                if ($proc) {
                                    $proc.WaitForExit(20000)
                                    $proc.Close()
                                }
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) SPOOLED TO HARDWARE: Sent to '$($health.PrinterName)'"
                            } else {
                                Start-Process -FilePath $pdfToPrintPath -Verb Print -WindowStyle Hidden -ErrorAction SilentlyContinue
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) SPOOLED TO HARDWARE: Sent via shell"
                            }

                            # --- REAL-TIME SPOOLER MONITOR & TIMER FREEZE ENGINE ---
                            Start-Sleep -Seconds 2
                            $spoolTimer       = 0
                            $faultTimer       = 0
                            $maxFaultWait     = 600 # 10 Minutes Human Intervention Cap
                            $printSuccess     = $false
                            $lastLoggedState  = "NORMAL"

                            # Dedicated Physical Print Duration Stopwatch
                            $physicalSw = [System.Diagnostics.Stopwatch]::StartNew()

                            while ($spoolTimer -lt $maxSpoolWait) {
                                $currentJobs = Get-PrintJob -PrinterName $health.PrinterName -ErrorAction SilentlyContinue
                                $jobCount = if ($currentJobs) { ($currentJobs | Measure-Object).Count } else { 0 }

                                $hCheck = Get-PrinterHealthStatus
                                $currentState = $hCheck.StateText

                                # State Change & Logging
                                if ($currentState -ne $lastLoggedState) {
                                    if ($currentState -eq "PAPER_OUT") {
                                        $currentPath = Update-FileState $currentPath $cleanOriginalName "_WAITING_PAPER"
                                        Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) HARDWARE ALERT: Paper Out detected (Dynamic Timeout Clock Frozen)."
                                    } elseif ($currentState -eq "PAPER_JAM") {
                                        $currentPath = Update-FileState $currentPath $cleanOriginalName "_PAPER_JAMMED"
                                        Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) HARDWARE ALERT: Paper Jam detected (Dynamic Timeout Clock Frozen)."
                                    } elseif ($currentState -eq "HARDWARE_FAULT") {
                                        if ($jobCount -gt 0) {
                                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) PRINTER BUSY: Physically printing in progress..."
                                        } else {
                                            $currentPath = Update-FileState $currentPath $cleanOriginalName "_HARDWARE_ERROR"
                                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) HARDWARE ALERT: General printer fault."
                                        }
                                    } elseif ($currentState -eq "NORMAL") {
                                        Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) HARDWARE RECOVERED: Printer state cleared (Dynamic Timeout Clock Resumed)."
                                    }
                                    $lastLoggedState = $currentState
                                }

                                # Physical Print Queue Check
                                if ($jobCount -eq 0) {
                                    $physicalSw.Stop()
                                    $physDurationText = "{0:N2}" -f $physicalSw.Elapsed.TotalSeconds
                                    Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) HARDWARE PRINT COMPLETE: Spooler queue clear (PHYSICAL PRINT TIME: ${physDurationText}s)"
                                    $printSuccess = $true
                                    break
                                } else {
                                    Start-Sleep -Seconds 3
                                    
                                    # TIMER FREEZE LOGIC:
                                    $isFaultActive = ($currentState -in @("PAPER_OUT", "PAPER_JAM")) -or ($currentState -eq "HARDWARE_FAULT" -and $jobCount -eq 0)

                                    if ($isFaultActive) {
                                        $faultTimer += 3
                                        if ($faultTimer -ge $maxFaultWait) {
                                            Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) HUMAN INTERVENTION TIMEOUT (${maxFaultWait}s): Unresolved hardware error ($currentState)."
                                            break
                                        }
                                    } else {
                                        $spoolTimer += 3
                                        $faultTimer = 0
                                    }
                                }
                            }

                            # Cleanup temporary converted PDF file
                            if ($tempPdfToClean -and (Test-Path $tempPdfToClean)) {
                                Remove-Item -Path $tempPdfToClean -Force -ErrorAction SilentlyContinue
                            }

                            if ($printSuccess) {
                                $moved = Move-ToFolder $currentPath $cleanOriginalName $donePrintFolder
                                $totalSecs = "{0:N2}" -f $sw.Elapsed.TotalSeconds
                                if ($moved) {
                                    Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) ARCHIVED: $cleanOriginalName moved to donePrint (TOTAL TIME: ${totalSecs}s / MAX ALLOWED: ${maxSpoolWait}s)"
                                }
                            } else {
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) TIMEOUT EXCEEDED: Isolating file..."
                                try {
                                    $stuckJobs = Get-PrintJob -PrinterName $health.PrinterName -ErrorAction SilentlyContinue
                                    if ($stuckJobs) { $stuckJobs | Remove-PrintJob -ErrorAction SilentlyContinue }
                                } catch {}

                                Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                                Add-Content -Path $log -Value "[ $timeStamp ] $(Get-ElapsedText $sw) MOVED TO retryPrint: $cleanOriginalName isolated."
                            }
                        }
                        $sw.Stop()
                    } catch {
                        if ($tempPdfToClean -and (Test-Path $tempPdfToClean)) { Remove-Item -Path $tempPdfToClean -Force -ErrorAction SilentlyContinue }
                        Add-Content -Path $log -Value "[ $timeStamp ] FILE RECOVERY: Exception on $cleanOriginalName. Moving to retryPrint..."
                        Move-ToFolder $currentPath $cleanOriginalName $retryPrintFolder | Out-Null
                    }
                }
            } catch {}

            # GARBAGE COLLECTION SWEEP
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Start-Sleep -Seconds 1.5
        }
    }

    Start-Job -ScriptBlock $scriptBlock -Name $jobName -ArgumentList $folderToWatch, $logFile, $configFile, $sumatraPath, $libreOfficePath | Out-Null
    Write-Host "SUCCESS: Active monitoring started for $folderToWatch" -ForegroundColor Green
}

function Stop-Watcher {
    Write-Host "Stopping folder watcher engine..." -ForegroundColor Cyan
    $job = Get-Job -Name $jobName -ErrorAction SilentlyContinue
    if ($job) {
        Stop-Job -Name $jobName -ErrorAction SilentlyContinue
        Remove-Job -Name $jobName -Force -ErrorAction SilentlyContinue
    }
    Write-Host "SUCCESS: Monitor deactivated." -ForegroundColor Red
}

# --- HEADLESS EXECUTION (TASK SCHEDULER) ---
if ($Headless) {
    $globalConfig.AutoPrint = $true
    Save-PipelineConfig -path $configFile -configObject $globalConfig
    Start-Watcher
    while ($true) { Start-Sleep -Seconds 10 }
}

# ------------------------------------------------------------------------------
# 4. INTERACTIVE MENU INTERFACE
# ------------------------------------------------------------------------------
while ($true) {
    Clear-Host
    $status = Get-WatcherStatus
    $statusColor = if ($status -eq "RUNNING") { "Green" } else { "Red" }
    $currentConfig = Get-PipelineConfig -path $configFile
    $printStatusColor = if ($currentConfig.AutoPrint) { "Green" } else { "DarkGray" }
    $currentPrinter = Get-DefaultPrinterName

    Write-Host "==================================================" -ForegroundColor DarkGray
    Write-Host "   WIN 11 FILTER, AUTOPRINT AND ARCHIVE PIPELINE  " -ForegroundColor White
    Write-Host "==================================================" -ForegroundColor DarkGray
    Write-Host " Watcher Engine Status:  " -NoNewline
    Write-Host "[ $status ]" -ForegroundColor $statusColor
    Write-Host " Auto-Print Mode:        " -NoNewline
    Write-Host "[ $(if($currentConfig.AutoPrint){'ON'}else{'OFF'}) ]" -ForegroundColor $printStatusColor
    Write-Host " Pipeline Version:       " -NoNewline
    Write-Host "[ v1.2.1-AuditTest ]" -ForegroundColor Yellow
    Write-Host " Target Active Printer:  $currentPrinter" -ForegroundColor Yellow
    Write-Host " Target Folder Path:     $folderToWatch" -ForegroundColor Gray
    Write-Host "==================================================" -ForegroundColor DarkGray
    Write-Host " 1. TURN ON Folder Watcher"
    Write-Host " 2. TURN OFF Folder Watcher"
    Write-Host " 3. TOGGLE Auto-Print Mode (ON/OFF)"
    Write-Host " 4. View System Filter and Print Logs"
    Write-Host " 5. Exit Menu"
    Write-Host "==================================================" -ForegroundColor DarkGray

    $choice = Read-Host "Select an option (1-5)"

    switch ($choice) {
        "1" { Start-Watcher; Start-Sleep -Seconds 2 }
        "2" { Stop-Watcher; Start-Sleep -Seconds 2 }
        "3" { 
            $currentConfig.AutoPrint = -not $currentConfig.AutoPrint
            Save-PipelineConfig -path $configFile -configObject $currentConfig
            Write-Host "Auto-Print shifted to: $($currentConfig.AutoPrint)" -ForegroundColor Cyan
            Start-Sleep -Seconds 1
        }
        "4" { 
            if (Test-Path $logFile) { Invoke-Item $logFile } 
            else { Write-Host "No logs generated yet." -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
        }
        "5" { 
            Write-Host "Exiting application. Stopping background jobs..." -ForegroundColor Yellow
            Stop-Watcher
            Start-Sleep -Seconds 1
            exit 
        }
    }
}