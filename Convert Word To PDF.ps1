# Define the folder paths
$wordFolderPath = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide\Editable format"
$pdfFolderPath  = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide"

$Host.UI.RawUI.WindowTitle = "Books to PDF"

# Keep the window open when launched from the desktop shortcut
function Wait-AndExit($code) {
    Write-Host ""
    Read-Host "Press Enter to close"
    exit $code
}

# Get all Word documents in the folder (skip Word's "~$" lock files)
$wordFiles = @(Get-ChildItem -Path $wordFolderPath -Filter "*.docx" | Where-Object { $_.Name -notlike '~$*' })
$total = $wordFiles.Count
$index = 0

Write-Host "Books to PDF: $total Word document(s) in '$wordFolderPath'" -ForegroundColor Cyan
Write-Host ""

foreach ($file in $wordFiles) {
    $index++
    $pdfPath = Join-Path -Path $pdfFolderPath -ChildPath ($file.BaseName + ".pdf")

    # Check if the PDF needs updating
    if (-not (Test-Path $pdfPath) -or ($file.LastWriteTime -gt (Get-Item $pdfPath).LastWriteTime)) {

        # Start Word inside the background job
        $job = Start-Job -ScriptBlock {
            param($fileFullName, $pdfPath)
            try {
                $wordApp = New-Object -ComObject Word.Application
                $wordApp.Visible = $false

                $wordPid = (Get-Process -Name "winword" | Sort-Object -Property StartTime -Descending | Select-Object -First 1).Id

                $doc = $wordApp.Documents.Open($fileFullName)
                $doc.ExportAsFixedFormat(
                    $pdfPath,
                    17,
                    $false,
                    1,
                    0, 0, 0, 0, $true, $true, 1, $true, $true, $true
                )
                $doc.Close([ref]$false)
                $wordApp.Quit()

                return @{ Success = $true; PID = $wordPid }
            } catch {
                if ($wordApp) { $wordApp.Quit() }
                throw $_
            }
        } -ArgumentList $file.FullName, $pdfPath

        # Intelligent Hang Detection & Live Progress Indicator
        $consecutiveIdleTicks = 0
        $isHung = $false
        $targetPid = $null
        $elapsedSeconds = 0

        # Loop while the massive 500-page file is rendering
        while ($job.State -eq "Running") {
            # 1. Update the user on screen that it's alive (rewrites the same line cleanly)
            Write-Host -NoNewline "`r[$index/$total] Rendering '$($file.Name)'... (Elapsed: ${elapsedSeconds}s) [ALIVE]" -ForegroundColor Yellow

            Start-Sleep -Seconds 2
            $elapsedSeconds += 2

            # 2. Grab the background Word process to check health
            $wordProc = Get-Process -Name "winword" -ErrorAction SilentlyContinue |
                        Sort-Object -Property StartTime -Descending |
                        Select-Object -First 1

            if ($wordProc) {
                $targetPid = $wordProc.Id

                # Snapshot CPU cycles to prove it hasn't flatlined
                $cpuBefore = $wordProc.Cpu
                Start-Sleep -Milliseconds 500
                $wordProcRefresh = Get-Process -Id $targetPid -ErrorAction SilentlyContinue

                if ($wordProcRefresh) {
                    $cpuDelta = $wordProcRefresh.Cpu - $cpuBefore

                    if ($cpuDelta -le 0.01) {
                        $consecutiveIdleTicks++
                    } else {
                        $consecutiveIdleTicks = 0 # It's working!
                    }
                }

                # Only abort if it's genuinely unresponsive at 0% CPU for a full 45 seconds
                if ($consecutiveIdleTicks -ge 18) {
                    $isHung = $true
                    break
                }
            }
        }

        # Finish the processing text line
        Write-Host ""

        # Handle an actual hang safely
        if ($isHung) {
            Remove-Job $job -Force
            if ($targetPid) {
                Stop-Process -Id $targetPid -Force -ErrorAction SilentlyContinue
                Write-Host "CRITICAL ERROR: Background Word instance (PID: $targetPid) flatlined at 0% CPU while processing $($file.Name) and was killed." -ForegroundColor Red
            }
            Wait-AndExit 1
        }

        # Safe verification of completed jobs
        $jobResult = Receive-Job $job -ErrorAction SilentlyContinue
        $jobError  = $job.ChildJobs[0].Error
        Remove-Job $job

        if ($jobError) {
            Write-Host "Error processing $($file.Name): $($jobError[0])" -ForegroundColor Red
            Wait-AndExit 1
        }

        # Log file completion details with final size
        if (Test-Path $pdfPath) {
            $pdfInfo = Get-Item $pdfPath
            $sizeInKB = [math]::Round(($pdfInfo.Length / 1KB), 2)
            Write-Host "[$index/$total] SUCCESS: Completed '$($pdfInfo.Name)' ($sizeInKB KB) in $elapsedSeconds seconds." -ForegroundColor Green
        }
    } else {
        Write-Host "[$index/$total] Skipped: '$($file.Name)' (PDF is already up to date)."
    }
}

Write-Host ""
Write-Host "All documents processed successfully." -ForegroundColor Cyan
Write-Host ""

# Upload new or changed PDFs to the Vercel Blob store
try {
    & (Join-Path $PSScriptRoot "Upload PDFs To Vercel.ps1") -PdfFolderPath $pdfFolderPath -BookNames $wordFiles.BaseName
} catch {
    Wait-AndExit 1
}
Wait-AndExit 0
