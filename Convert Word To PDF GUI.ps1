# Graphical version of "Convert Word To PDF.ps1".
# Same conversion, hang detection and Vercel upload, shown in a window with progress bars and a log.
# Runs in Windows PowerShell 5.1 (powershell.exe), which starts in STA mode as WinForms needs.

# Define the folder paths
$wordFolderPath   = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide\Editable format"
$pdfFolderPath    = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide"
$uploadScript     = Join-Path $PSScriptRoot "Upload PDFs To Vercel.ps1"
$hangLimitSeconds = 45   # Word at 0% CPU for this long is treated as hung

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Get all Word documents in the folder (skip Word's "~$" lock files)
$script:wordFiles   = @(Get-ChildItem -Path $wordFolderPath -Filter "*.docx" | Where-Object { $_.Name -notlike '~$*' })
$script:total       = $script:wordFiles.Count
$script:next        = 0          # files started so far
$script:done        = 0          # files converted, skipped or failed
$script:state       = "Idle"     # Idle, Converting, Uploading, Finished, Failed
$script:job         = $null
$script:jobStart    = $null
$script:lastSample  = $null
$script:lastCpu     = $null
$script:idleSeconds = 0
$script:wordPid     = $null
$script:currentName = ""
$script:currentPdf  = ""

# ---------- Window ----------

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Books to PDF"
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.ClientSize      = New-Object System.Drawing.Size(640, 470)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false

function New-Label($text, $x, $y, $width, $height) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.SetBounds($x, $y, $width, $height)
    $form.Controls.Add($label)
    return $label
}

$lblTitle = New-Label "Converting The Guide to PDF" 16 12 608 32
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)

$lblOverall = New-Label "Documents: 0 of $script:total" 16 52 608 20

$barOverall = New-Object System.Windows.Forms.ProgressBar
$barOverall.SetBounds(16, 76, 608, 24)
$barOverall.Minimum = 0
$barOverall.Maximum = [math]::Max($script:total, 1)
$form.Controls.Add($barOverall)

$lblCurrent = New-Label "Starting..." 16 110 608 22

# Word does not report a percentage, so the file bar only shows activity
$barCurrent = New-Object System.Windows.Forms.ProgressBar
$barCurrent.SetBounds(16, 136, 608, 16)
$barCurrent.Style = "Marquee"
$barCurrent.MarqueeAnimationSpeed = 30
$form.Controls.Add($barCurrent)

$lblElapsed = New-Label "" 16 160 608 20

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.SetBounds(16, 188, 608, 220)
$txtLog.Multiline  = $true
$txtLog.ReadOnly   = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Font       = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.SetBounds(528, 422, 96, 32)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# ---------- Helpers ----------

function Write-Log($text) {
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format "HH:mm:ss"), $text))
}

function Show-Overall {
    $barOverall.Value = [math]::Min($script:done, $barOverall.Maximum)
    $lblOverall.Text  = "Documents: $script:done of $script:total"
}

function Stop-Word($processId) {
    if ($processId) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }
}

function Set-Finished($text, $color) {
    $timer.Stop()
    $barCurrent.Style = "Continuous"
    $barCurrent.Value = 0
    $lblCurrent.ForeColor = $color
    $lblCurrent.Text = $text
    $lblElapsed.Text = ""
    Write-Log $text
}

function Set-Failed($text) {
    $script:state = "Failed"
    Set-Finished $text ([System.Drawing.Color]::Firebrick)
}

function Set-Succeeded($text) {
    $script:state = "Finished"
    Set-Finished $text ([System.Drawing.Color]::Green)
}

# Starts the next Word export, skipping PDFs that are already up to date. Starts the upload when nothing is left.
function Start-NextStep {
    while ($script:next -lt $script:total) {
        $file    = $script:wordFiles[$script:next]
        $pdfPath = Join-Path -Path $pdfFolderPath -ChildPath ($file.BaseName + ".pdf")
        $script:next++

        if ((Test-Path $pdfPath) -and -not ($file.LastWriteTime -gt (Get-Item $pdfPath).LastWriteTime)) {
            $script:done++
            Write-Log "[$script:next/$script:total] Skipped '$($file.Name)' (PDF is already up to date)."
            Show-Overall
            continue
        }

        # Start Word inside the background job
        $script:job = Start-Job -ScriptBlock {
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

        $script:state       = "Converting"
        $script:currentName = $file.Name
        $script:currentPdf  = $pdfPath
        $script:jobStart    = Get-Date
        $script:lastSample  = $script:jobStart
        $script:lastCpu     = $null
        $script:idleSeconds = 0
        $script:wordPid     = $null

        $lblCurrent.ForeColor = [System.Drawing.Color]::Black
        $lblCurrent.Text = "Rendering '$($file.Name)' ($script:next of $script:total)"
        $barCurrent.Style = "Marquee"
        Write-Log "[$script:next/$script:total] Rendering '$($file.Name)'..."
        return
    }

    Start-Upload
}

function Start-Upload {
    $bookNames = @($script:wordFiles.BaseName)
    $script:job = Start-Job -FilePath $uploadScript -ArgumentList $pdfFolderPath, $bookNames
    $script:state      = "Uploading"
    $script:jobStart   = Get-Date
    $lblCurrent.ForeColor = [System.Drawing.Color]::Black
    $lblCurrent.Text   = "Uploading PDFs to Vercel Blob..."
    $barCurrent.Style  = "Marquee"
    Write-Log "All documents processed. Uploading to Vercel Blob..."
}

function Complete-Conversion {
    $elapsed  = [int]((Get-Date) - $script:jobStart).TotalSeconds
    $pdfPath  = $script:currentPdf
    Receive-Job $script:job -ErrorAction SilentlyContinue | Out-Null
    $jobError = $script:job.ChildJobs[0].Error
    Remove-Job $script:job
    $script:job = $null
    $script:done++

    if ($jobError) {
        Set-Failed "Error processing '$($script:currentName)': $($jobError[0])"
        return
    }

    if (Test-Path $pdfPath) {
        $pdfInfo  = Get-Item $pdfPath
        $sizeInKB = [math]::Round(($pdfInfo.Length / 1KB), 2)
        Write-Log "SUCCESS: Completed '$($pdfInfo.Name)' ($sizeInKB KB) in $elapsed seconds."
    }
    Show-Overall
    Start-NextStep
}

# Watches Word's CPU use. Samples every 2 seconds and kills Word after $hangLimitSeconds at 0% CPU.
function Test-WordHealth {
    $wordProc = Get-Process -Name "winword" -ErrorAction SilentlyContinue |
                Sort-Object -Property StartTime -Descending |
                Select-Object -First 1
    if (-not $wordProc) { return }

    $now = Get-Date
    if (($now - $script:lastSample).TotalSeconds -lt 2) { return }

    $script:wordPid = $wordProc.Id
    $cpu = $wordProc.Cpu
    if (($null -ne $script:lastCpu) -and (($cpu - $script:lastCpu) -le 0.01)) {
        $script:idleSeconds += ($now - $script:lastSample).TotalSeconds
    } else {
        $script:idleSeconds = 0   # It's working!
    }
    $script:lastCpu    = $cpu
    $script:lastSample = $now

    if ($script:idleSeconds -ge $hangLimitSeconds) {
        Stop-Job $script:job -ErrorAction SilentlyContinue
        Remove-Job $script:job -Force -ErrorAction SilentlyContinue
        $script:job = $null
        Stop-Word $script:wordPid
        Set-Failed "CRITICAL ERROR: Background Word instance (PID: $($script:wordPid)) flatlined at 0% CPU while processing $($script:currentName) and was killed."
    }
}

function Complete-Upload {
    # Flush any log lines the upload wrote after the last tick
    Receive-Job $script:job -ErrorAction SilentlyContinue |
        ForEach-Object { if ("$_".Trim()) { Write-Log ("$_" -replace "`r", "") } }
    $jobError = $script:job.ChildJobs[0].Error
    $jobFailed = $script:job.State -eq "Failed"
    Remove-Job $script:job
    $script:job = $null

    if ($jobFailed -or $jobError) {
        $reason = if ($jobError) { $jobError[0] } else { "upload job failed" }
        Set-Failed "Vercel upload failed: $reason"
        return
    }
    $elapsed = [int]((Get-Date) - $script:jobStart).TotalSeconds
    Set-Succeeded "All documents processed and uploaded in $elapsed seconds."
}

# Runs every 500 ms on the UI thread: advances the current job and updates the window
function Update-Progress {
    if ($script:state -eq "Converting") {
        if ($script:job.State -eq "Running") {
            $lblElapsed.Text = "Elapsed: $([int]((Get-Date) - $script:jobStart).TotalSeconds)s  [ALIVE]"
            Test-WordHealth
        } else {
            Complete-Conversion
        }
    }
    elseif ($script:state -eq "Uploading") {
        # Show the upload log while it runs
        Receive-Job $script:job -ErrorAction SilentlyContinue |
            ForEach-Object { if ("$_".Trim()) { Write-Log ("$_" -replace "`r", "") } }
        if ($script:job.State -eq "Running") {
            $lblElapsed.Text = "Elapsed: $([int]((Get-Date) - $script:jobStart).TotalSeconds)s  [ALIVE]"
        } else {
            Complete-Upload
        }
    }
}

# ---------- Wiring ----------

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({ Update-Progress })

$form.Add_Shown({
    Write-Log "Books to PDF: $script:total Word document(s) in '$wordFolderPath'"
    $timer.Start()
    Start-NextStep
})

# Closing during a run asks first, then stops the job and the Word instance it started
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:state -eq "Converting" -or $script:state -eq "Uploading") {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "A conversion or upload is still running. Stop it and close?",
            "Books to PDF",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $e.Cancel = $true
            return
        }
        if ($script:job) {
            Stop-Job $script:job -ErrorAction SilentlyContinue
            Remove-Job $script:job -Force -ErrorAction SilentlyContinue
        }
        if ($script:state -eq "Converting") { Stop-Word $script:wordPid }
    }
    $timer.Stop()
})

[void]$form.ShowDialog()
$timer.Dispose()
$form.Dispose()
