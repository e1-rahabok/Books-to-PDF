# Fake Word and Vercel jobs for "Convert Word To PDF GUI.ps1 -Demo".
# Nothing is converted, uploaded or sent. Each job waits a few seconds and replays canned output,
# so the window can be tried out safely. Demo files live in the temp folder, not in the Box folders.

$demoRoot       = Join-Path $env:TEMP "BooksToPDF-Demo"
$demoWordFolder = Join-Path $demoRoot "Word"
$demoPdfFolder  = Join-Path $demoRoot "PDF"
$demoNames      = @("Executive Summary", "PART 1 - Concepts", "PART 2 - Consumption", "PART 3 - Counters", "PART 4 - Miscellaneous")
$demoFailName   = "PART 3 - Counters"   # this document fails to convert, to show the error state. Set to "" for none.

# Fresh folders on every start: empty .docx files, and a PDF for the first document that is newer than its Word file,
# so that one shows as "Up to date" as it would on a real run.
function Reset-DemoFolders {
    Remove-Item $demoRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $demoWordFolder, $demoPdfFolder | Out-Null
    foreach ($name in $demoNames) {
        $docx = Join-Path $demoWordFolder "$name.docx"
        New-Item -ItemType File -Path $docx | Out-Null
        (Get-Item $docx).LastWriteTime = (Get-Date).AddHours(-2)
    }
    New-Item -ItemType File -Path (Join-Path $demoPdfFolder "$($demoNames[0]).pdf") | Out-Null
}

# A fake job. State flips to Completed after $durationMs. Timeline lines are handed out by Receive-Job when their time comes.
function New-DemoJob($durationMs, $timeline, $failure) {
    $errors = @()
    if ($failure) { $errors = @($failure) }
    $job = [pscustomobject]@{
        Start     = Get-Date
        Ends      = (Get-Date).AddMilliseconds($durationMs)
        Timeline  = @($timeline)
        Sent      = 0
        ChildJobs = @([pscustomobject]@{ Error = $errors })
    }
    $job | Add-Member -MemberType ScriptProperty -Name State -Value { if ((Get-Date) -ge $this.Ends) { "Completed" } else { "Running" } }
    return $job
}

function New-DemoConvertJob($docxPath, $pdfPath) {
    $name    = [IO.Path]::GetFileNameWithoutExtension($docxPath)
    $failure = $null
    if ($name -eq $demoFailName) {
        $failure = "Word could not open the document (simulated)"
    } else {
        New-Item -ItemType File -Path $pdfPath -Force | Out-Null
    }
    return New-DemoJob 2500 @() $failure
}

function New-DemoUploadJob {
    $timeline = @()
    $at    = 0
    $count = $demoNames.Count
    $timeline += @{ At = $at; Text = "Vercel Blob: checking $count PDF(s) for changes, uploading only changed files to 'downloads/'" }
    $at += 400
    for ($i = 0; $i -lt $count; $i++) {
        $name = $demoNames[$i]
        $n    = $i + 1
        if ($i -eq 0) {
            $timeline += @{ At = $at; Text = "[$n/$count] Unchanged: '$name.pdf' (not uploaded)" }
            $at += 400
            continue
        }
        $timeline += @{ At = $at; Text = "[$n/$count] Uploading '$name.pdf' (0.4 MB)..." }
        $at += 1200
        $timeline += @{ At = $at; Text = "`r[$n/$count] UPLOADED: https://demo.invalid/downloads/$name.pdf" }
        $at += 400
    }
    return New-DemoJob ($at + 300) $timeline $null
}

# Replaces Start-Job. Upload jobs are started with -FilePath, Word conversions with -ScriptBlock.
function Start-Job {
    param($ScriptBlock, $FilePath, $ArgumentList)
    if ($FilePath) { return New-DemoUploadJob }
    return New-DemoConvertJob $ArgumentList[0] $ArgumentList[1]
}

function Receive-Job {
    param($Job, $ErrorAction)
    $elapsed = ((Get-Date) - $Job.Start).TotalMilliseconds
    while ($Job.Sent -lt $Job.Timeline.Count -and $Job.Timeline[$Job.Sent].At -le $elapsed) {
        $Job.Timeline[$Job.Sent].Text
        $Job.Sent++
    }
}

function Remove-Job { }
function Stop-Job { }

# In demo mode the real Word process check is off, so no real Word instance is ever looked at or killed
function Test-WordHealth { }
