# Graphical version of "Convert Word To PDF.ps1".
# Same conversion, hang detection and Vercel upload, shown in a dark window with a Start button,
# one status row per document and a log. Nothing runs until Start is pressed.
# Runs in Windows PowerShell 5.1 (powershell.exe), which starts in STA mode as WinForms needs.
# -Demo runs the window on fake data (see "Demo Fake Jobs.ps1"): nothing is converted or uploaded.
param([switch]$Demo)

# Define the folder paths
$wordFolderPath   = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide\Editable format"
$pdfFolderPath    = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide"
$uploadScript     = Join-Path $PSScriptRoot "Upload PDFs To Vercel.ps1"
$iconPath         = Join-Path $PSScriptRoot "Books to PDF.ico"
$hangLimitSeconds = 45   # Word at 0% CPU for this long is treated as hung
$maxRows          = 5    # the window has room for this many document rows
$errorLog         = Join-Path $env:TEMP "BooksToPDF-error.log"   # unexpected errors are written here
$lastRunFile      = Join-Path $env:LOCALAPPDATA "BooksToPDF\last-run.txt"   # when the last run finished

if ($Demo) {
    . (Join-Path $PSScriptRoot "Demo Fake Jobs.ps1")
    Reset-DemoFolders
    $lastRunFile   = Join-Path $demoRoot "last-run.txt"
    $wordFolderPath = $demoWordFolder
    $pdfFolderPath  = $demoPdfFolder
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Errors thrown in window events go to the error log instead of the .NET crash dialog.
# Must be set before any window control is created.
[System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[System.Windows.Forms.Application]::add_ThreadException({
    param($source, $e)
    Write-ErrorReport $e.Exception
})

# Plain WinForms cannot draw rounded corners or gradients, so the look is built here.
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class Theme {
    public static readonly Color Top     = Color.FromArgb(9, 16, 42);
    public static readonly Color Bottom  = Color.FromArgb(14, 44, 112);
    public static readonly Color Card    = Color.FromArgb(20, 34, 78);
    public static readonly Color Border  = Color.FromArgb(52, 80, 140);
    public static readonly Color Track   = Color.FromArgb(38, 60, 118);
    public static readonly Color Accent  = Color.FromArgb(56, 120, 255);
    public static readonly Color Accent2 = Color.FromArgb(104, 196, 255);
    public static readonly Color Text    = Color.FromArgb(238, 244, 255);
    public static readonly Color Muted   = Color.FromArgb(140, 162, 206);
    public static readonly Color Good    = Color.FromArgb(128, 222, 120);
    public static readonly Color Bad     = Color.FromArgb(255, 112, 112);

    public static void PaintBackground(Graphics g, Rectangle r) {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (LinearGradientBrush lb = new LinearGradientBrush(r, Top, Bottom, LinearGradientMode.Vertical)) {
            g.FillRectangle(lb, r);
        }
        // Soft glow in the top-right corner
        using (GraphicsPath path = new GraphicsPath()) {
            path.AddEllipse(r.Right - r.Width, r.Top - r.Height / 2, r.Width * 2, r.Height);
            using (PathGradientBrush pb = new PathGradientBrush(path)) {
                pb.CenterColor = Color.FromArgb(80, Accent);
                pb.SurroundColors = new Color[] { Color.FromArgb(0, Accent) };
                g.FillPath(pb, path);
            }
        }
    }

    public static GraphicsPath RoundRect(RectangleF r, float radius) {
        float d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
        GraphicsPath p = new GraphicsPath();
        if (d <= 0) {
            p.AddRectangle(r);
            return p;
        }
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}

// Rounded, opaque panel with a thin border
public class Card : Panel {
    public Card() {
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Theme.Card;
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (Width < 1 || Height < 1) return;
        // Clip the corners so the window gradient shows through
        using (GraphicsPath p = Theme.RoundRect(new RectangleF(0, 0, Width, Height), 16)) {
            Region = new System.Drawing.Region(p);
        }
        Invalidate();
    }

    protected override void OnPaintBackground(PaintEventArgs e) {
        e.Graphics.Clear(BackColor);
    }

    protected override void OnPaint(PaintEventArgs e) {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using (GraphicsPath p = Theme.RoundRect(new RectangleF(0.5f, 0.5f, Width - 1, Height - 1), 16))
        using (Pen pen = new Pen(Theme.Border)) {
            e.Graphics.DrawPath(pen, p);
        }
    }
}

// Rounded progress bar. Marquee mode slides a highlight back and forth, for steps without a percentage.
public class GradientBar : Control {
    int max = 100;
    int val;
    bool marquee;
    int offset;
    Timer timer;

    public GradientBar() {
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
        timer = new Timer();
        timer.Interval = 25;
        timer.Tick += delegate {
            offset += 6;
            if (offset > Width) offset = -(Width / 3);
            Invalidate();
        };
    }

    public int Maximum { get { return max; } set { max = Math.Max(1, value); Invalidate(); } }
    public int Value { get { return val; } set { val = Math.Max(0, Math.Min(max, value)); Invalidate(); } }

    public bool Marquee {
        get { return marquee; }
        set {
            marquee = value;
            offset = -(Width / 3);
            timer.Enabled = value;
            Invalidate();
        }
    }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        RectangleF track = new RectangleF(0, 0, Width - 1, Height - 1);
        float radius = Math.Max(1, Height / 2f);
        using (GraphicsPath tp = Theme.RoundRect(track, radius))
        using (SolidBrush tb = new SolidBrush(Theme.Track)) {
            g.FillPath(tb, tp);

            float fillW = marquee ? Width / 3f : Width * (float)val / max;
            float x = marquee ? offset : 0;
            if (fillW < 2) return;

            g.SetClip(tp);
            using (GraphicsPath fp = Theme.RoundRect(new RectangleF(x, 0, fillW, Height - 1), radius))
            using (LinearGradientBrush lb = new LinearGradientBrush(new PointF(x, 0), new PointF(x + fillW, 0), Theme.Accent, Theme.Accent2)) {
                g.FillPath(lb, fp);
            }
            g.ResetClip();
        }
    }
}

// Pill-shaped gradient button with a hover highlight. Greyed out while disabled.
public class PillButton : Control {
    bool hover;

    public PillButton() {
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
        ForeColor = Color.White;
        Cursor = Cursors.Hand;
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        Color from = !Enabled ? Theme.Track : (hover ? Theme.Accent2 : Theme.Accent);
        Color to   = !Enabled ? Theme.Track : (hover ? Color.FromArgb(170, 225, 255) : Theme.Accent2);
        Color text = Enabled ? ForeColor : Theme.Muted;

        using (GraphicsPath p = Theme.RoundRect(new RectangleF(0, 0, Width - 1, Height - 1), Height / 2f))
        using (LinearGradientBrush lb = new LinearGradientBrush(new PointF(0, 0), new PointF(Width, 0), from, to)) {
            g.FillPath(lb, p);
        }

        using (Font f = new Font("Segoe UI Semibold", 10.5f))
        using (SolidBrush tb = new SolidBrush(text))
        using (StringFormat sf = new StringFormat()) {
            sf.Alignment = StringAlignment.Center;
            sf.LineAlignment = StringAlignment.Center;
            g.DrawString(Text, f, tb, new RectangleF(0, 0, Width, Height), sf);
        }
    }
}

// Main window: dark gradient background and a dark title bar on Windows 11
public class AppForm : Form {
    public AppForm() {
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaintBackground(PaintEventArgs e) {
        Theme.PaintBackground(e.Graphics, ClientRectangle);
    }

    protected override void OnHandleCreated(EventArgs e) {
        base.OnHandleCreated(e);
        int dark = 1;
        Native.DwmSetWindowAttribute(Handle, 20, ref dark, 4);      // dark title bar
        int corner = 2;
        Native.DwmSetWindowAttribute(Handle, 33, ref corner, 4);    // rounded corners
        int caption = Theme.Top.R | (Theme.Top.G << 8) | (Theme.Top.B << 16);
        Native.DwmSetWindowAttribute(Handle, 35, ref caption, 4);   // title bar colour (Windows 11 only, ignored elsewhere)
    }
}

public static class Native {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
'@

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
$script:currentRow  = -1         # row of the document being rendered
$script:currentName = ""
$script:currentPdf  = ""
$script:uploadName  = ""         # PDF base name the upload is working on now
$script:rowLabels   = @()        # status text per document row
$script:rowInfo     = @()        # size and last modified line per document row
$script:rowBars     = @()        # progress bar per document row

# ---------- Window ----------

function New-Label($parent, $text, $x, $y, $width, $height, $font, $color) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.SetBounds($x, $y, $width, $height)
    $label.BackColor = [System.Drawing.Color]::Transparent
    $label.ForeColor = $color
    if ($font) { $label.Font = $font }
    $parent.Controls.Add($label)
    return $label
}

$fontSection = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$fontRow     = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$fontStatus  = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)

$form = New-Object AppForm
$form.Text            = "Vercel Update"
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.ClientSize      = New-Object System.Drawing.Size(780, 868)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
if (Test-Path $iconPath) { $form.Icon = New-Object System.Drawing.Icon($iconPath) }

# Header
$lblTitle    = New-Label $form "Vercel Update" 28 18 500 36 (New-Object System.Drawing.Font("Segoe UI Variable Display", 18, [System.Drawing.FontStyle]::Bold)) ([Theme]::Text)
$lblSubtitle = New-Label $form "Microsoft Word -> PDF -> Upload to Vercel" 28 56 500 22 $null ([Theme]::Muted)
$lblLastRun  = New-Label $form "" 28 84 724 20 $null ([Theme]::Muted)
$lblStatus   = New-Label $form "" 28 104 724 20 $null ([Theme]::Muted)

$btnStart = New-Object PillButton
$btnStart.Text = "Start"
$btnStart.SetBounds(420, 800, 160, 44)
$btnStart.Add_Click({
    try { Start-Run } catch { Write-ErrorReport $_.Exception }
})
$form.Controls.Add($btnStart)

# Document list: one row per slot, slots beyond the number of files stay empty
$cardFiles = New-Object Card
$cardFiles.SetBounds(24, 130, 732, 380)
$form.Controls.Add($cardFiles)

$lblSection1 = New-Label $cardFiles "DOCUMENTS" 20 14 200 20 $fontSection ([Theme]::Muted)
$lblCount    = New-Label $cardFiles "" 512 14 200 20 $fontSection ([Theme]::Muted)
$lblCount.TextAlign = "MiddleRight"

for ($k = 0; $k -lt $maxRows; $k++) {
    $y    = 46 + $k * 66
    $name = if ($k -lt $script:total) { $script:wordFiles[$k].BaseName } else { "" }

    $lblName = New-Label $cardFiles $name 20 $y 460 22 $fontRow ([Theme]::Text)
    $lblRowStatus = New-Label $cardFiles "" 492 $y 220 22 $fontStatus ([Theme]::Muted)
    $lblRowStatus.TextAlign = "MiddleRight"
    $lblInfo = New-Label $cardFiles "" 20 ($y + 26) 692 18 $null ([Theme]::Muted)

    $bar = New-Object GradientBar
    $bar.SetBounds(20, $y + 50, 692, 8)
    $cardFiles.Controls.Add($bar)

    $script:rowLabels += $lblRowStatus
    $script:rowInfo   += $lblInfo
    $script:rowBars   += $bar
    if ($k -ge $script:total) { $bar.Visible = $false }
}

# Log card
$cardLog = New-Object Card
$cardLog.SetBounds(24, 524, 732, 262)
$form.Controls.Add($cardLog)

$lblSection2 = New-Label $cardLog "LOG" 20 14 200 20 $fontSection ([Theme]::Muted)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.SetBounds(18, 40, 696, 206)
$txtLog.Multiline   = $true
$txtLog.ReadOnly    = $true
$txtLog.BorderStyle = "None"
$txtLog.BackColor   = [Theme]::Card
$txtLog.ForeColor   = [Theme]::Text
$txtLog.ScrollBars  = "Vertical"
$txtLog.Font        = New-Object System.Drawing.Font("Consolas", 9.5)
$cardLog.Controls.Add($txtLog)

# Close button
$btnClose = New-Object PillButton
$btnClose.Text = "Close"
$btnClose.SetBounds(592, 800, 160, 44)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# ---------- Helpers ----------

function Write-Log($text) {
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format "HH:mm:ss"), $text))
}

# An error that escaped a handler: keep the full details in the error log and leave the window usable
function Write-ErrorReport($exception) {
    $context = "state=$script:state, document='$script:currentName', done=$script:done of $script:total"
    Add-Content -Path $errorLog -Encoding UTF8 -Value ("[{0}] {1}`r`n{2}`r`n" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $context, $exception.ToString())
    if ($script:state -eq "Converting") {
        if ($script:job) { Stop-Job $script:job -ErrorAction SilentlyContinue; Remove-Job $script:job -Force -ErrorAction SilentlyContinue }
        $script:job = $null
        Stop-Word $script:wordPid
    }
    Set-Failed "Unexpected error, details in $errorLog"
}

function Show-Overall {
    $lblCount.Text = "$script:done of $script:total done"
    Update-RowInfo
}

# Size and last modified time for each document: the PDF once it exists, otherwise the Word file
function Update-RowInfo {
    $culture = [Globalization.CultureInfo]::InvariantCulture
    for ($k = 0; $k -lt [math]::Min($script:total, $maxRows); $k++) {
        $pdf = Join-Path $pdfFolderPath ($script:wordFiles[$k].BaseName + ".pdf")
        if (Test-Path $pdf) { $kind = "PDF"; $item = Get-Item $pdf } else { $kind = "Word"; $item = $script:wordFiles[$k] }
        $script:rowInfo[$k].Text = "{0}   {1:N1} MB   |   {2}" -f $kind, ($item.Length / 1MB), $item.LastWriteTime.ToString("dd-MMM-yyyy HH:mm", $culture)
    }
}

function Stop-Word($processId) {
    if ($processId) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }
}

# Sets the status text and bar of one document row. Rows beyond the list are ignored.
function Set-Row($i, $text, $color, $value, $busy) {
    if ($i -lt 0 -or $i -ge $script:rowLabels.Count) { return }
    $script:rowLabels[$i].Text      = $text
    $script:rowLabels[$i].ForeColor = $color
    $script:rowBars[$i].Marquee     = $busy
    if (-not $busy) { $script:rowBars[$i].Value = $value }
}

function Get-RowIndex($baseName) {
    for ($k = 0; $k -lt [math]::Min($script:total, $maxRows); $k++) {
        if ($script:wordFiles[$k].BaseName -eq $baseName) { return $k }
    }
    return -1
}

function Reset-Rows {
    for ($k = 0; $k -lt [math]::Min($script:total, $maxRows); $k++) {
        Set-Row $k "Waiting" ([Theme]::Muted) 0 $false
    }
}

function Show-LastRun {
    $text = "never"
    if (Test-Path $lastRunFile) { $text = (Get-Content $lastRunFile -Raw).Trim() }
    $lblLastRun.Text = "Last run: $text"
}

function Save-LastRun {
    $stamp = (Get-Date).ToString("dd-MMM-yyyy HH:mm", [Globalization.CultureInfo]::InvariantCulture)
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $lastRunFile) | Out-Null
        Set-Content -Path $lastRunFile -Value $stamp -Encoding UTF8
    } catch {
        Write-Log "Could not save the last run time: $($_.Exception.Message)"
    }
    Show-LastRun
}

function Set-Finished($text, $color) {
    $timer.Stop()
    Save-LastRun
    for ($k = 0; $k -lt $script:rowBars.Count; $k++) { $script:rowBars[$k].Marquee = $false }
    $lblStatus.ForeColor = $color
    $lblStatus.Text      = $text
    $btnStart.Text       = "Run again"
    $btnStart.Enabled    = $true
}

function Set-Failed($text) {
    $script:state = "Failed"
    Write-Log $text
    Set-Finished "Stopped: see the log for details" ([Theme]::Bad)
}

function Set-Succeeded($text) {
    $script:state = "Finished"
    Write-Log $text
    Set-Finished $text ([Theme]::Good)
}

# Reads one line of upload output and moves the matching document row along
function Write-UploadLine($line) {
    $text = ("$line" -replace "`r", "").Trim()
    if (-not $text) { return }
    Write-Log $text

    if ($text -match "Unchanged: '(.+?)\.pdf'") {
        Set-Row (Get-RowIndex $matches[1]) "Unchanged" ([Theme]::Muted) 100 $false
    }
    if ($text -match "Upload failed for '(.+?)\.pdf'") {
        Set-Row (Get-RowIndex $matches[1]) "Upload failed" ([Theme]::Bad) 0 $false
    }
    if ($text -match "Uploading '(.+?)\.pdf'") {
        $script:uploadName = $matches[1]
        Set-Row (Get-RowIndex $script:uploadName) "Uploading" ([Theme]::Accent2) 0 $true
    }
    if ($text -match "UPLOADED:") {
        Set-Row (Get-RowIndex $script:uploadName) "Uploaded" ([Theme]::Good) 100 $false
    }
}

# Called by the Start button. Does nothing while a run is already going.
function Start-Run {
    if ($script:state -eq "Converting" -or $script:state -eq "Uploading") { return }
    if ($script:total -eq 0) {
        $lblStatus.ForeColor = [Theme]::Bad
        $lblStatus.Text = "No Word documents found in the folder."
        return
    }

    $script:next  = 0
    $script:done  = 0
    $script:state = "Converting"
    Reset-Rows
    Show-Overall
    $btnStart.Enabled   = $false
    $btnStart.Text      = "Running..."
    $lblStatus.ForeColor = [Theme]::Muted
    $lblStatus.Text     = "Converting $script:total document(s), then uploading to Vercel Blob..."
    Write-Log "Started."
    $timer.Start()
    Start-NextStep
}

# Starts the next Word export, skipping PDFs that are already up to date. Starts the upload when nothing is left.
function Start-NextStep {
    while ($script:next -lt $script:total) {
        $row     = $script:next
        $file    = $script:wordFiles[$row]
        $pdfPath = Join-Path -Path $pdfFolderPath -ChildPath ($file.BaseName + ".pdf")
        $script:next++

        if ((Test-Path $pdfPath) -and -not ($file.LastWriteTime -gt (Get-Item $pdfPath).LastWriteTime)) {
            $script:done++
            Write-Log "[$script:next/$script:total] Skipped '$($file.Name)' (PDF is already up to date)."
            Set-Row $row "Up to date" ([Theme]::Muted) 100 $false
            Show-Overall
            continue
        }

        # Start Word inside the background job
        $script:job = Start-Job -ScriptBlock {
            param($fileFullName, $pdfPath)
            try {
                $wordApp = New-Object -ComObject Word.Application
                $wordApp.Visible = $false
                $wordApp.DisplayAlerts = 0   # no prompts: a prompt would stop the job with nobody to answer it

                $wordPid = (Get-Process -Name "winword" | Sort-Object -Property StartTime -Descending | Select-Object -First 1).Id

                # Read-only, so a document already open in Word (for example, one being edited) does not block the export
                $doc = $wordApp.Documents.Open($fileFullName, $false, $true)
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
        $script:currentRow  = $row
        $script:currentName = $file.Name
        $script:currentPdf  = $pdfPath
        $script:jobStart    = Get-Date
        $script:lastSample  = $script:jobStart
        $script:lastCpu     = $null
        $script:idleSeconds = 0
        $script:wordPid     = $null

        Set-Row $row "Converting" ([Theme]::Accent2) 0 $true
        Write-Log "[$script:next/$script:total] Rendering '$($file.Name)'..."
        return
    }

    Start-Upload
}

function Start-Upload {
    # Documents that are ready for upload wait in the list until their upload line arrives
    for ($k = 0; $k -lt [math]::Min($script:total, $maxRows); $k++) {
        $text = $script:rowLabels[$k].Text
        if ($text -eq "Converted" -or $text -eq "Up to date") { Set-Row $k "Queued" ([Theme]::Muted) 100 $false }
    }

    $bookNames = @($script:wordFiles.BaseName)
    $script:job = Start-Job -FilePath $uploadScript -ArgumentList $pdfFolderPath, $bookNames
    $script:state      = "Uploading"
    $script:jobStart   = Get-Date
    $script:uploadName = ""
    $lblStatus.Text    = "Uploading PDFs to Vercel Blob..."
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
        Set-Row $script:currentRow "Failed" ([Theme]::Bad) 0 $false
        Set-Failed "Error processing '$($script:currentName)': $($jobError[0])"
        return
    }

    if (Test-Path $pdfPath) {
        $pdfInfo  = Get-Item $pdfPath
        $sizeInKB = [math]::Round(($pdfInfo.Length / 1KB), 2)
        Write-Log "SUCCESS: Completed '$($pdfInfo.Name)' ($sizeInKB KB) in $elapsed seconds."
    }
    Set-Row $script:currentRow "Converted" ([Theme]::Good) 100 $false
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
        Set-Row $script:currentRow "Failed" ([Theme]::Bad) 0 $false
        Set-Failed "CRITICAL ERROR: Background Word instance (PID: $($script:wordPid)) flatlined at 0% CPU while processing $($script:currentName) and was killed."
    }
}

function Complete-Upload {
    # Flush any log lines the upload wrote after the last tick
    Receive-Job $script:job -ErrorAction SilentlyContinue | ForEach-Object { Write-UploadLine $_ }
    $jobError  = $script:job.ChildJobs[0].Error
    $jobFailed = $script:job.State -eq "Failed"
    Remove-Job $script:job
    $script:job = $null

    # Rows still waiting were not uploaded (for example, no Vercel token)
    for ($k = 0; $k -lt [math]::Min($script:total, $maxRows); $k++) {
        $text = $script:rowLabels[$k].Text
        if ($text -eq "Queued" -or $text -eq "Uploading") { Set-Row $k "Not uploaded" ([Theme]::Muted) 100 $false }
    }

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
            $lblStatus.Text = "Converting '$($script:currentName)'  -  $([int]((Get-Date) - $script:jobStart).TotalSeconds)s"
            Test-WordHealth
        } else {
            Complete-Conversion
        }
    }
    elseif ($script:state -eq "Uploading") {
        # Show the upload progress row by row while it runs
        Receive-Job $script:job -ErrorAction SilentlyContinue | ForEach-Object { Write-UploadLine $_ }
        if ($script:job.State -eq "Running") {
            $lblStatus.Text = "Uploading to Vercel Blob  -  $([int]((Get-Date) - $script:jobStart).TotalSeconds)s"
        } else {
            Complete-Upload
        }
    }
}

# ---------- Wiring ----------

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    try { Update-Progress } catch { Write-ErrorReport $_.Exception }
})

$form.Add_Shown({
    Show-Overall
    Show-LastRun
    $lblStatus.Text = "Ready: $script:total document(s). Press Start to convert and upload."
    Write-Log "Books to PDF: $script:total Word document(s) in '$wordFolderPath'"
    if ($script:total -gt $maxRows) {
        Write-Log "The list shows the first $maxRows documents; all $script:total are still processed."
    }
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

# Demo mode: define the fake jobs again, so they replace the real Test-WordHealth defined above
if ($Demo) { . (Join-Path $PSScriptRoot "Demo Fake Jobs.ps1") }

[void]$form.ShowDialog()
$timer.Dispose()
$form.Dispose()
