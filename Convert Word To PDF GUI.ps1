# Graphical version of "Convert Word To PDF.ps1".
# Same conversion, hang detection and Vercel upload, shown in a dark window with progress ring, bars and a log.
# Runs in Windows PowerShell 5.1 (powershell.exe), which starts in STA mode as WinForms needs.

# Define the folder paths
$wordFolderPath   = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide\Editable format"
$pdfFolderPath    = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide"
$uploadScript     = Join-Path $PSScriptRoot "Upload PDFs To Vercel.ps1"
$hangLimitSeconds = 45   # Word at 0% CPU for this long is treated as hung

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Plain WinForms cannot draw rounded corners, gradients or a progress ring, so the look is built here.
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

// Circular progress with a big number in the middle
public class Ring : Control {
    int max = 100;
    int val;
    string caption = "";
    string subCaption = "";

    public Ring() {
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }

    public int Maximum { get { return max; } set { max = Math.Max(1, value); Invalidate(); } }
    public int Value { get { return val; } set { val = Math.Max(0, Math.Min(max, value)); Invalidate(); } }
    public string Caption { get { return caption; } set { caption = value; Invalidate(); } }
    public string SubCaption { get { return subCaption; } set { subCaption = value; Invalidate(); } }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        float thick = 14;
        float size = Math.Min(Width, Height) - 20;
        RectangleF arc = new RectangleF((Width - size) / 2f + thick / 2f, (Height - size) / 2f + thick / 2f, size - thick, size - thick);

        using (Pen track = new Pen(Theme.Track, thick)) {
            g.DrawEllipse(track, arc);
        }

        float sweep = 360f * val / max;
        if (sweep > 0.5f) {
            using (Pen glow = new Pen(Color.FromArgb(70, Theme.Accent), thick + 10)) {
                glow.StartCap = LineCap.Round;
                glow.EndCap = LineCap.Round;
                g.DrawArc(glow, arc, -90, sweep);
            }
            using (LinearGradientBrush lb = new LinearGradientBrush(new PointF(0, 0), new PointF(Width, Height), Theme.Accent2, Theme.Accent))
            using (Pen p = new Pen(lb, thick)) {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                g.DrawArc(p, arc, -90, sweep);
            }
        }

        using (Font big = new Font("Segoe UI Variable Display", 28f, FontStyle.Bold))
        using (Font small = new Font("Segoe UI", 10f))
        using (SolidBrush text = new SolidBrush(Theme.Text))
        using (SolidBrush muted = new SolidBrush(Theme.Muted))
        using (StringFormat sf = new StringFormat()) {
            sf.Alignment = StringAlignment.Center;
            sf.LineAlignment = StringAlignment.Center;
            g.DrawString(caption, big, text, new RectangleF(0, Height * 0.22f, Width, Height * 0.4f), sf);
            g.DrawString(subCaption, small, muted, new RectangleF(0, Height * 0.58f, Width, Height * 0.2f), sf);
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

// Pill-shaped gradient button with a hover highlight
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

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        using (GraphicsPath p = Theme.RoundRect(new RectangleF(0, 0, Width - 1, Height - 1), Height / 2f))
        using (LinearGradientBrush lb = new LinearGradientBrush(new PointF(0, 0), new PointF(Width, 0),
                   hover ? Theme.Accent2 : Theme.Accent,
                   hover ? Color.FromArgb(170, 225, 255) : Theme.Accent2)) {
            g.FillPath(lb, p);
        }

        using (Font f = new Font("Segoe UI Semibold", 10.5f))
        using (SolidBrush tb = new SolidBrush(ForeColor))
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
$script:currentName = ""
$script:currentPdf  = ""

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
$fontFile    = New-Object System.Drawing.Font("Segoe UI Semibold", 12)

$form = New-Object AppForm
$form.Text            = "Books to PDF"
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.ClientSize      = New-Object System.Drawing.Size(460, 712)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false

# Header
$lblTitle    = New-Label $form "Books to PDF" 28 20 400 36 (New-Object System.Drawing.Font("Segoe UI Variable Display", 18, [System.Drawing.FontStyle]::Bold)) ([Theme]::Text)
$lblSubtitle = New-Label $form "The Guide  -  Word to PDF" 28 58 400 22 $null ([Theme]::Muted)

# Overall progress ring
$ring = New-Object Ring
$ring.SetBounds(130, 96, 200, 200)
$ring.Maximum = [math]::Max($script:total, 1)
$form.Controls.Add($ring)

# Current step card
$cardCurrent = New-Object Card
$cardCurrent.SetBounds(24, 320, 412, 124)
$form.Controls.Add($cardCurrent)

$lblSection1 = New-Label $cardCurrent "CURRENT FILE" 20 14 200 20 $fontSection ([Theme]::Muted)
$lblBadge    = New-Label $cardCurrent "READY" 252 14 140 20 $fontSection ([Theme]::Muted)
$lblBadge.TextAlign = "MiddleRight"
$lblCurrent  = New-Label $cardCurrent "Starting..." 20 38 372 26 $fontFile ([Theme]::Text)

$barCurrent = New-Object GradientBar
$barCurrent.SetBounds(20, 72, 372, 8)
$cardCurrent.Controls.Add($barCurrent)

$lblElapsed = New-Label $cardCurrent "" 20 92 372 20 $null ([Theme]::Muted)

# Log card
$cardLog = New-Object Card
$cardLog.SetBounds(24, 462, 412, 174)
$form.Controls.Add($cardLog)

$lblSection2 = New-Label $cardLog "LOG" 20 14 200 20 $fontSection ([Theme]::Muted)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.SetBounds(18, 40, 376, 122)
$txtLog.Multiline   = $true
$txtLog.ReadOnly    = $true
$txtLog.BorderStyle = "None"
$txtLog.BackColor   = [Theme]::Card
$txtLog.ForeColor   = [Theme]::Text
$txtLog.ScrollBars  = "Vertical"
$txtLog.Font        = New-Object System.Drawing.Font("Consolas", 9)
$cardLog.Controls.Add($txtLog)

# Close button
$btnClose = New-Object PillButton
$btnClose.Text = "Close"
$btnClose.SetBounds(284, 652, 152, 40)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# ---------- Helpers ----------

function Write-Log($text) {
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format "HH:mm:ss"), $text))
}

function Show-Overall {
    $ring.Value      = $script:done
    $ring.Caption    = "$script:done / $script:total"
    $ring.SubCaption = "documents"
}

function Set-Badge($text, $color) {
    $lblBadge.Text      = $text
    $lblBadge.ForeColor = $color
}

function Stop-Word($processId) {
    if ($processId) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }
}

function Set-Finished($text, $badge, $color) {
    $timer.Stop()
    $barCurrent.Marquee = $false
    $barCurrent.Value   = 0
    $lblCurrent.ForeColor = $color
    $lblCurrent.Text = $text
    $lblElapsed.Text = ""
    Set-Badge $badge $color
    Write-Log $text
}

function Set-Failed($text) {
    $script:state = "Failed"
    Set-Finished $text "FAILED" ([Theme]::Bad)
}

function Set-Succeeded($text) {
    $script:state = "Finished"
    Set-Finished $text "DONE" ([Theme]::Good)
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

        $lblCurrent.ForeColor = [Theme]::Text
        $lblCurrent.Text = "Rendering '$($file.Name)' ($script:next of $script:total)"
        Set-Badge "RENDERING" ([Theme]::Accent2)
        $barCurrent.Marquee = $true
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
    $lblCurrent.ForeColor = [Theme]::Text
    $lblCurrent.Text   = "Uploading PDFs to Vercel Blob..."
    Set-Badge "UPLOADING" ([Theme]::Accent2)
    $barCurrent.Marquee = $true
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
            $lblElapsed.Text = "Elapsed: $([int]((Get-Date) - $script:jobStart).TotalSeconds)s"
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
            $lblElapsed.Text = "Elapsed: $([int]((Get-Date) - $script:jobStart).TotalSeconds)s"
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
    Show-Overall
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
