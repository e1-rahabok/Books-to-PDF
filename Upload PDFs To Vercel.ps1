# Uploads the book PDFs to the Vercel Blob store, replacing older copies.
# The read-write token is stored DPAPI-encrypted outside the repository (see "Set Vercel Token.ps1").
param(
    [string]$PdfFolderPath = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide",
    [string[]]$BookNames,                     # PDF base names to consider; defaults to every PDF in the folder
    [string]$BlobPrefix    = "downloads/",
    [string]$TokenPath     = "$HOME\.secrets\vercel-blob-isvara-media.xml"
)

$apiUrl     = "https://blob.vercel-storage.com"
$apiVersion = "11"

if (-not (Test-Path $TokenPath)) {
    Write-Host "Vercel upload skipped: no token at '$TokenPath'. Run 'Set Vercel Token.ps1' once to add it." -ForegroundColor Yellow
    return
}
$token   = (Import-Clixml $TokenPath).GetNetworkCredential().Password
$headers = @{ authorization = "Bearer $token"; "x-api-version" = $apiVersion }

# What is already in the store, keyed by pathname
$remote = @{}
$cursor = $null
do {
    $query = "prefix=$([uri]::EscapeDataString($BlobPrefix))&limit=1000"
    if ($cursor) { $query += "&cursor=$([uri]::EscapeDataString($cursor))" }
    $page = Invoke-RestMethod -Method Get -Uri "$apiUrl/?$query" -Headers $headers
    foreach ($blob in $page.blobs) { $remote[$blob.pathname] = $blob }
    $cursor = $page.cursor
} while ($page.hasMore)

$pdfFiles = @(Get-ChildItem -Path $PdfFolderPath -Filter "*.pdf" |
              Where-Object { -not $BookNames -or $BookNames -contains $_.BaseName })
$total = $pdfFiles.Count
$index = 0

Write-Host "Vercel Blob: checking $total PDF(s) against '$BlobPrefix'" -ForegroundColor Cyan

foreach ($pdf in $pdfFiles) {
    $index++
    $pathname = $BlobPrefix + $pdf.Name
    $existing = $remote[$pathname]

    if ($existing -and ([datetime]$existing.uploadedAt).ToUniversalTime() -ge $pdf.LastWriteTimeUtc) {
        Write-Host "[$index/$total] Skipped upload: '$($pdf.Name)' (Vercel copy is up to date)."
        continue
    }

    Write-Host -NoNewline "[$index/$total] Uploading '$($pdf.Name)' ($([math]::Round($pdf.Length / 1MB, 1)) MB)..." -ForegroundColor Yellow
    $putHeaders = $headers + @{
        "x-vercel-blob-access"    = "public"
        "x-content-type"          = "application/pdf"
        "x-add-random-suffix"     = "0"
        "x-allow-overwrite"       = "1"
        "x-cache-control-max-age" = "60"   # so readers get a replaced PDF within a minute
    }
    try {
        $result = Invoke-RestMethod -Method Put -Uri "$apiUrl/?pathname=$([uri]::EscapeDataString($pathname))" `
                                    -Headers $putHeaders -InFile $pdf.FullName -ContentType "application/pdf"
        Write-Host "`r[$index/$total] UPLOADED: $($result.url)" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "Upload failed for '$($pdf.Name)': $($_.Exception.Message) $($_.ErrorDetails.Message)" -ForegroundColor Red
        throw
    }
}
