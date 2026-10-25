# Uploads the book PDFs whose content changed to the Vercel Blob store, replacing any file with the same name.
# The read-write token is stored DPAPI-encrypted outside the repository (see "Set Vercel Token.ps1").
param(
    [string]$PdfFolderPath = "C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide",
    [string[]]$BookNames,                     # PDF base names to consider; defaults to every PDF in the folder
    [string]$BlobPrefix    = "downloads/",
    [string]$TokenPath     = "$HOME\.secrets\vercel-blob-isvara-media.xml",
    [string]$StatePath     = "$HOME\.secrets\vercel-upload-state.json"   # SHA-256 of each PDF at its last upload
)

$apiUrl     = "https://blob.vercel-storage.com"
$apiVersion = "11"

if (-not (Test-Path $TokenPath)) {
    Write-Host "Vercel upload skipped: no token at '$TokenPath'. Run 'Set Vercel Token.ps1' once to add it." -ForegroundColor Yellow
    return
}
$token   = (Import-Clixml $TokenPath).GetNetworkCredential().Password
$headers = @{ authorization = "Bearer $token"; "x-api-version" = $apiVersion }

$pdfFiles = @(Get-ChildItem -Path $PdfFolderPath -Filter "*.pdf" |
              Where-Object { -not $BookNames -or $BookNames -contains $_.BaseName })
$total = $pdfFiles.Count
$index = 0

# Hashes saved at the last successful upload, keyed by file name
$uploadedHashes = @{}
if (Test-Path $StatePath) {
    (Get-Content $StatePath -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $uploadedHashes[$_.Name] = $_.Value }
}

Write-Host "Vercel Blob: checking $total PDF(s) for changes, uploading only changed files to '$BlobPrefix'" -ForegroundColor Cyan

foreach ($pdf in $pdfFiles) {
    $index++
    $pathname = $BlobPrefix + $pdf.Name

    $hash = (Get-FileHash -Path $pdf.FullName -Algorithm SHA256).Hash
    if ($uploadedHashes[$pdf.Name] -eq $hash) {
        Write-Host "[$index/$total] Unchanged: '$($pdf.Name)' (not uploaded)" -ForegroundColor DarkGray
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
        $uploadedHashes[$pdf.Name] = $hash
        $uploadedHashes | ConvertTo-Json | Set-Content -Path $StatePath -Encoding UTF8
    } catch {
        Write-Host ""
        Write-Host "Upload failed for '$($pdf.Name)': $($_.Exception.Message) $($_.ErrorDetails.Message)" -ForegroundColor Red
        throw
    }
}
