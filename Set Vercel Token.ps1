# Run once: saves the Vercel Blob read-write token DPAPI-encrypted, outside this repository.
# Only this Windows user on this laptop can decrypt it.
$tokenPath = "$HOME\.secrets\vercel-blob-isvara-media.xml"

Write-Host "Find the token in Vercel > Storage > isvara-media > .env.local tab (BLOB_READ_WRITE_TOKEN)."
$secure = Read-Host "Paste BLOB_READ_WRITE_TOKEN (input hidden)" -AsSecureString

New-Item -ItemType Directory -Force (Split-Path $tokenPath) | Out-Null
New-Object System.Management.Automation.PSCredential("isvara-media", $secure) | Export-Clixml $tokenPath

Write-Host "Saved to $tokenPath" -ForegroundColor Green
