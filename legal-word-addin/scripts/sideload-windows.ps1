# Copy manifest into Word's Wef folder (Windows desktop Word).
# After: restart Word, open a document, check Insert (or Home) for the add-in group.
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Manifest = Join-Path $Root "manifest.xml"
if (-not (Test-Path $Manifest)) {
  Write-Error "Missing manifest.xml at $Manifest"
}
$Wef = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
New-Item -ItemType Directory -Force -Path $Wef | Out-Null
$Dest = Join-Path $Wef "legal-ai-assistant.xml"
Copy-Item -LiteralPath $Manifest -Destination $Dest -Force
Write-Host "Copied manifest to:"
Write-Host "  $Dest"
Write-Host ""
Write-Host "1) Close Word completely."
Write-Host "2) Run: npm run dev:all   (HTTPS on https://localhost:3000)"
Write-Host "3) Open Word, open a real document, look for the add-in on the ribbon."
Write-Host ""
Write-Host "If nothing appears, use Insert > Add-ins > My Add-ins > Upload My Add-in and pick manifest.xml from the project folder."
