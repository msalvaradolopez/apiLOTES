# Fix conexion project references
Write-Host "Fixing conexion project references..." -ForegroundColor Cyan

$projectFile = ".\conexion\conexion.vbproj"
$content = Get-Content $projectFile -Raw

# Update iTextSharp reference from 5.5.13.3 to 5.5.13.4
$content = $content -replace 'iTextSharp.5.5.13.3', 'iTextSharp.5.5.13.4'
$content = $content -replace 'itextsharp, Version=5.5.13.3', 'itextsharp, Version=5.5.13.4'

# Update BouncyCastle reference
$content = $content -replace 'BouncyCastle.Crypto, Version=1.8.9.0, Culture=neutral, PublicKeyToken=0e99375e54769942', 'BouncyCastle.Cryptography, Version=2.4.0.0, Culture=neutral, PublicKeyToken=072edcf4a5328938, processorArchitecture=MSIL'
$content = $content -replace 'packages\\BouncyCastle.1.8.9\\lib\\BouncyCastle.Crypto.dll', 'packages\BouncyCastle.Cryptography.2.4.0\lib\net461\BouncyCastle.Cryptography.dll'

# Update iTextSharp path to use net461
$content = $content -replace 'packages\\iTextSharp.5.5.13.[34]\\lib\\itextsharp.dll', 'packages\iTextSharp.5.5.13.4\lib\net461\itextsharp.dll'

Set-Content $projectFile -Value $content -NoNewline

Write-Host "Project file updated successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Please close Visual Studio and run this script, then reopen the solution." -ForegroundColor Yellow
