# Script para descargar e instalar .NET Framework 4.8 Runtime en Windows Server
# Ejecutar como Administrador

$downloadUrl = "https://go.microsoft.com/fwlink/?linkid=2088631"
$installerPath = "$env:TEMP\ndp48-web.exe"

Write-Host "Descargando .NET Framework 4.8 Runtime..." -ForegroundColor Green
Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

Write-Host "Instalando .NET Framework 4.8..." -ForegroundColor Green
Start-Process -FilePath $installerPath -ArgumentList "/q /norestart" -Wait

Write-Host "Instalacion completada. Se recomienda reiniciar el servidor." -ForegroundColor Yellow
