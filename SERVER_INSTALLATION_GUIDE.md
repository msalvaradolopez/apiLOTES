# Guía de Instalación - Runtime para Windows Server

## Requisitos del Sistema
- Windows Server 2012 R2 o superior
- Windows Server 2016, 2019, o 2022 (recomendado)

## Pasos de Instalación

### 1. Instalar .NET Framework 4.8

**Opción A: Descarga Manual**
1. Visitar: https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48
2. Descargar "Runtime" (NO Developer Pack)
3. Ejecutar el instalador como Administrador

**Opción B: Usar PowerShell (Recomendado)**
```powershell
# Ejecutar como Administrador
.\install-dotnet48-runtime.ps1
```

### 2. Habilitar IIS y ASP.NET

```powershell
# Ejecutar como Administrador en PowerShell

# Instalar IIS con características necesarias
Install-WindowsFeature -Name Web-Server -IncludeManagementTools
Install-WindowsFeature -Name Web-Asp-Net45
Install-WindowsFeature -Name Web-Net-Ext45
Install-WindowsFeature -Name Web-ISAPI-Ext
Install-WindowsFeature -Name Web-ISAPI-Filter
Install-WindowsFeature -Name Web-Http-Redirect
Install-WindowsFeature -Name Web-WebSockets

# Verificar instalación
Get-WindowsFeature -Name Web-* | Where-Object {$_.Installed -eq $true}
```

### 3. Registrar ASP.NET con IIS

```powershell
# Ejecutar como Administrador
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\aspnet_regiis.exe -i
```

### 4. Verificar Instalación

```powershell
# Verificar versión de .NET Framework instalada
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" | Select-Object Version, Release
```

El número de Release para .NET 4.8 debe ser **528040** o superior.

### 5. Configurar Application Pool en IIS

Después de publicar la aplicación:
1. Abrir IIS Manager
2. Crear o configurar Application Pool:
   - .NET CLR Version: **v4.0**
   - Managed Pipeline Mode: **Integrated**
   - Identity: **ApplicationPoolIdentity** (o cuenta específica)

### 6. Publicar Aplicación

Copiar archivos publicados a:
```
C:\inetpub\wwwroot\apiLOTES\
```

O usar la ubicación de tu preferencia y configurar el sitio en IIS.

## Verificación Final

```powershell
# Verificar que IIS está corriendo
Get-Service -Name W3SVC

# Probar la aplicación
# Navegar a: http://localhost/apiLOTES
```

## Notas Importantes

- **REINICIAR** el servidor después de instalar .NET Framework 4.8
- La aplicación incluye todas las DLLs necesarias (BouncyCastle, iTextSharp, etc.)
- No se requiere .NET Core o .NET 5+
- Verificar firewall para permitir tráfico HTTP/HTTPS

## Solución de Problemas

### Error: "Could not load file or assembly"
```powershell
# Verificar que ASP.NET está registrado
aspnet_regiis.exe -i -enable
```

### Error: HTTP 500.19
- Verificar permisos de carpeta
- El Application Pool debe tener acceso de lectura

### Verificar permisos
```powershell
icacls "C:\inetpub\wwwroot\apiLOTES" /grant "IIS AppPool\DefaultAppPool:(OI)(CI)RX"
```
