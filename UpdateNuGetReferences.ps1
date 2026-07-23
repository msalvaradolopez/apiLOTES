# PowerShell script to update NuGet package references in lotesAPI.vbproj
# Close Visual Studio before running this script

Write-Host "Updating NuGet package references..." -ForegroundColor Green

$projectFile = "lotesAPI.vbproj"

if (-not (Test-Path $projectFile)) {
    Write-Host "Error: $projectFile not found!" -ForegroundColor Red
    exit 1
}

# Read the project file
$content = Get-Content $projectFile -Raw

# Update Microsoft.CodeDom.Providers.DotNetCompilerPlatform references
$content = $content -replace 'Microsoft\.CodeDom\.Providers\.DotNetCompilerPlatform\.2\.0\.1\\build\\net45', 'Microsoft.CodeDom.Providers.DotNetCompilerPlatform.4.1.0\build\net48'
$content = $content -replace 'Microsoft\.CodeDom\.Providers\.DotNetCompilerPlatform\.2\.0\.1\\lib\\net45', 'Microsoft.CodeDom.Providers.DotNetCompilerPlatform.4.1.0\lib\net472'

# Update System.Web.Cors
$content = $content -replace 'System\.Web\.Cors, Version=5\.2\.7\.0', 'System.Web.Cors, Version=5.3.0.0'
$content = $content -replace 'Microsoft\.AspNet\.Cors\.5\.2\.7\\lib', 'Microsoft.AspNet.Cors.5.3.0\lib'

# Update System.Web.Http.Cors
$content = $content -replace 'System\.Web\.Http\.Cors, Version=5\.2\.7\.0', 'System.Web.Http.Cors, Version=5.3.0.0'
$content = $content -replace 'Microsoft\.AspNet\.WebApi\.Cors\.5\.2\.7\\lib', 'Microsoft.AspNet.WebApi.Cors.5.3.0\lib'

# Update Microsoft.Web.Infrastructure
$content = $content -replace 'Microsoft\.Web\.Infrastructure, Version=1\.0\.0\.0', 'Microsoft.Web.Infrastructure, Version=2.0.0.0'
$content = $content -replace 'Microsoft\.Web\.Infrastructure\.1\.0\.0\.0\\lib', 'Microsoft.Web.Infrastructure.2.0.0\lib'

# Update Newtonsoft.Json
$content = $content -replace 'Newtonsoft\.Json\.12\.0\.2\\lib', 'Newtonsoft.Json.13.0.3\lib'
$content = $content -replace '<HintPath>packages\\Newtonsoft\.Json\.12\.0\.2', '<HintPath>packages\Newtonsoft.Json.13.0.3'

# Update System.Net.Http.Formatting
$content = $content -replace 'System\.Net\.Http\.Formatting, Version=5\.2\.7\.0', 'System.Net.Http.Formatting, Version=6.0.0.0'
$content = $content -replace 'Microsoft\.AspNet\.WebApi\.Client\.5\.2\.7\\lib', 'Microsoft.AspNet.WebApi.Client.6.0.0\lib'

# Update System.Web.Http
$content = $content -replace 'System\.Web\.Http, Version=5\.2\.7\.0', 'System.Web.Http, Version=5.3.0.0'
$content = $content -replace 'Microsoft\.AspNet\.WebApi\.Core\.5\.2\.7\\lib', 'Microsoft.AspNet.WebApi.Core.5.3.0\lib'

# Update System.Web.Http.WebHost
$content = $content -replace 'System\.Web\.Http\.WebHost, Version=5\.2\.7\.0', 'System.Web.Http.WebHost, Version=5.3.0.0'
$content = $content -replace 'Microsoft\.AspNet\.WebApi\.WebHost\.5\.2\.7\\lib', 'Microsoft.AspNet.WebApi.WebHost.5.3.0\lib'

# Update System.Web.Mvc
$content = $content -replace 'System\.Web\.Mvc, Version=5\.2\.7\.0', 'System.Web.Mvc, Version=5.3.0.0'
$content = $content -replace 'Microsoft\.AspNet\.Mvc\.5\.2\.7\\lib', 'Microsoft.AspNet.Mvc.5.3.0\lib'

# Update System.Web.Razor
$content = $content -replace 'Microsoft\.AspNet\.Razor\.3\.2\.7\\lib', 'Microsoft.AspNet.Razor.3.3.0\lib'

# Update System.Web.WebPages, Helpers, Deployment, Razor (keep Version=3.0.0.0 but update path to 3.3.0)
$content = $content -replace 'Microsoft\.AspNet\.WebPages\.3\.2\.7\\lib', 'Microsoft.AspNet.WebPages.3.3.0\lib'

# Add new package references for System.Buffers, System.Memory, etc.
$newReferences = @"
    <Reference Include="System.Buffers, Version=4.0.3.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Buffers.4.5.1\lib\net461\System.Buffers.dll</HintPath>
    </Reference>
    <Reference Include="System.Memory, Version=4.0.1.2, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Memory.4.5.5\lib\net461\System.Memory.dll</HintPath>
    </Reference>
    <Reference Include="System.Numerics" />
    <Reference Include="System.Numerics.Vectors, Version=4.1.4.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Numerics.Vectors.4.5.0\lib\net46\System.Numerics.Vectors.dll</HintPath>
    </Reference>
    <Reference Include="System.Runtime.CompilerServices.Unsafe, Version=6.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Runtime.CompilerServices.Unsafe.6.0.0\lib\net461\System.Runtime.CompilerServices.Unsafe.dll</HintPath>
    </Reference>
    <Reference Include="System.Threading.Tasks.Extensions, Version=4.2.0.1, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Threading.Tasks.Extensions.4.5.4\lib\net461\System.Threading.Tasks.Extensions.dll</HintPath>
    </Reference>
    <Reference Include="System.Xml.Linq" />
"@

# Find the location to insert new references (before the closing </ItemGroup> after Antlr3.Runtime)
$insertionPoint = $content.IndexOf("</ItemGroup>", $content.IndexOf("Antlr3.Runtime"))
if ($insertionPoint -gt 0) {
    $content = $content.Insert($insertionPoint, $newReferences + "`r`n  ")
}

# Save the updated project file
Set-Content -Path $projectFile -Value $content -Encoding UTF8

Write-Host "Project file updated successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Delete the packages folder: Remove-Item -Recurse -Force .\packages\" -ForegroundColor White
Write-Host "2. Restore packages: dotnet restore" -ForegroundColor White
Write-Host "3. Open Visual Studio and rebuild the solution" -ForegroundColor White
