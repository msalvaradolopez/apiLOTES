# NuGet Package Update - Complete Refactoring Guide

## Current Status
The packages.config file has been updated with the latest NuGet package versions, but the .vbproj file still references the old versions. This mismatch is causing build errors.

## Packages Updated

### Main Packages (lotesAPI)
- **jQuery**: 3.4.1 ? 3.7.1
- **Microsoft.AspNet.Cors**: 5.2.7 ? 5.3.0
- **Microsoft.AspNet.Mvc**: 5.2.7 ? 5.3.0
- **Microsoft.AspNet.Razor**: 3.2.7 ? 3.3.0
- **Microsoft.AspNet.WebApi**: 5.2.7 ? 5.3.0
- **Microsoft.AspNet.WebApi.Client**: 5.2.7 ? 6.0.0
- **Microsoft.AspNet.WebApi.Core**: 5.2.7 ? 5.3.0
- **Microsoft.AspNet.WebApi.Cors**: 5.2.7 ? 5.3.0
- **Microsoft.AspNet.WebApi.WebHost**: 5.2.7 ? 5.3.0
- **Microsoft.AspNet.WebPages**: 3.2.7 ? 3.3.0
- **Microsoft.CodeDom.Providers.DotNetCompilerPlatform**: 2.0.1 ? 4.1.0
- **Microsoft.Web.Infrastructure**: 1.0.0.0 ? 2.0.0
- **Newtonsoft.Json**: 12.0.2 ? 13.0.3

### New Dependencies Added
- System.Buffers 4.5.1
- System.Memory 4.5.5
- System.Numerics.Vectors 4.5.0
- System.Runtime.CompilerServices.Unsafe 6.0.0
- System.Threading.Tasks.Extensions 4.5.4

### Spanish Localization Packages (kept at 5.2.7)
- Microsoft.AspNet.Mvc.es
- Microsoft.AspNet.Razor.es
- Microsoft.AspNet.WebApi.Client.es
- Microsoft.AspNet.WebApi.Core.es
- Microsoft.AspNet.WebApi.WebHost.es
- Microsoft.AspNet.WebPages.es

### Conexion Project
- **BouncyCastle**: 1.8.9 ? BouncyCastle.Cryptography 2.4.0
- **iTextSharp**: 5.5.13.3 ? 5.5.13.4

## Solution Steps

### Option 1: Automated Update (Recommended)

1. **Close Visual Studio completely**

2. **Run the PowerShell script:**
   ```powershell
   cd C:\Salvador\apiLOTES
   .\UpdateNuGetReferences.ps1
   ```

3. **Clean and restore packages:**
   ```powershell
   Remove-Item -Recurse -Force .\packages\
   dotnet restore
   ```

4. **Open Visual Studio and rebuild the solution**

### Option 2: Manual Update

If you prefer to update manually or the script doesn't work:

1. **Close Visual Studio**

2. **Open lotesAPI.vbproj in a text editor**

3. **Update Line 3:**
   ```xml
   FROM: <Import Project="packages\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.2.0.1\build\net45\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.props" ...
   TO:   <Import Project="packages\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.4.1.0\build\net48\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.props" ...
   ```

4. **Update all assembly references (lines 55-130):**
   - System.Web.Cors: Change `5.2.7.0` ? `5.3.0.0` and path `5.2.7` ? `5.3.0`
   - System.Web.Http.Cors: Change `5.2.7.0` ? `5.3.0.0` and path `5.2.7` ? `5.3.0`
   - Microsoft.Web.Infrastructure: Change `1.0.0.0` ? `2.0.0.0` and path `1.0.0.0` ? `2.0.0`
   - Newtonsoft.Json: Change path `12.0.2` ? `13.0.3`
   - System.Net.Http.Formatting: Change `5.2.7.0` ? `6.0.0.0` and path `5.2.7` ? `6.0.0`
   - System.Web.Http: Change `5.2.7.0` ? `5.3.0.0` and path `5.2.7` ? `5.3.0`
   - System.Web.Http.WebHost: Change `5.2.7.0` ? `5.3.0.0` and path `5.2.7` ? `5.3.0`
   - System.Web.Mvc: Change `5.2.7.0` ? `5.3.0.0` and path `5.2.7` ? `5.3.0`
   - System.Web.Razor: Change path `3.2.7` ? `3.3.0`
   - System.Web.WebPages (all): Change path `3.2.7` ? `3.3.0`
   - Microsoft.CodeDom.Providers.DotNetCompilerPlatform: Change path `2.0.1\lib\net45` ? `4.1.0\lib\net472`

5. **Add new references before the closing `</ItemGroup>` tag (after Antlr3.Runtime):**
   ```xml
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
   ```

6. **Update line 351 (near the end):**
   ```xml
   FROM: <Error Condition="!Exists('packages\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.2.0.1\build\net45\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.props')" ...
   TO:   <Error Condition="!Exists('packages\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.4.1.0\build\net48\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.props')" ...
   ```

7. **Save the file**

8. **Clean and restore:**
   ```powershell
   cd C:\Salvador\apiLOTES
   Remove-Item -Recurse -Force .\packages\
   dotnet restore
   ```

9. **Open Visual Studio and rebuild**

## Files Already Updated
? packages.config (main project)
? conexion\packages.config
? Web.config (binding redirects and compiler configuration)

## Files That Need Update
? lotesAPI.vbproj (use script or manual update above)

## Expected Results After Update
- All NuGet packages will be at their latest compatible versions
- Solution will build without errors
- No code refactoring required - only package updates
- All functionality remains the same

## Troubleshooting

### If you still get XDocument errors:
1. Ensure System.Xml.Linq reference is added (included in the script)
2. Clean the solution (Build ? Clean Solution)
3. Close Visual Studio
4. Delete bin and obj folders
5. Reopen and rebuild

### If you get version conflict errors:
1. Check Web.config has all binding redirects (already updated)
2. Ensure packages folder is completely deleted before restore
3. Run: `dotnet restore --force`

### If packages fail to restore:
1. Clear NuGet cache: `dotnet nuget locals all --clear`
2. Restore again: `dotnet restore`

## Next Steps After Successful Build
1. Test the application to ensure everything works
2. Commit the changes to Git
3. Update any documentation regarding package versions
