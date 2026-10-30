<#
.SYNOPSIS
Creates the installation package for Windows:
 Builds the .Net library and the console runner with the .NET SDK
 Exports the 32bit and 64bit type libraries
 Generates the COM registration file for InnoSetup (SeleniumBasicSetup.pas)
 Creates the setup installer with InnoSetup 6 [https://jrsoftware.org/isinfo.php]

.PARAMETER Version
Optional new version (ex: 3.0.1.0) written to Selenium\Properties\AssemblyInfo.cs

.EXAMPLE
powershell -ExecutionPolicy Bypass -File build-setup.ps1
#>
param(
    [string]$Version,
    # Internal: runs the type library export and registration steps in a child process
    [string]$Export
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$ReleaseDir = Join-Path $PSScriptRoot 'Selenium\bin\Release'
$AssemblyPath = Join-Path $ReleaseDir 'Selenium.dll'

$SetupToolsSource = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

public static class SetupTools {

    [ComImport, Guid("00020406-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeLib {
        // Only SaveAllChanges is called, the other slots just keep the vtable order
        void CreateTypeInfo();
        void SetName();
        void SetVersion();
        void SetGuid();
        void SetDocString();
        void SetHelpFileName();
        void SetHelpContext();
        void SetLcid();
        void SetLibFlags();
        void SaveAllChanges();
    }

    [DllImport("oleaut32.dll", PreserveSig = false)]
    static extern ITypeLib LoadRegTypeLib(ref Guid rguid, short wVerMajor, short wVerMinor, int lcid);

    [DllImport("oleaut32.dll", PreserveSig = false, CharSet = CharSet.Unicode)]
    static extern ITypeLib LoadTypeLibEx(string szFile, int regkind);

    class ExporterSink : ITypeLibExporterNotifySink {
        public void ReportEvent(ExporterEventKind eventKind, int eventCode, string eventMsg) {
            if (eventKind != ExporterEventKind.NOTIF_TYPECONVERTED)
                Console.WriteLine("  " + eventMsg);
        }

        // Referenced assemblies (mscorlib...) resolve to their registered type library
        public object ResolveRef(Assembly assembly) {
            Guid guid = Marshal.GetTypeLibGuidForAssembly(assembly);
            int major, minor;
            Marshal.GetTypeLibVersionForAssembly(assembly, out major, out minor);
            try {
                return LoadRegTypeLib(ref guid, (short)major, (short)minor, 0);
            } catch (COMException) {
                string path = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), assembly.GetName().Name + ".tlb");
                return LoadTypeLibEx(path, 2 /* REGKIND_NONE */);
            }
        }
    }

    public static void ExportTypeLib(string assemblyPath, string tlbPath, bool win64) {
        Assembly assembly = Assembly.LoadFrom(assemblyPath);
        var flags = win64 ? TypeLibExporterFlags.ExportAs64Bit : TypeLibExporterFlags.ExportAs32Bit;
        var tlb = (ICreateTypeLib)new TypeLibConverter().ConvertAssemblyToTypeLib(
            assembly, tlbPath, flags, new ExporterSink());
        tlb.SaveAllChanges();
    }

    static readonly string[] PROXYSTUBS = {
        "{00020424-0000-0000-C000-000000000046}",  // Dual: PSOAInterface
        "{00020424-0000-0000-C000-000000000046}",  // IUnknown: PSOAInterface
        "{00020420-0000-0000-C000-000000000046}"   // IDispatch: PSDispatch
    };

    static T GetAttribute<T>(ICustomAttributeProvider obj) where T : Attribute {
        var attributes = obj.GetCustomAttributes(typeof(T), false);
        return attributes.Length == 0 ? null : (T)attributes[0];
    }

    /// <summary>
    /// Writes the InnoSetup code registering the COM visible types (included by SeleniumBasicSetup.iss).
    /// </summary>
    public static void WriteRegistration(string assemblyPath, string outPath) {
        Assembly assembly = Assembly.LoadFrom(assemblyPath);
        string filename = Path.GetFileName(assemblyPath);
        string basename = Path.GetFileNameWithoutExtension(assemblyPath);
        var typelibVersion = GetAttribute<TypeLibVersionAttribute>(assembly);

        var classes = new List<string>();
        var interfaces = new List<string>();
        var values = new List<string>();

        foreach (Type type in assembly.GetExportedTypes()) {
            var comvisible = GetAttribute<ComVisibleAttribute>(type);
            if (comvisible == null || !comvisible.Value)
                continue;
            string guid = "{" + GetAttribute<GuidAttribute>(type).Value + "}";
            if (type.IsClass) {
                if (type.GetConstructor(Type.EmptyTypes) != null) {
                    string progid = GetAttribute<ProgIdAttribute>(type).Value;
                    classes.Add(string.Format("RegClass(lib, '{0}', '{1}', '{2}');", guid, progid, type.FullName));
                }
            } else if (type.IsValueType) {
                values.Add(string.Format("RegRecord(lib, '{0}', '{1}');", guid, type.FullName));
            } else if (type.IsInterface) {
                var interfaceType = GetAttribute<InterfaceTypeAttribute>(type).Value;
                interfaces.Add(string.Format("RegInterface(lib, '{0}', '{1}', '{2}');", guid, type.Name, PROXYSTUBS[(int)interfaceType]));
            }
        }

        var sb = new StringBuilder();
        sb.Append("{\n");
        sb.Append(" Code generated at each build by build-setup.ps1\n");
        sb.Append(" This is a subset of SeleniumBasicSetup.iss\n");
        sb.Append("}\n\n");
        sb.AppendFormat("lib.Guid := '{{{0}}}';\n", GetAttribute<GuidAttribute>(assembly).Value);
        sb.AppendFormat("lib.FullName := '{0}';\n", assembly.FullName);
        sb.AppendFormat("lib.Description := '{0}';\n", GetAttribute<AssemblyDescriptionAttribute>(assembly).Description);
        sb.AppendFormat("lib.TypeVersion := '{0}.{1}';\n", typelibVersion.MajorVersion, typelibVersion.MinorVersion);
        sb.AppendFormat("lib.PathDll := ExpandConstant('{{app}}\\{0}');\n", filename);
        sb.AppendFormat("lib.PathTlb32 := ExpandConstant('{{app}}\\{0}32.tlb');\n", basename);
        sb.AppendFormat("lib.PathTlb64 := ExpandConstant('{{app}}\\{0}64.tlb');\n", basename);
        sb.AppendFormat("lib.Runtime := '{0}';\n", assembly.ImageRuntimeVersion);
        sb.Append("\n");
        sb.Append("RegTypeLib(lib);\n");
        foreach (var lines in new[] { classes, interfaces, values }) {
            sb.Append("\n");
            lines.Sort(StringComparer.Ordinal);
            foreach (string line in lines)
                sb.Append(line).Append("\n");
        }
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(false));
    }
}
'@

function Invoke-Native([string]$exe, [string[]]$arguments) {
    Write-Host "> $exe $($arguments -join ' ')" -ForegroundColor DarkGray
    & $exe @arguments
    if ($LASTEXITCODE -ne 0) { throw "Command failed with exit code ${LASTEXITCODE}: $exe" }
}

function Find-InnoSetup {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    throw "InnoSetup 6 not found. Install it with: winget install --id JRSoftware.InnoSetup -e"
}

# Child process: loads the assembly, so the parent can rebuild it afterwards
if ($Export) {
    Add-Type -TypeDefinition $SetupToolsSource
    Write-Host 'Export type library 32bits ...'
    [SetupTools]::ExportTypeLib($AssemblyPath, (Join-Path $ReleaseDir 'Selenium32.tlb'), $false)
    Write-Host 'Export type library 64bits ...'
    [SetupTools]::ExportTypeLib($AssemblyPath, (Join-Path $ReleaseDir 'Selenium64.tlb'), $true)
    Write-Host 'Generate registration file ...'
    [SetupTools]::WriteRegistration($AssemblyPath, (Join-Path $PSScriptRoot 'SeleniumBasicSetup.pas'))
    exit 0
}

$assemblyInfoPath = Join-Path $PSScriptRoot 'Selenium\Properties\AssemblyInfo.cs'
$assemblyInfo = [IO.File]::ReadAllText($assemblyInfoPath)
if ($Version) {
    if ($Version -notmatch '^\d+\.\d+\.\d+\.\d+$') { throw "Invalid version: $Version" }
    Write-Host "Update version number to $Version ..."
    $assemblyInfo = $assemblyInfo -replace '(Assembly(File)?Version\(")[.\d]+("\))', "`${1}$Version`${3}"
    [IO.File]::WriteAllText($assemblyInfoPath, $assemblyInfo, (New-Object Text.UTF8Encoding $true))
}
Write-Host "Version: $([regex]::Match($assemblyInfo, 'AssemblyFileVersion\("([.\d]+)"\)').Groups[1].Value)"

$iscc = Find-InnoSetup

Write-Host 'Delete previous builds ...'
foreach ($dir in 'Selenium\bin\Release', 'Selenium\obj\Release', 'VbsConsole\bin\Release', 'VbsConsole\obj\Release') {
    if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
}

Write-Host 'Build .Net library and console runner ...'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
Invoke-Native 'dotnet' @('build', 'SeleniumBasic.sln', '-c', 'Release', '-nologo', '-v', 'q')

Invoke-Native "$PSHOME\powershell.exe" @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Export', '1')

Write-Host 'Build setup package ...'
Invoke-Native $iscc @('/Q', "/O$PSScriptRoot", (Join-Path $PSScriptRoot 'SeleniumBasicSetup.iss'))

Write-Host "`nDone: $(Get-ChildItem $PSScriptRoot -Filter 'SeleniumBasic-*.exe' | Sort-Object LastWriteTime | Select-Object -Last 1 -ExpandProperty Name)"
