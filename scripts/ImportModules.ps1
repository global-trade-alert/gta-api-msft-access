<#
.SYNOPSIS
    Imports SGEPT API Integration VBA modules into Microsoft Access databases.

.DESCRIPTION
    This PowerShell script automates the process of importing the required VBA modules
    (JsonConverter.bas and modGtaSync.bas) into any Microsoft Access database for
    SGEPT API integration. It also attempts to configure the necessary VBA references.

.PARAMETER DatabasePath
    The full path to the Access database (.accdb or .mdb file) where modules will be imported.
    
.PARAMETER ModulesPath
    The path to the folder containing the VBA modules. Defaults to "../code" relative to script location.
    
.PARAMETER SetReferences
    Switch parameter. If specified, attempts to set the required VBA references automatically.
    Note: This may require elevated permissions and might not work in all environments.

.PARAMETER Force
    Switch parameter. If specified, overwrites existing modules with the same names without prompting.

.EXAMPLE
    .\ImportModules.ps1 -DatabasePath "C:\MyProject\GtaDemo.accdb"
    
    Imports VBA modules into the specified database using default module path.

.EXAMPLE
    .\ImportModules.ps1 -DatabasePath ".\samples\GtaDemo.accdb" -SetReferences -Force
    
    Imports modules, sets references, and overwrites existing modules without prompting.

.EXAMPLE
    .\ImportModules.ps1 -DatabasePath "C:\MyDB.accdb" -ModulesPath "C:\VBAModules"
    
    Imports modules from a custom path into the specified database.

.NOTES
    Requirements:
    - Microsoft Access 2016 or later
    - PowerShell 5.1 or later
    - Administrative rights may be required for setting VBA references
    - Access must be closed before running this script
    
    Author: SGEPT Integration Team
    Version: 1.0.0
    License: MIT
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Path to the Access database file")]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Database file not found: $_"
        }
        if ($_ -notmatch '\.(accdb|mdb)$') {
            throw "File must be an Access database (.accdb or .mdb): $_"
        }
        return $true
    })]
    [string]$DatabasePath,
    
    [Parameter(Mandatory=$false, HelpMessage="Path to folder containing VBA modules")]
    [string]$ModulesPath,
    
    [Parameter(Mandatory=$false, HelpMessage="Attempt to set VBA references automatically")]
    [switch]$SetReferences,
    
    [Parameter(Mandatory=$false, HelpMessage="Overwrite existing modules without prompting")]
    [switch]$Force
)

# Script configuration
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Module configuration
$RequiredModules = @(
    @{
        Name = "JsonConverter"
        FileName = "JsonConverter.bas"
        Description = "VBA-JSON Library by Tim Hall"
    },
    @{
        Name = "modGtaSync"
        FileName = "modGtaSync.bas"
        Description = "SGEPT API Integration Logic"
    }
)

# VBA References to set (if SetReferences is specified)
$RequiredReferences = @(
    @{
        Name = "Microsoft XML, v6.0"
        GUID = "{F5078F18-C551-11D3-89B9-0000F81FE221}"
        Version = "6.0"
    },
    @{
        Name = "Microsoft Scripting Runtime"
        GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"
        Version = "1.0"
    }
)

function Write-Header {
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "SGEPT API Integration - VBA Module Importer" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Progress {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] WARNING: $Message" -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ERROR: $Message" -ForegroundColor Red
}

function Get-ModulesPath {
    if ($ModulesPath) {
        return $ModulesPath
    }
    
    # Default to ../code relative to script location
    $scriptDir = Split-Path -Parent $MyInvocation.ScriptName
    $defaultPath = Join-Path (Split-Path -Parent $scriptDir) "code"
    
    if (Test-Path $defaultPath) {
        return $defaultPath
    }
    
    throw "Modules path not found. Please specify -ModulesPath parameter or ensure modules are in: $defaultPath"
}

function Test-ModuleFiles {
    param([string]$ModulesDirectory)
    
    Write-Progress "Checking for required VBA modules..."
    
    $missingFiles = @()
    foreach ($module in $RequiredModules) {
        $filePath = Join-Path $ModulesDirectory $module.FileName
        if (-not (Test-Path $filePath)) {
            $missingFiles += $module.FileName
        } else {
            Write-Verbose "Found: $($module.FileName)"
        }
    }
    
    if ($missingFiles.Count -gt 0) {
        throw "Missing VBA module files: $($missingFiles -join ', ') in directory: $ModulesDirectory"
    }
    
    Write-Progress "All required VBA modules found."
}

function New-AccessApplication {
    Write-Progress "Starting Microsoft Access..."
    
    try {
        $accessApp = New-Object -ComObject Access.Application
        $accessApp.Visible = $false
        Write-Verbose "Access application created successfully."
        return $accessApp
    }
    catch {
        throw "Failed to create Access application. Ensure Microsoft Access is installed and not already running. Error: $($_.Exception.Message)"
    }
}

function Open-Database {
    param([object]$AccessApp, [string]$DatabasePath)
    
    Write-Progress "Opening database: $(Split-Path -Leaf $DatabasePath)"
    
    try {
        $resolvedPath = Resolve-Path $DatabasePath
        $accessApp.OpenCurrentDatabase($resolvedPath.Path)
        Write-Verbose "Database opened successfully."
    }
    catch {
        throw "Failed to open database: $DatabasePath. Error: $($_.Exception.Message)"
    }
}

function Test-ModuleExists {
    param([object]$AccessApp, [string]$ModuleName)
    
    try {
        $modules = $AccessApp.CurrentProject.AllModules
        for ($i = 0; $i -lt $modules.Count; $i++) {
            if ($modules.Item($i).Name -eq $ModuleName) {
                return $true
            }
        }
        return $false
    }
    catch {
        Write-Warning "Could not check for existing module: $ModuleName"
        return $false
    }
}

function Import-VBAModule {
    param([object]$AccessApp, [string]$ModulePath, [string]$ModuleName, [string]$Description)
    
    try {
        # Check if module already exists
        if (Test-ModuleExists -AccessApp $AccessApp -ModuleName $ModuleName) {
            if (-not $Force) {
                $response = Read-Host "Module '$ModuleName' already exists. Overwrite? (y/N)"
                if ($response -notmatch '^[Yy]') {
                    Write-Progress "Skipped: $ModuleName"
                    return $true
                }
            }
            
            # Delete existing module
            Write-Verbose "Removing existing module: $ModuleName"
            $AccessApp.DoCmd.DeleteObject(5, $ModuleName)  # 5 = acModule
        }
        
        # Import the module
        Write-Progress "Importing: $ModuleName ($Description)"
        $AccessApp.DoCmd.TransferText(0, "", "", $ModulePath, $false, "")
        
        # Alternative method for VBA modules
        $AccessApp.LoadFromText(5, $ModuleName, $ModulePath)  # 5 = acModule
        
        Write-Verbose "Successfully imported: $ModuleName"
        return $true
    }
    catch {
        Write-Error "Failed to import module '$ModuleName': $($_.Exception.Message)"
        return $false
    }
}

function Set-VBAReferences {
    param([object]$AccessApp)
    
    if (-not $SetReferences) {
        Write-Progress "Skipping VBA references setup (use -SetReferences to enable)"
        return
    }
    
    Write-Progress "Attempting to set VBA references..."
    
    try {
        $vbProject = $AccessApp.VBE.ActiveVBProject
        $references = $vbProject.References
        
        foreach ($ref in $RequiredReferences) {
            try {
                # Check if reference already exists
                $found = $false
                foreach ($existingRef in $references) {
                    if ($existingRef.Name -eq $ref.Name) {
                        Write-Verbose "Reference already set: $($ref.Name)"
                        $found = $true
                        break
                    }
                }
                
                if (-not $found) {
                    Write-Verbose "Adding reference: $($ref.Name)"
                    $references.AddFromGuid($ref.GUID, [int]$ref.Version.Split('.')[0], [int]$ref.Version.Split('.')[1])
                    Write-Progress "Added reference: $($ref.Name)"
                }
            }
            catch {
                Write-Warning "Could not add reference '$($ref.Name)': $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Warning "VBA references setup failed: $($_.Exception.Message)"
        Write-Host "You may need to set VBA references manually:" -ForegroundColor Yellow
        foreach ($ref in $RequiredReferences) {
            Write-Host "  - $($ref.Name)" -ForegroundColor Yellow
        }
    }
}

function Save-Database {
    param([object]$AccessApp)
    
    Write-Progress "Saving database..."
    
    try {
        $AccessApp.DoCmd.Save()
        Write-Verbose "Database saved successfully."
    }
    catch {
        Write-Warning "Could not save database: $($_.Exception.Message)"
    }
}

function Close-AccessApplication {
    param([object]$AccessApp)
    
    if ($AccessApp) {
        Write-Progress "Closing Access application..."
        try {
            $AccessApp.CloseCurrentDatabase()
            $AccessApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AccessApp) | Out-Null
            Write-Verbose "Access application closed successfully."
        }
        catch {
            Write-Warning "Error closing Access application: $($_.Exception.Message)"
        }
    }
}

function Write-Summary {
    param([bool]$Success, [int]$ImportedCount, [int]$TotalCount)
    
    Write-Host ""
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "IMPORT SUMMARY" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    
    if ($Success) {
        Write-Host "✅ Import completed successfully!" -ForegroundColor Green
        Write-Host "📊 Modules imported: $ImportedCount of $TotalCount" -ForegroundColor Green
        
        if ($SetReferences) {
            Write-Host "🔗 VBA references setup attempted" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "Next Steps:" -ForegroundColor Yellow
        Write-Host "1. Open your Access database" -ForegroundColor White
        Write-Host "2. Press Alt+F11 to verify VBA modules are imported" -ForegroundColor White
        Write-Host "3. Check VBA references (Tools > References) if needed" -ForegroundColor White
        Write-Host "4. Configure your API key in tblSettings table" -ForegroundColor White
        Write-Host "5. Run SyncGTA to test the integration" -ForegroundColor White
    }
    else {
        Write-Host "❌ Import failed or incomplete" -ForegroundColor Red
        Write-Host "📊 Modules imported: $ImportedCount of $TotalCount" -ForegroundColor Red
        Write-Host ""
        Write-Host "Troubleshooting:" -ForegroundColor Yellow
        Write-Host "1. Ensure Access is not already running" -ForegroundColor White
        Write-Host "2. Check that you have write permissions to the database" -ForegroundColor White
        Write-Host "3. Try running PowerShell as Administrator" -ForegroundColor White
        Write-Host "4. Import modules manually if automation fails" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "For support, see: https://github.com/your-repo/gta-api-msft-access" -ForegroundColor Cyan
}

# Main execution
try {
    Write-Header
    
    # Validate inputs
    $modulesDirectory = Get-ModulesPath
    Test-ModuleFiles -ModulesDirectory $modulesDirectory
    
    # Initialize Access
    $accessApp = $null
    $importedCount = 0
    $totalCount = $RequiredModules.Count
    
    try {
        $accessApp = New-AccessApplication
        Open-Database -AccessApp $accessApp -DatabasePath $DatabasePath
        
        # Import each module
        foreach ($module in $RequiredModules) {
            $modulePath = Join-Path $modulesDirectory $module.FileName
            $success = Import-VBAModule -AccessApp $accessApp -ModulePath $modulePath -ModuleName $module.Name -Description $module.Description
            
            if ($success) {
                $importedCount++
            }
        }
        
        # Set VBA references if requested
        Set-VBAReferences -AccessApp $accessApp
        
        # Save database
        Save-Database -AccessApp $accessApp
        
        Write-Summary -Success ($importedCount -eq $totalCount) -ImportedCount $importedCount -TotalCount $totalCount
    }
    finally {
        Close-AccessApplication -AccessApp $accessApp
    }
}
catch {
    Write-Error $_.Exception.Message
    Write-Summary -Success $false -ImportedCount $importedCount -TotalCount $totalCount
    exit 1
} 