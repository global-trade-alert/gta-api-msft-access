# PowerShell VBA Module Importer

This folder contains PowerShell scripts to automate the setup of SGEPT API integration in Microsoft Access databases.

## ImportModules.ps1

Automates the import of required VBA modules (`JsonConverter.bas` and `modGtaSync.bas`) into any Access database.

### Quick Start

```powershell
# Basic usage - import modules into your database
.\ImportModules.ps1 -DatabasePath "C:\MyProject\GtaDemo.accdb"

# Full automation - import modules, set references, overwrite existing
.\ImportModules.ps1 -DatabasePath ".\samples\GtaDemo.accdb" -SetReferences -Force
```

### Prerequisites

- **Microsoft Access 2016 or later** installed
- **PowerShell 5.1 or later** (check with `$PSVersionTable.PSVersion`)
- **Access database must be closed** before running script
- **Administrative rights** may be required for automatic reference setup

### Parameters

| Parameter | Required | Description | Default |
|-----------|----------|-------------|---------|
| `-DatabasePath` | Yes | Full path to Access database (.accdb/.mdb) | None |
| `-ModulesPath` | No | Path to folder containing VBA modules | `../code` |
| `-SetReferences` | No | Attempt to set VBA references automatically | False |
| `-Force` | No | Overwrite existing modules without prompting | False |

### Usage Examples

#### Basic Import
```powershell
.\ImportModules.ps1 -DatabasePath "C:\Projects\MyGtaIntegration.accdb"
```
- Imports VBA modules using default module path
- Prompts before overwriting existing modules
- Skips VBA reference setup (manual setup required)

#### Full Automation
```powershell
.\ImportModules.ps1 -DatabasePath ".\samples\GtaDemo.accdb" -SetReferences -Force
```
- Imports modules from default location
- Automatically sets VBA references (XML v6.0, Scripting Runtime)
- Overwrites existing modules without prompting
- Best for initial setup and testing

#### Custom Module Path
```powershell
.\ImportModules.ps1 -DatabasePath "C:\MyDB.accdb" -ModulesPath "C:\CustomVBA"
```
- Uses custom location for VBA module files
- Useful if you've moved or copied the modules elsewhere

#### Multiple Databases
```powershell
# Setup multiple databases in a loop
$databases = @("DB1.accdb", "DB2.accdb", "DB3.accdb")
foreach ($db in $databases) {
    .\ImportModules.ps1 -DatabasePath $db -SetReferences -Force
}
```

### What the Script Does

#### Phase 1: Validation
- ✅ Checks that database file exists and is valid Access format
- ✅ Verifies VBA module files are present
- ✅ Validates PowerShell execution environment

#### Phase 2: Access COM Automation
- 🔧 Starts Microsoft Access in background (invisible mode)
- 📂 Opens the specified database
- 🔍 Checks for existing modules with same names

#### Phase 3: Module Import
- 📄 Imports `JsonConverter.bas` (VBA-JSON library)
- 📄 Imports `modGtaSync.bas` (main integration logic)
- ⚠️ Prompts for overwrite confirmation (unless `-Force` specified)

#### Phase 4: VBA References (Optional)
- 🔗 Adds "Microsoft XML, v6.0" reference
- 🔗 Adds "Microsoft Scripting Runtime" reference
- ⚠️ Requires `-SetReferences` parameter

#### Phase 5: Cleanup
- 💾 Saves database changes
- 🚪 Closes Access application cleanly
- 📊 Displays import summary

### Troubleshooting

#### "Access application creation failed"
**Cause:** Access not installed, already running, or COM registration issues

**Solutions:**
1. Close any open Access instances
2. Restart PowerShell session
3. Run as Administrator
4. Repair Microsoft Office installation

#### "Database file not found"
**Cause:** Incorrect path or file doesn't exist

**Solutions:**
1. Use full absolute paths: `C:\Full\Path\To\Database.accdb`
2. Verify file exists: `Test-Path "YourDatabase.accdb"`
3. Check file permissions

#### "VBA module files missing"
**Cause:** Script can't find JsonConverter.bas or modGtaSync.bas

**Solutions:**
1. Ensure files are in `../code` relative to script location
2. Use `-ModulesPath` parameter to specify custom location
3. Download missing files from project repository

#### "Module import failed"
**Cause:** VBA import errors or file corruption

**Solutions:**
1. Try manual import in Access VBA editor
2. Check module file syntax and encoding
3. Ensure database isn't read-only or corrupted

#### "VBA references setup failed"
**Cause:** Insufficient permissions or missing COM libraries

**Solutions:**
1. Run PowerShell as Administrator
2. Set references manually in Access (Tools > References)
3. Install/repair missing Microsoft components

### Manual Fallback

If automation fails, import modules manually:

1. **Open Access database**
2. **Press Alt+F11** (VBA Editor)
3. **Right-click in Project Explorer**
4. **Choose "Import File"**
5. **Select JsonConverter.bas**, click Open
6. **Repeat for modGtaSync.bas**
7. **Set References** (Tools > References):
   - ☑️ Microsoft XML, v6.0
   - ☑️ Microsoft Scripting Runtime

### Security Considerations

#### Execution Policy
PowerShell may require execution policy changes:
```powershell
# Check current policy
Get-ExecutionPolicy

# Enable script execution (if needed)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### COM Security
- Script uses Microsoft Access COM automation
- Some corporate environments may block COM operations
- Administrative rights may be required for VBA reference setup

### Performance Notes

- **Typical execution time**: 15-30 seconds per database
- **Resource usage**: Minimal (Access runs in background)
- **Parallel processing**: Not recommended (Access COM limitations)

### Advanced Usage

#### Batch Processing
```powershell
# Import into all databases in a directory
Get-ChildItem -Path "C:\Databases" -Filter "*.accdb" | ForEach-Object {
    Write-Host "Processing: $($_.Name)"
    .\ImportModules.ps1 -DatabasePath $_.FullName -SetReferences -Force
}
```

#### Custom Module Sets
```powershell
# Modify $RequiredModules array in script for different module sets
# Useful for custom implementations or additional modules
```

#### Integration with CI/CD
```powershell
# Example for automated deployment
try {
    .\ImportModules.ps1 -DatabasePath $env:TARGET_DATABASE -SetReferences -Force
    Write-Host "Deployment successful"
    exit 0
}
catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    exit 1
}
```

---

## Additional Scripts

This folder may contain additional PowerShell utilities:

- **DatabaseSetup.ps1** - Complete database creation and configuration
- **ScheduleSync.ps1** - Windows Task Scheduler automation
- **UpdateModules.ps1** - Update existing installations with new module versions

---

## Support

For PowerShell script issues:
1. Check the detailed error messages in script output
2. Verify all prerequisites are met
3. Try manual VBA import as fallback
4. Report issues on the project GitHub repository

For SGEPT API integration help:
- See main project documentation: `../docs/README.md`
- Architecture overview: `../docs/Architecture.md`
- Setup guide: `../samples/GtaDemo_SetupGuide.md` 