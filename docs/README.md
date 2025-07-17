# SGEPT API Integration for Microsoft Access

> **Pull the 50 most recent global trade interventions directly into your Microsoft Access database with a single click.**

## What This Project Does

This integration automatically downloads trade intervention data from the Global Trade Alert (GTA) database via the SGEPT API into your local Microsoft Access database. Think of it as a data pipeline that keeps your local trade analysis up-to-date with the latest global trade restrictions, anti-dumping measures, and other trade policy changes.

📋 **See the complete system architecture**: [Architecture Overview](Architecture.md) | [Visual Diagram](Architecture_Diagram_Source.md)

### Business Value

- ✅ **Stay Current**: Get the latest trade interventions without manual data entry
- ✅ **Save Time**: Eliminate copy-paste from websites or manual Excel imports  
- ✅ **Ensure Accuracy**: Direct API connection prevents human transcription errors
- ✅ **Enable Analysis**: Query and analyze trade data using familiar Access tools
- ✅ **Track Changes**: See exactly what changed since your last sync with built-in change logs

### What You Get

When setup is complete, you'll have:

1. **Automated Data Sync**: Press one button to update your trade intervention database
2. **Rich Data**: 15+ fields per intervention including dates, jurisdictions, product codes, and evaluations  
3. **Change Tracking**: Complete audit trail showing what records were added, updated, or unchanged
4. **Ready-to-Use Queries**: Pre-built reports for common analysis needs
5. **Professional Setup**: Production-ready system with error handling and logging

---

## Prerequisites

### What You Need

- **Microsoft Access 2016 or later** (part of Microsoft Office Professional)
- **Internet connection** for API access
- **SGEPT API key** ([request here](https://www.globaltradealert.org/api))
- **Windows computer** with administrative rights to install VBA references
- **Basic familiarity with Access** (opening databases, running queries)

### API Access Levels

**Demo API Key** (Free)
- Basic intervention data (title, type, dates, jurisdictions, product codes)
- Perfect for testing and evaluation
- Request at: [SGEPT API Registration](https://www.globaltradealert.org/api)

**Full API Key** (Available upon request)
- All demo features PLUS detailed descriptions and source URLs
- Required for production analysis
- Contact SGEPT support to upgrade from demo access

---

## Quick Start (For Experienced Users)

1. **Download**: Clone this repository or download the ZIP file
2. **Setup Database**: Run `samples/GtaDemo_TableSchemas.sql` in a new Access database
3. **Import VBA**: Import `code/JsonConverter.bas` and `code/modGtaSync.bas` modules
4. **Configure References**: Enable Microsoft XML v6.0 and Scripting Runtime in VBA
5. **Add API Key**: Enter your SGEPT API key in the `tblSettings` table
6. **Sync**: Run `SyncGTA` function in VBA Immediate Window
7. **Analyze**: Use pre-built queries or create your own

📖 **Need more detail?** Follow the complete setup guide below.

---

## Complete Setup Guide

### Step 1: Download the Project Files

**Option A: Download ZIP File**
1. Click the green "Code" button at the top of this page
2. Select "Download ZIP"
3. Extract the files to a folder like `C:\GTA-Integration`

**Option B: Clone Repository (if you use Git)**
```bash
git clone https://github.com/your-repo/gta-api-msft-access.git
cd gta-api-msft-access
```

### Step 2: Create Your Access Database

1. **Open Microsoft Access**
2. **Create New Database**:
   - Choose "Blank desktop database"
   - Name it: `GtaDemo.accdb`
   - Save in the `samples/` folder of your project
3. **Click "Create"**

### Step 3: Set Up Database Tables

**Method A: Using SQL Scripts (Recommended)**
1. In Access, go to **Create** → **Query Design**
2. Close the "Show Table" dialog that appears
3. Click **SQL View** in the ribbon
4. Open the file `samples/GtaDemo_TableSchemas.sql` in Notepad
5. Copy and paste **each CREATE TABLE statement separately** into Access
6. Click **Run** (!) for each statement
7. You should now see three tables: `tblSettings`, `tblGTAInterventions`, `tblSyncLog`

**Method B: Manual Table Creation**
If you prefer not to use SQL, follow the detailed table specifications in `samples/GtaDemo_SetupGuide.md`.

### Step 4: Configure VBA References

1. **Open VBA Editor**: Press `Alt + F11`
2. **Access References**: Go to **Tools** → **References**
3. **Check Required Libraries**:
   - ☑️ Microsoft XML, v6.0
   - ☑️ Microsoft Scripting Runtime  
   - ☑️ Microsoft VBA Object Library (usually already checked)
4. **Click OK**

### Step 5: Import VBA Code Modules

**Option A: Automated Import (Recommended)**
```powershell
# Navigate to the scripts folder and run the PowerShell importer
cd scripts
.\ImportModules.ps1 -DatabasePath "..\samples\GtaDemo.accdb" -SetReferences -Force
```

**Option B: Manual Import**
1. **In VBA Editor**: Right-click in the Project Explorer
2. **Import First Module**: Choose **Import File** → select `code/JsonConverter.bas`
3. **Import Second Module**: Choose **Import File** → select `code/modGtaSync.bas`
4. **Verify**: You should now see both modules in your Project Explorer

💡 **PowerShell Automation**: The `scripts/ImportModules.ps1` automates VBA import and reference setup. See `scripts/README.md` for details.

### Step 6: Get Your API Key

1. **Visit**: [SGEPT API Registration](https://www.globaltradealert.org/api)
2. **Register**: Provide your organization details
3. **Request Demo Access**: Start with demo key for testing
4. **Save Your Key**: You'll receive it via email

### Step 7: Configure the API Key

1. **Open** the `tblSettings` table in Access
2. **Find the APIKey row**: Look for `setting_name = 'APIKey'`
3. **Enter Your Key**: Paste your SGEPT API key in the `setting_value` column
4. **Save** the table

### Step 8: Test Your First Sync

1. **Open VBA Editor**: Press `Alt + F11`
2. **Open Immediate Window**: Press `Ctrl + G`
3. **Run Sync**: Type `SyncGTA` and press Enter
4. **Watch Progress**: You'll see status messages in the Immediate Window
5. **Check Results**: Open the `tblGTAInterventions` table to see your data

**Success Looks Like:**
```
2024-12-01 14:30:52 [SyncGTA] Starting GTA data synchronization (PageSize: 50)
2024-12-01 14:30:53 [SyncGTA] Making API request to https://api.globaltradealert.org/api/v1/data/
2024-12-01 14:31:15 [SyncGTA] Sync completed successfully. Records: 50, Time: 23.1s
```

---

## Daily Operations

### Running a Sync

**Manual Sync (Recommended for Testing)**
1. Open your Access database
2. Press `Alt + F11` to open VBA
3. Press `Ctrl + G` for Immediate Window  
4. Type `SyncGTA` and press Enter

**Custom Page Size**
```vba
SyncGTA 25    ' Sync 25 records (faster)
SyncGTA 100   ' Sync 100 records (more data)
```

### Viewing Your Data

**Main Data Table**: `tblGTAInterventions`
- Open this table to see all intervention records
- Sort by `date_announced` to see newest first
- Filter by `intervention_type` to focus on specific measures

**Change Log**: `tblSyncLog`  
- See exactly what happened during each sync
- Filter by `log_level = "ERROR"` to find problems
- Use `session_id` to group related log entries

### Using Pre-Built Reports

The system includes ready-to-use queries for common analysis:

1. **Recent Activity**: `qryRecentInterventions` - Last 30 days
2. **Summary by Type**: `qryInterventionSummary` - Count by intervention type
3. **By Country**: `qryByJurisdiction` - Activity by implementing country
4. **Trade Defense**: `qryTradeDefenseMeasures` - Anti-dumping, safeguards only
5. **Sync History**: `qrySyncActivity` - Recent sync operations

**To run a query:**
1. Click **Queries** in the Access navigation pane
2. Double-click any query name to run it
3. Results open in a new window for viewing or exporting

---

## Understanding Your Data

### Core Fields (Available with Demo API Key)

| Field | What It Means | Example |
|-------|---------------|---------|
| `intervention_id` | Unique identifier for each trade measure | 54321 |
| `state_act_title` | Official name of the measure | "Import tariffs on steel products" |
| `intervention_type` | Category of trade measure | "Import tariff", "Anti-dumping" |
| `gta_evaluation` | Impact assessment | "Red" (harmful), "Green" (liberalizing) |
| `date_announced` | When measure was announced | 2024-03-15 |
| `implementation_date` | When measure takes effect | 2024-04-01 |
| `implementing_jurisdiction_name` | Country implementing the measure | "United States" |
| `affected_jurisdictions` | Countries affected by the measure | "China, Germany, Japan" |
| `targeted_products_hs6` | Products affected (6-digit HS codes) | "720711, 720712, 720719" |

### Enhanced Fields (Full API Key Required)

| Field | What It Means | Why It Matters |
|-------|---------------|----------------|
| `intervention_description` | Detailed description of the measure | Understand the full context and scope |
| `source` | Official government source URL | Verify information and get original documents |

### Understanding Product and Sector Codes

**HS6 Codes**: 6-digit Harmonized System codes identify specific products
- Example: `720711` = "Semi-finished products of iron/steel, rectangular cross-section"
- [Lookup tool](https://www.trade.gov/harmonized-system-hs-codes)

**CPC3 Codes**: 3-digit Common Product Classification codes identify broader sectors  
- Example: `351` = "Basic chemicals, fertilizers, plastics"
- [Reference guide](https://unstats.un.org/unsd/classifications/Econ/CPC)

---

## Troubleshooting

### Common Setup Issues

**"Object library not registered" Error**
- **Cause**: VBA references not properly enabled
- **Solution**: Repeat Step 4 (Configure VBA References)
- **Verify**: Ensure checkmarks are present and no "MISSING" labels appear

**"API key not found" Error**  
- **Cause**: API key not entered in settings table
- **Solution**: Double-check Step 7 (Configure API Key)
- **Verify**: Open `tblSettings` and confirm API key value is populated

**"No data returned" Error**
- **Cause**: Invalid API key or internet connectivity
- **Solution**: 
  1. Verify internet connection
  2. Test API key validity (try smaller page size: `SyncGTA 10`)
  3. Contact SGEPT support if key appears valid

**"Permission denied" Error**
- **Cause**: Access security settings blocking VBA/internet access
- **Solution**:
  1. Enable macros when opening the database
  2. Add database location to Access "Trusted Locations"
  3. Check corporate firewall settings

### Common Operation Issues

**Sync Takes Too Long**
- **Cause**: Large page size or slow internet
- **Solution**: Use smaller page size (`SyncGTA 25`)
- **Normal**: 50 records typically takes 15-30 seconds

**Some Records Not Updating**
- **Cause**: Working as designed - system only updates records that actually changed
- **Verification**: Check `tblSyncLog` for "unchanged, skipping" messages
- **This is good**: Saves time and prevents unnecessary database writes

**Missing Description or Source Fields**
- **Cause**: Demo API key doesn't include these fields
- **Solution**: Request full API access from SGEPT support
- **Workaround**: Focus analysis on available fields for now

### Getting Help

**Check the Sync Log First**
1. Open `tblSyncLog` table
2. Look for recent ERROR or WARNING entries
3. The `message` field usually explains what went wrong

**Run Diagnostic Queries**
- `qryErrorLog` - Shows only problems
- `qryLatestSyncSummary` - Overview of most recent sync
- `qryDataQuality` - Checks data completeness

**Contact Information**
- **Technical Issues**: Check this project's Issues page on GitHub
- **API Access**: Contact SGEPT support at [support email]
- **VBA/Access Help**: Microsoft Access community forums

---

## Advanced Topics

### Scheduling Automatic Syncs

**Option 1: Windows Task Scheduler**
1. Create a batch file that opens Access and runs the sync
2. Schedule it to run daily/weekly using Windows Task Scheduler
3. Requires some command-line scripting knowledge

**Option 2: Access Automation**
1. Create an Access macro that runs `SyncGTA`
2. Set up AutoExec macro to run on database open
3. Use Windows Scheduled Tasks to open the database

**Option 3: PowerShell Automation**
Use the included PowerShell scripts for comprehensive automation:
```powershell
# Automated VBA module import and reference setup
.\scripts\ImportModules.ps1 -DatabasePath "YourDatabase.accdb" -SetReferences -Force

# Batch processing multiple databases
Get-ChildItem "*.accdb" | ForEach-Object {
    .\scripts\ImportModules.ps1 -DatabasePath $_.FullName -SetReferences -Force
}
```
See `scripts/README.md` for complete automation options.

### Integrating with Other Systems

**Excel Integration**
1. Link Access tables to Excel using Data → Get Data → From Database
2. Create pivot tables and charts from the intervention data
3. Refresh connection to get latest data

**Power BI Integration**  
1. Connect Power BI to your Access database
2. Build dashboards and visualizations
3. Share reports with your organization

**Custom VBA Extensions**
- Modify `modGtaSync.bas` to add new fields or processing logic
- Create custom queries for your specific analysis needs
- Add email notifications when sync completes

### Data Management

**Managing Database Size**
- The system keeps all historical data by default
- Consider archiving old records if database becomes large
- Sample archiving query provided in `samples/GtaDemo_SampleQueries.sql`

**Backup Strategy**
- Regular backups of the .accdb file
- Export critical data to Excel/CSV for additional backup
- Version control for VBA code changes

**Performance Optimization**
- Keep the database compact and repaired (Database Tools → Compact)
- Monitor sync log for performance trends
- Consider reducing page size if syncs become slow

---

## File Structure

After setup, your project should look like this:

```
gta-api-msft-access/
├── docs/
│   ├── README.md                    # This file
│   ├── Architecture.md              # Detailed system architecture
│   └── Architecture_Diagram_Source.md # Visual diagram (Mermaid source)
├── code/
│   ├── JsonConverter.bas            # VBA JSON library
│   └── modGtaSync.bas              # Main sync logic
├── samples/
│   ├── GtaDemo.accdb               # Your database (after setup)
│   ├── GtaDemo_TableSchemas.sql    # Database creation scripts
│   ├── GtaDemo_SetupGuide.md       # Detailed setup instructions  
│   └── GtaDemo_SampleQueries.sql   # Pre-built analysis queries
├── scripts/
│   ├── ImportModules.ps1           # PowerShell VBA module importer
│   └── README.md                   # PowerShell scripts documentation
├── .github/
│   └── workflows/                  # CI/CD automation
├── LICENSE                         # MIT license for code
└── CHANGELOG.md                    # Version history
```

---

## API Reference

### Sync Function Options

```vba
' Basic sync (50 records, default)
SyncGTA

' Custom page size (1-1000 records)
SyncGTA 25      ' Faster, fewer records
SyncGTA 100     ' More data per sync

' Note: API has rate limits, don't sync too frequently
```

### Settings Table Configuration

| Setting | Default | Purpose | Valid Values |
|---------|---------|---------|--------------|
| APIKey | (empty) | Your SGEPT API key | String from SGEPT |
| PageSize | 50 | Records per sync | 1-1000 |
| SyncEnabled | True | Enable/disable sync | True/False |
| LastSyncDate | (auto) | Timestamp tracking | Auto-managed |

### Log Levels

| Level | Meaning | Action Needed |
|-------|---------|---------------|
| SUCCESS | Operation completed successfully | None |
| INFO | General progress information | None |
| WARNING | Non-critical issue occurred | Review but no action required |
| ERROR | Operation failed | Investigation required |

---

## License and Credits

### Code License
This project's code is licensed under the **MIT License** - see `LICENSE` file for details.

### Documentation License  
Documentation is licensed under **Creative Commons BY 4.0** - you're free to share and adapt with attribution.

### Third-Party Components

- **VBA-JSON Library**: By Tim Hall, MIT License
- **SGEPT API**: By Global Trade Alert initiative
- **Harmonized System Codes**: World Customs Organization
- **CPC Codes**: United Nations Statistics Division

### Acknowledgments

- Global Trade Alert team for providing the SGEPT API
- Microsoft Access community for VBA guidance
- Contributors to this project (see GitHub contributors page)

---

## Getting Started Checklist

Print this checklist and check off each step:

**Setup Phase**
- [ ] Downloaded project files
- [ ] Created new Access database (`GtaDemo.accdb`)
- [ ] Created database tables using SQL scripts
- [ ] Enabled VBA references (XML v6.0, Scripting Runtime)
- [ ] Imported VBA modules (`JsonConverter.bas`, `modGtaSync.bas`)
- [ ] Requested SGEPT API key
- [ ] Entered API key in `tblSettings` table

**Testing Phase**
- [ ] Ran first sync (`SyncGTA` in VBA Immediate Window)
- [ ] Verified data in `tblGTAInterventions` table
- [ ] Checked sync log in `tblSyncLog` table
- [ ] Tested pre-built queries (e.g., `qryRecentInterventions`)

**Production Phase**
- [ ] Documented sync schedule for your organization
- [ ] Trained team members on data access and queries
- [ ] Set up backup procedures
- [ ] Considered automation options (if needed)

**You're Ready!** 🎉

Your SGEPT API integration is complete. You can now pull the latest global trade intervention data into Access with a single command, analyze it using familiar Access tools, and stay current with global trade policy developments.

---

*For the latest version of this documentation and project updates, visit: [GitHub Repository URL]* 