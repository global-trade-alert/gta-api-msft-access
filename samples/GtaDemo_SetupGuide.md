# GtaDemo.accdb Setup Guide

This guide walks you through creating the Access database for SGEPT API integration step by step.

## Prerequisites

- Microsoft Access 2016 or later
- SGEPT API key (demo or full access)
- Administrative rights to install VBA references

## Step 1: Create New Access Database

1. Open Microsoft Access
2. Click **Blank desktop database**
3. Name the database: `GtaDemo.accdb`
4. Choose location: `samples/` folder in your project
5. Click **Create**

## Step 2: Enable Required References

1. Press `Alt + F11` to open VBA Editor
2. Go to **Tools** > **References**
3. Check the following boxes:
   - ☑️ **Microsoft XML, v6.0** (for HTTP requests)
   - ☑️ **Microsoft Scripting Runtime** (for Dictionary objects)
   - ☑️ **Microsoft VBA Object Library** (usually pre-selected)
4. Click **OK**

## Step 3: Create Database Tables

### Option A: Using SQL View
1. In Access, go to **Create** > **Query Design**
2. Close the "Show Table" dialog
3. Go to **Design** > **SQL View**
4. Copy and paste the contents of `GtaDemo_TableSchemas.sql`
5. Run each CREATE TABLE statement separately

### Option B: Using Table Design View
Create the tables manually using the field specifications below:

#### Table 1: tblSettings
| Field Name | Data Type | Field Size | Required | Default |
|------------|-----------|------------|----------|---------|
| ID | AutoNumber | Long Integer | Yes | - |
| setting_name | Short Text | 50 | Yes | - |
| setting_value | Short Text | 255 | No | - |
| description | Long Text | - | No | - |
| last_updated | Date/Time | - | No | Now() |

**Primary Key:** ID  
**Unique Index:** setting_name

#### Table 2: tblGTAInterventions
| Field Name | Data Type | Field Size | Required | Group |
|------------|-----------|------------|----------|-------|
| intervention_id | Number (Long) | - | Yes | Core |
| state_act_title | Short Text | 255 | No | Core |
| intervention_description | Long Text | - | No | Core* |
| intervention_type | Short Text | 100 | No | Core |
| gta_evaluation | Short Text | 50 | No | Core |
| date_announced | Date/Time | - | No | Dates |
| implementation_date | Date/Time | - | No | Dates |
| removal_date | Date/Time | - | No | Dates |
| last_updated | Date/Time | - | No | Dates |
| implementing_jurisdiction_name | Short Text | 255 | No | Geographic |
| affected_jurisdictions | Long Text | - | No | Geographic |
| targeted_products_hs6 | Long Text | - | No | Economic |
| targeted_sectors_cpc3 | Long Text | - | No | Economic |
| source | Long Text | - | No | Admin* |
| sync_source | Short Text | 50 | No | Admin |

**Primary Key:** intervention_id  
**Indexes:** date_announced, intervention_type, gta_evaluation, last_updated

#### Table 3: tblSyncLog (Change Log)
| Field Name | Data Type | Field Size | Required | Purpose |
|------------|-----------|------------|----------|---------|
| log_id | AutoNumber | Long Integer | Yes | Primary Key |
| log_timestamp | Date/Time | - | Yes | When event occurred |
| session_id | Short Text | 50 | No | Groups related sync operations |
| source_function | Short Text | 50 | No | Which VBA function logged this |
| log_level | Short Text | 10 | No | INFO, WARNING, ERROR, SUCCESS |
| message | Long Text | 500 | No | Detailed log message |
| intervention_id | Number (Long) | - | No | Optional link to intervention |

**Primary Key:** log_id  
**Indexes:** log_timestamp, session_id, log_level

*Fields marked with asterisk (*) require full API access

## Step 4: Insert Default Settings

Run this SQL in a new query:

```sql
INSERT INTO tblSettings (setting_name, setting_value, description) VALUES 
    ('APIKey', '', 'SGEPT API key - required for data synchronization'),
    ('LastSyncDate', '', 'Timestamp of last successful API sync'),
    ('PageSize', '50', 'Number of records to fetch per API request (max 1000)'),
    ('SyncEnabled', 'True', 'Enable/disable automatic synchronization');
```

## Step 5: Import VBA Modules

1. In VBA Editor (`Alt + F11`)
2. Right-click in Project Explorer
3. Choose **Import File**
4. Import these modules:
   - `code/JsonConverter.bas`
   - `code/modGtaSync.bas`

## Step 6: Configure API Key

1. Open the `tblSettings` table
2. Find the row where `setting_name = 'APIKey'`
3. Enter your SGEPT API key in the `setting_value` field
4. Save the table

## Step 7: Test the Integration

1. Press `Ctrl + G` to open Immediate Window in VBA
2. Type: `SyncGTA` and press Enter
3. Check the `tblGTAInterventions` table for imported data
4. **View the sync log:** Open `tblSyncLog` to see what happened during sync

## Step 8: Using the Change Log

The `tblSyncLog` table provides complete visibility into sync operations:

### View Recent Activity
- Open `tblSyncLog` table directly
- Use the `qrySyncActivity` query for formatted view
- Check `qryLatestSyncSummary` for high-level summary

### Understanding Log Levels
- **SUCCESS:** Records inserted/updated successfully  
- **INFO:** General sync progress information
- **WARNING:** Non-critical issues (missing optional fields)
- **ERROR:** Failed operations that need attention

### Session Tracking
Each sync run gets a unique `session_id` like `SYNC_20241201_143052_123` that groups all related log entries together.

## Troubleshooting

### Common Issues

**Error: "Object library not registered"**
- Solution: Re-enable the references in Step 2

**Error: "API key not found"**
- Solution: Verify the API key is correctly entered in tblSettings

**Error: "No data returned"**
- Solution: Check API key validity and internet connection

**Error: "Permission denied"**
- Solution: Enable macros and trusted locations in Access security settings

### API Access Levels

**Demo API Key:**
- Basic intervention data available
- Missing: intervention_description, source fields

**Full API Key:**
- Complete data including descriptions and sources
- Request trial access from SGEPT support

## Next Steps

1. **Automation:** Set up Windows Task Scheduler to run sync automatically
2. **Reporting:** Create Access queries and reports for data analysis
3. **Integration:** Link to Excel or Power BI for advanced analytics

## File Structure

After setup, your project should look like:
```
samples/
├── GtaDemo.accdb           # Your Access database
├── GtaDemo_TableSchemas.sql # Table creation scripts
└── GtaDemo_SetupGuide.md   # This guide

code/
├── JsonConverter.bas       # VBA JSON library
└── modGtaSync.bas         # Main sync module
``` 