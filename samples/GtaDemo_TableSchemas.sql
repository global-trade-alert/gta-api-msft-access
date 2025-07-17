-- ======================================================================
-- SGEPT API Access Integration - Database Schema
-- File: GtaDemo_TableSchemas.sql
-- 
-- Instructions: Run these CREATE TABLE statements in a new Access database
-- to create the required table structures for GTA intervention sync.
-- ======================================================================

-- ======================================================================
-- Table 1: tblSettings (API Configuration)
-- ======================================================================
CREATE TABLE tblSettings (
    ID AUTOINCREMENT PRIMARY KEY,
    setting_name TEXT(50) NOT NULL,
    setting_value TEXT(255),
    description TEXT(500),
    last_updated DATETIME DEFAULT Now(),
    CONSTRAINT UK_tblSettings_name UNIQUE (setting_name)
);

-- Insert default settings
INSERT INTO tblSettings (setting_name, setting_value, description) VALUES 
    ('APIKey', '', 'SGEPT API key - required for data synchronization'),
    ('LastSyncDate', '', 'Timestamp of last successful API sync'),
    ('PageSize', '50', 'Number of records to fetch per API request (max 1000)'),
    ('SyncEnabled', 'True', 'Enable/disable automatic synchronization');

-- ======================================================================
-- Table 3: tblSyncLog (User-Accessible Change Log)
-- ======================================================================
CREATE TABLE tblSyncLog (
    log_id AUTOINCREMENT PRIMARY KEY,
    log_timestamp DATETIME DEFAULT Now(),
    session_id TEXT(50),                           -- Groups related log entries
    source_function TEXT(50),                      -- Function that generated the log
    log_level TEXT(10),                           -- INFO, WARNING, ERROR, SUCCESS
    message TEXT(500),                            -- Log message content
    intervention_id LONG,                         -- Optional: Link to specific intervention
    
    INDEX IX_tblSyncLog_Timestamp (log_timestamp),
    INDEX IX_tblSyncLog_Session (session_id),
    INDEX IX_tblSyncLog_Level (log_level)
);

-- ======================================================================
-- Table 2: tblGTAInterventions (Main Data Table)
-- ======================================================================
CREATE TABLE tblGTAInterventions (
    -- === GROUP 1: CORE INTERVENTION INFORMATION ===
    intervention_id LONG PRIMARY KEY,
    state_act_title TEXT(255),
    intervention_description MEMO,                    -- Requires full API access
    intervention_type TEXT(100),
    gta_evaluation TEXT(50),                         -- Red/Amber/Green triangle
    
    -- === GROUP 2: KEY DATES ===
    date_announced DATETIME,
    implementation_date DATETIME,
    removal_date DATETIME,
    last_updated DATETIME DEFAULT Now(),
    
    -- === GROUP 3: GEOGRAPHIC SCOPE ===
    implementing_jurisdiction_name TEXT(255),
    affected_jurisdictions TEXT(500),               -- Comma-separated list
    
    -- === GROUP 4: ECONOMIC TARGETING ===
    targeted_products_hs6 MEMO,                     -- HS 6-digit codes (comma-separated)
    targeted_sectors_cpc3 TEXT(500),                -- CPC 3-digit codes (comma-separated)
    
    -- === GROUP 5: ADMINISTRATIVE ===
    source TEXT(500),                               -- Requires full API access
    sync_source TEXT(50) DEFAULT 'SGEPT_API'
);

-- ======================================================================
-- Indexes for Performance
-- ======================================================================
CREATE INDEX IX_tblGTAInterventions_DateAnnounced ON tblGTAInterventions (date_announced);
CREATE INDEX IX_tblGTAInterventions_InterventionType ON tblGTAInterventions (intervention_type);
CREATE INDEX IX_tblGTAInterventions_GTAEvaluation ON tblGTAInterventions (gta_evaluation);
CREATE INDEX IX_tblGTAInterventions_LastUpdated ON tblGTAInterventions (last_updated);

-- ======================================================================
-- Field Descriptions / Comments
-- ======================================================================
/*
FIELD AVAILABILITY BY API ACCESS LEVEL:

DEMO API KEY (Available immediately):
- intervention_id, state_act_title, intervention_type, gta_evaluation
- date_announced, implementation_date, removal_date
- implementing_jurisdiction_name, affected_jurisdictions  
- targeted_products_hs6, targeted_sectors_cpc3

FULL API KEY (Available upon request for trial purposes):
- intervention_description (detailed description)
- source (official source URL/reference)

DATA TYPES EXPLAINED:
- LONG: 32-bit integer for intervention_id (primary key)
- TEXT(n): Variable text field with maximum length n
- MEMO: Large text field for descriptions and code lists
- DATETIME: Date and time values
- AUTOINCREMENT: Auto-incrementing integer for settings table

MAST CHAPTER D FOCUS:
This schema is optimized for MAST Chapter D (Contingent trade-protective measures)
including anti-dumping, safeguards, and countervailing measures.
*/ 