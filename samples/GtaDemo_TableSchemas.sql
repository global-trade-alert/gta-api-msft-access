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
    ID COUNTER PRIMARY KEY,
    setting_name TEXT(50) NOT NULL UNIQUE,
    setting_value TEXT(255),
    description MEMO,
    last_updated DATETIME
);

-- Insert default settings (one at a time)
INSERT INTO tblSettings (setting_name, setting_value, description) VALUES ('APIKey', '', 'SGEPT API key - required for data synchronization'); 
INSERT INTO tblSettings (setting_name, setting_value, description) VALUES ('LastSyncDate', '', 'Timestamp of last successful API sync'); 
INSERT INTO tblSettings (setting_name, setting_value, description) VALUES ('PageSize', '50', 'Number of records to fetch per API request (max 1000)'); 
INSERT INTO tblSettings (setting_name, setting_value, description) VALUES ('SyncEnabled', 'True', 'Enable/disable automatic synchronization');

-- ======================================================================
-- Table 3: tblSyncLog (User-Accessible Change Log)
-- ======================================================================
CREATE TABLE tblSyncLog (
    log_id COUNTER PRIMARY KEY,
    log_timestamp DATETIME,
    session_id TEXT(50),
    source_function TEXT(50),
    log_level TEXT(10),
    message MEMO,
    intervention_id LONG
);

-- ======================================================================
-- Table 2: tblGTAInterventions (Main Data Table)
-- ======================================================================
CREATE TABLE tblGTAInterventions (
    intervention_id LONG PRIMARY KEY,
    state_act_title TEXT(255),
    intervention_description MEMO,
    intervention_type TEXT(100),
    gta_evaluation TEXT(50),
    date_announced DATETIME,
    implementation_date DATETIME,
    removal_date DATETIME,
    last_updated DATETIME,
    implementing_jurisdiction_name TEXT(255),
    affected_jurisdictions MEMO,
    targeted_products_hs6 MEMO,
    targeted_sectors_cpc3 MEMO,
    source MEMO,
    sync_source TEXT(50)
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