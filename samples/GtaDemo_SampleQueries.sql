-- ======================================================================
-- SGEPT API Access Integration - Sample Queries
-- File: GtaDemo_SampleQueries.sql
-- 
-- Instructions: Create these queries in Access for data analysis
-- Copy each query into Access Query Design > SQL View
-- ======================================================================

-- ======================================================================
-- Query 1: Recent MAST Chapter D Interventions (Last 30 Days)
-- ======================================================================
-- Name: qryRecentInterventions
SELECT 
    intervention_id,
    state_act_title,
    intervention_type,
    gta_evaluation,
    date_announced,
    implementing_jurisdiction_name,
    affected_jurisdictions
FROM tblGTAInterventions
WHERE date_announced >= DateAdd("d", -30, Date())
ORDER BY date_announced DESC;

-- ======================================================================
-- Query 2: Interventions by Type and Evaluation
-- ======================================================================
-- Name: qryInterventionSummary
SELECT 
    intervention_type,
    gta_evaluation,
    Count(*) AS intervention_count,
    Min(date_announced) AS earliest_date,
    Max(date_announced) AS latest_date
FROM tblGTAInterventions
WHERE intervention_type Is Not Null
GROUP BY intervention_type, gta_evaluation
ORDER BY intervention_type, gta_evaluation;

-- ======================================================================
-- Query 3: Interventions by Implementing Jurisdiction
-- ======================================================================
-- Name: qryByJurisdiction
SELECT 
    implementing_jurisdiction_name,
    Count(*) AS total_interventions,
    Sum(IIf(gta_evaluation = "Red", 1, 0)) AS harmful_measures,
    Sum(IIf(gta_evaluation = "Green", 1, 0)) AS liberalizing_measures,
    Sum(IIf(gta_evaluation = "Amber", 1, 0)) AS amber_measures
FROM tblGTAInterventions
WHERE implementing_jurisdiction_name Is Not Null
GROUP BY implementing_jurisdiction_name
HAVING Count(*) > 0
ORDER BY total_interventions DESC;

-- ======================================================================
-- Query 4: Anti-Dumping and Safeguard Measures Only
-- ======================================================================
-- Name: qryTradeDefenseMeasures
SELECT 
    intervention_id,
    state_act_title,
    intervention_type,
    date_announced,
    implementation_date,
    removal_date,
    implementing_jurisdiction_name,
    affected_jurisdictions,
    targeted_products_hs6
FROM tblGTAInterventions
WHERE intervention_type IN ("Anti-dumping", "Safeguard", "Anti-subsidy", "Anti-circumvention")
ORDER BY date_announced DESC;

-- ======================================================================
-- Query 5: Interventions with Product Codes (HS6)
-- ======================================================================
-- Name: qryWithProductCodes
SELECT 
    intervention_id,
    state_act_title,
    intervention_type,
    date_announced,
    implementing_jurisdiction_name,
    targeted_products_hs6,
    targeted_sectors_cpc3
FROM tblGTAInterventions
WHERE targeted_products_hs6 Is Not Null AND targeted_products_hs6 <> ""
ORDER BY date_announced DESC;

-- ======================================================================
-- Query 6: Interventions Pending Implementation
-- ======================================================================
-- Name: qryPendingImplementation
SELECT 
    intervention_id,
    state_act_title,
    intervention_type,
    date_announced,
    implementing_jurisdiction_name,
    affected_jurisdictions,
    DateDiff("d", date_announced, Date()) AS days_since_announced
FROM tblGTAInterventions
WHERE implementation_date Is Null AND removal_date Is Null
ORDER BY date_announced;

-- ======================================================================
-- Query 7: Monthly Intervention Trends
-- ======================================================================
-- Name: qryMonthlyTrends
SELECT 
    Format(date_announced, "yyyy-mm") AS announcement_month,
    Count(*) AS total_interventions,
    Sum(IIf(gta_evaluation = "Red", 1, 0)) AS harmful_count,
    Sum(IIf(gta_evaluation = "Green", 1, 0)) AS liberalizing_count
FROM tblGTAInterventions
WHERE date_announced Is Not Null
GROUP BY Format(date_announced, "yyyy-mm")
ORDER BY Format(date_announced, "yyyy-mm") DESC;

-- ======================================================================
-- Query 8: Settings Configuration View
-- ======================================================================
-- Name: qrySettings
SELECT 
    setting_name,
    setting_value,
    description,
    last_updated
FROM tblSettings
ORDER BY setting_name;

-- ======================================================================
-- Query 9: Data Quality Check
-- ======================================================================
-- Name: qryDataQuality
SELECT 
    "Total Records" AS metric,
    Count(*) AS value
FROM tblGTAInterventions
UNION ALL
SELECT 
    "Records with Product Codes" AS metric,
    Count(*) AS value
FROM tblGTAInterventions
WHERE targeted_products_hs6 Is Not Null AND targeted_products_hs6 <> ""
UNION ALL
SELECT 
    "Records with Implementation Date" AS metric,
    Count(*) AS value
FROM tblGTAInterventions
WHERE implementation_date Is Not Null
UNION ALL
SELECT 
    "Records with Removal Date" AS metric,
    Count(*) AS value
FROM tblGTAInterventions
WHERE removal_date Is Not Null;

-- ======================================================================
-- Query 10: Search Interventions by Keyword
-- ======================================================================
-- Name: qrySearchByKeyword
-- Note: Modify the LIKE condition to search for specific terms
SELECT 
    intervention_id,
    state_act_title,
    intervention_type,
    date_announced,
    implementing_jurisdiction_name,
    affected_jurisdictions
FROM tblGTAInterventions
WHERE state_act_title LIKE "*steel*" OR
      intervention_type LIKE "*steel*" OR
      affected_jurisdictions LIKE "*steel*"
ORDER BY date_announced DESC;

-- ======================================================================
-- SYNC LOG QUERIES (User-Accessible Change Log)
-- ======================================================================

-- ======================================================================
-- Query 11: Recent Sync Activity (Last 7 Days)
-- ======================================================================
-- Name: qrySyncActivity
SELECT 
    log_timestamp,
    session_id,
    log_level,
    source_function,
    message
FROM tblSyncLog
WHERE log_timestamp >= DateAdd("d", -7, Date())
ORDER BY log_timestamp DESC;

-- ======================================================================
-- Query 12: Latest Sync Session Summary
-- ======================================================================
-- Name: qryLatestSyncSummary
SELECT 
    session_id,
    Min(log_timestamp) AS sync_start,
    Max(log_timestamp) AS sync_end,
    Count(*) AS total_log_entries,
    Sum(IIf(log_level = "SUCCESS", 1, 0)) AS successful_operations,
    Sum(IIf(log_level = "ERROR", 1, 0)) AS errors,
    Sum(IIf(log_level = "WARNING", 1, 0)) AS warnings,
    Sum(IIf(message LIKE "*Creating new*", 1, 0)) AS new_records,
    Sum(IIf(message LIKE "*Updating*", 1, 0)) AS updated_records,
    Sum(IIf(message LIKE "*unchanged*", 1, 0)) AS unchanged_records
FROM tblSyncLog
WHERE session_id = (SELECT TOP 1 session_id FROM tblSyncLog ORDER BY log_timestamp DESC)
GROUP BY session_id;

-- ======================================================================
-- Query 13: Error Log (Problems Only)
-- ======================================================================
-- Name: qryErrorLog
SELECT 
    log_timestamp,
    session_id,
    source_function,
    message
FROM tblSyncLog
WHERE log_level IN ("ERROR", "WARNING")
ORDER BY log_timestamp DESC;

-- ======================================================================
-- Query 14: Sync Performance History
-- ======================================================================
-- Name: qrySyncPerformance
SELECT 
    session_id,
    Min(log_timestamp) AS sync_start,
    Max(log_timestamp) AS sync_end,
    DateDiff("s", Min(log_timestamp), Max(log_timestamp)) AS duration_seconds,
    Sum(IIf(message LIKE "*Creating new*", 1, 0)) AS new_records,
    Sum(IIf(message LIKE "*Updating*", 1, 0)) AS updated_records,
    Sum(IIf(message LIKE "*unchanged*", 1, 0)) AS unchanged_records
FROM tblSyncLog
WHERE session_id LIKE "SYNC_*"
GROUP BY session_id
ORDER BY sync_start DESC;

-- ======================================================================
-- Query 15: Change Log with Intervention Details
-- ======================================================================
-- Name: qryChangeLogDetailed
SELECT 
    l.log_timestamp,
    l.session_id,
    l.log_level,
    l.message,
    i.intervention_id,
    i.state_act_title,
    i.intervention_type,
    i.implementing_jurisdiction_name
FROM tblSyncLog l
LEFT JOIN tblGTAInterventions i ON l.message LIKE "*" & i.intervention_id & "*"
WHERE l.log_level IN ("SUCCESS") AND l.message LIKE "*intervention*"
ORDER BY l.log_timestamp DESC;

-- ======================================================================
-- Usage Instructions:
-- ======================================================================
/*
TO CREATE THESE QUERIES IN ACCESS:

1. Open your GtaDemo.accdb database
2. Go to Create > Query Design
3. Close the "Show Table" dialog
4. Click "SQL View" button
5. Copy and paste one query at a time
6. Save with the suggested query name (e.g., "qryRecentInterventions")
7. Repeat for each query

CUSTOMIZATION TIPS:

- Modify date ranges in WHERE clauses
- Add/remove fields as needed
- Change sorting options
- Update keyword searches in Query 10
- Adjust grouping criteria for summaries

PERFORMANCE NOTES:

- Queries run faster with proper indexes (already included in schema)
- Large datasets may benefit from date range filters
- Use Parameters for dynamic filtering in production
*/ 