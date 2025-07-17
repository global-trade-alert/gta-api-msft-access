Attribute VB_Name = "modGtaSync"
Option Explicit

''
' SGEPT API Access Integration for Microsoft Access
' 
' This module provides functionality to sync GTA intervention data
' from the SGEPT API into local Access database tables.
'
' Requirements:
' - JsonConverter.bas module (VBA-JSON by Tim Hall)
' - Reference to Microsoft Scripting Runtime
' - Reference to Microsoft XML, v6.0 (for HTTP requests)
' - Tables: tblSettings, tblGTAInterventions
' 
' API Access Levels:
' - Demo API Key: Basic fields (intervention_id, state_act_title, dates, jurisdictions, etc.)
' - Full API Key: Includes intervention_description and source fields
'                 (Available upon request for trial purposes)
'
' @author SGEPT Integration Team
' @version 1.0.0
''

' ============================================= '
' Constants
' ============================================= '

' API Configuration
Private Const API_BASE_URL As String = "https://api.globaltradealert.org"
Private Const API_ENDPOINT As String = "/api/v1/data/"
Private Const API_TIMEOUT As Long = 30000  ' 30 seconds

' Pagination and Filtering
Private Const DEFAULT_PAGE_SIZE As Long = 50
Private Const MAST_CHAPTER_D As Long = 4  ' Contingent trade-protective measures

' Database Configuration
Private Const SETTINGS_TABLE As String = "tblSettings"
Private Const INTERVENTIONS_TABLE As String = "tblGTAInterventions"
Private Const SYNC_LOG_TABLE As String = "tblSyncLog"
Private Const API_KEY_SETTING As String = "APIKey"

' ============================================= '
' Public Entry Points
' ============================================= '

''
' Main synchronization function - pulls latest GTA interventions
' from SGEPT API and updates local database
'
' @method SyncGTA
' @param {Long} Optional PageSize - Number of records to fetch (default: 50)
' @return {Boolean} Success status
''
Public Function SyncGTA(Optional ByVal PageSize As Long = DEFAULT_PAGE_SIZE) As Boolean
    On Error GoTo ErrHandler
    
    Dim startTime As Double
    Dim recordsProcessed As Long
    Dim apiKey As String
    Dim jsonResponse As Object
    Dim httpStatus As Long
    
    ' Initialize
    startTime = Timer
    SyncGTA = False
    
    ' Validate page size
    If PageSize <= 0 Or PageSize > 1000 Then
        PageSize = DEFAULT_PAGE_SIZE
    End If
    
    ' Log operation start
    LogMessage "SyncGTA", "Starting GTA data synchronization (PageSize: " & PageSize & ")"
    
    ' Step 1: Retrieve API key from settings
    apiKey = GetApiKeyFromSettings()
    If Len(apiKey) = 0 Then
        Err.Raise vbObjectError + 1001, "SyncGTA", "API Key not found in settings table. Please configure APIKey in " & SETTINGS_TABLE
    End If
    
    ' Step 2: Make API request
    LogMessage "SyncGTA", "Making API request to " & API_BASE_URL & API_ENDPOINT
    Set jsonResponse = MakeApiRequest(apiKey, PageSize, httpStatus)
    
    If httpStatus <> 200 Then
        Err.Raise vbObjectError + 1002, "SyncGTA", "API request failed with HTTP status: " & httpStatus
    End If
    
    ' Step 3: Process response and update database
    recordsProcessed = ProcessApiResponse(jsonResponse)
    
    ' Step 4: Show success message
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    MsgBox "GTA Sync completed successfully!" & vbCrLf & _
           "Records processed: " & recordsProcessed & vbCrLf & _
           "Time elapsed: " & Format(elapsedTime, "0.0") & " seconds", _
           vbInformation, "SGEPT API Sync"
    
    LogMessage "SyncGTA", "Sync completed successfully. Records: " & recordsProcessed & ", Time: " & Format(elapsedTime, "0.0") & "s"
    SyncGTA = True
    
    Exit Function
    
ErrHandler:
    Dim errorMsg As String
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    
    LogMessage "SyncGTA", "ERROR - " & errorMsg
    
    MsgBox "GTA Sync failed!" & vbCrLf & vbCrLf & _
           errorMsg & vbCrLf & vbCrLf & _
           "Please check your API key and internet connection.", _
           vbCritical, "SGEPT API Sync Error"
    
    SyncGTA = False
End Function

' ============================================= '
' Private Helper Functions
' ============================================= '

''
' Retrieve API key from settings table
'
' @method GetApiKeyFromSettings
' @return {String} API key or empty string if not found
''
Private Function GetApiKeyFromSettings() As String
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim sql As String
    
    GetApiKeyFromSettings = ""
    
    ' Query settings table for API key
    sql = "SELECT SettingValue FROM " & SETTINGS_TABLE & " WHERE SettingName = '" & API_KEY_SETTING & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        GetApiKeyFromSettings = Nz(rs("SettingValue"), "")
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrHandler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    ' Return empty string on error
    GetApiKeyFromSettings = ""
End Function

''
' Make HTTP request to SGEPT API
'
' @method MakeApiRequest
' @param {String} apiKey - API authentication key
' @param {Long} pageSize - Number of records to request
' @param {Long} ByRef httpStatus - HTTP status code returned
' @return {Object} Parsed JSON response
''
Private Function MakeApiRequest(ByVal apiKey As String, ByVal pageSize As Long, ByRef httpStatus As Long) As Object
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonPayload As String
    Dim responseText As String
    
    Set MakeApiRequest = Nothing
    httpStatus = 0
    
    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    ' Build JSON payload for MAST chapter D filter
    jsonPayload = "{" & _
                  """limit"":" & pageSize & "," & _
                  """sorting"":[""-date_announced""]," & _
                  """request_data"":{" & _
                  """mast_chapters"":[" & MAST_CHAPTER_D & "]" & _
                  "}" & _
                  "}"
    
    ' Configure and send request
    With http
        .Open "POST", API_BASE_URL & API_ENDPOINT, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "APIKey", apiKey
        .setRequestHeader "User-Agent", "SGEPT-Access-Integration/1.0"
        .Send jsonPayload
        
        ' Allow UI updates during request
        DoEvents
        
        httpStatus = .Status
        responseText = .responseText
    End With
    
    ' Parse JSON response if successful
    If httpStatus = 200 Then
        If Len(responseText) > 0 Then
            Set MakeApiRequest = JsonConverter.ParseJson(responseText)
        Else
            Err.Raise vbObjectError + 1003, "MakeApiRequest", "Empty response from API"
        End If
    End If
    
    Set http = Nothing
    Exit Function
    
ErrHandler:
    If Not http Is Nothing Then Set http = Nothing
    Err.Raise Err.Number, "MakeApiRequest", Err.Description
End Function

''
' Process API response and update database
'
' @method ProcessApiResponse
' @param {Object} jsonResponse - Parsed JSON response from API
' @return {Long} Number of records processed
''
Private Function ProcessApiResponse(ByVal jsonResponse As Object) As Long
    On Error GoTo ErrHandler
    
    Dim resultsArray As Object
    Dim intervention As Object
    Dim recordCount As Long
    Dim i As Long
    Dim rs As Object
    Dim sql As String
    
    ProcessApiResponse = 0
    
    ' Validate response structure
    If Not jsonResponse.Exists("results") Then
        Err.Raise vbObjectError + 1004, "ProcessApiResponse", "API response missing 'results' array"
    End If
    
    Set resultsArray = jsonResponse("results")
    recordCount = resultsArray.Count
    
    If recordCount = 0 Then
        LogMessage "ProcessApiResponse", "No interventions returned from API"
        Exit Function
    End If
    
    ' Open recordset for inserting data
    sql = "SELECT * FROM " & INTERVENTIONS_TABLE & " WHERE 1=0"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    ' Process each intervention
    For i = 1 To recordCount
        Set intervention = resultsArray(i)
        
        ' Allow UI updates during processing
        If i Mod 10 = 0 Then DoEvents
        
        ' Insert/update intervention record
        InsertInterventionRecord rs, intervention
        
        ProcessApiResponse = ProcessApiResponse + 1
    Next i
    
    rs.Close
    Set rs = Nothing
    Set resultsArray = Nothing
    
    LogMessage "ProcessApiResponse", "Processed " & ProcessApiResponse & " intervention records"
    
    Exit Function
    
ErrHandler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Err.Raise Err.Number, "ProcessApiResponse", Err.Description
End Function

''
' Insert or update intervention record in database
'
' @method InsertInterventionRecord
' @param {Object} rs - Open recordset for interventions table
' @param {Object} intervention - JSON intervention object
''
Private Sub InsertInterventionRecord(ByRef rs As Object, ByVal intervention As Object)
    On Error GoTo ErrHandler
    
    Dim interventionId As Long
    Dim existingRs As Object
    Dim implementingJurisdictions As String
    Dim affectedJurisdictions As String
    Dim targetedProducts As String
    Dim targetedSectors As String
    Dim isNewRecord As Boolean
    Dim hasChanges As Boolean
    
    ' Extract intervention ID
    If Not intervention.Exists("intervention_id") Then
        LogMessage "InsertInterventionRecord", "Warning: Intervention missing ID, skipping"
        Exit Sub
    End If
    
    interventionId = CLng(intervention("intervention_id"))
    
    ' Check if record already exists and get current data for comparison
    Set existingRs = CurrentDb.OpenRecordset("SELECT * FROM " & INTERVENTIONS_TABLE & " WHERE intervention_id = " & interventionId)
    
    isNewRecord = existingRs.EOF
    hasChanges = False
    
    If isNewRecord Then
        ' === NEW RECORD: INSERT ===
        rs.AddNew
        hasChanges = True
        LogMessageWithId "InsertInterventionRecord", "Creating new intervention ID: " & interventionId, interventionId
    Else
        ' === EXISTING RECORD: CHECK FOR CHANGES ===
        hasChanges = RecordHasChanges(existingRs, intervention)
        
        If hasChanges Then
            existingRs.Edit
            LogMessageWithId "InsertInterventionRecord", "Updating intervention ID: " & interventionId & " (changes detected)", interventionId
        Else
            LogMessageWithId "InsertInterventionRecord", "Intervention ID " & interventionId & " unchanged, skipping", interventionId
            existingRs.Close
            Set existingRs = Nothing
            Exit Sub
        End If
    End If
    
    ' === POPULATE/UPDATE FIELDS (COMMON CODE FOR INSERT AND UPDATE) ===
    If isNewRecord Then
        rs("intervention_id") = interventionId
    End If
    
    ' === GROUP 1: CORE INTERVENTION INFORMATION ===
    ' Available with demo API key
    If intervention.Exists("state_act_title") Then
        If isNewRecord Then
            rs("state_act_title") = Left(CStr(intervention("state_act_title")), 255)
        Else
            existingRs("state_act_title") = Left(CStr(intervention("state_act_title")), 255)
        End If
    End If
    
    If intervention.Exists("intervention_type") Then
        If isNewRecord Then
            rs("intervention_type") = Left(CStr(intervention("intervention_type")), 100)
        Else
            existingRs("intervention_type") = Left(CStr(intervention("intervention_type")), 100)
        End If
    End If
    
    If intervention.Exists("gta_evaluation") Then
        If isNewRecord Then
            rs("gta_evaluation") = Left(CStr(intervention("gta_evaluation")), 50)
        Else
            existingRs("gta_evaluation") = Left(CStr(intervention("gta_evaluation")), 50)
        End If
    End If
    
    ' Requires full API access (available upon request for trial purposes)
    If intervention.Exists("intervention_description") Then
        If isNewRecord Then
            rs("intervention_description") = Left(CStr(intervention("intervention_description")), 1000)
        Else
            existingRs("intervention_description") = Left(CStr(intervention("intervention_description")), 1000)
        End If
    End If
    
    ' === GROUP 2: KEY DATES ===
    If intervention.Exists("date_announced") Then
        If isNewRecord Then
            rs("date_announced") = CDate(intervention("date_announced"))
        Else
            existingRs("date_announced") = CDate(intervention("date_announced"))
        End If
    End If
    
    If intervention.Exists("implementation_date") Then
        If isNewRecord Then
            rs("implementation_date") = CDate(intervention("implementation_date"))
        Else
            existingRs("implementation_date") = CDate(intervention("implementation_date"))
        End If
    End If
    
    If intervention.Exists("removal_date") Then
        If isNewRecord Then
            rs("removal_date") = CDate(intervention("removal_date"))
        Else
            existingRs("removal_date") = CDate(intervention("removal_date"))
        End If
    End If
    
    If isNewRecord Then
        rs("last_updated") = Now()
    Else
        existingRs("last_updated") = Now()
    End If
    
    ' === GROUP 3: GEOGRAPHIC SCOPE ===
    If intervention.Exists("implementing_jurisdictions") Then
        implementingJurisdictions = ExtractJurisdictionNames(intervention("implementing_jurisdictions"))
        If isNewRecord Then
            rs("implementing_jurisdiction_name") = Left(implementingJurisdictions, 255)
        Else
            existingRs("implementing_jurisdiction_name") = Left(implementingJurisdictions, 255)
        End If
    End If
    
    If intervention.Exists("affected_jurisdictions") Then
        affectedJurisdictions = ExtractJurisdictionNames(intervention("affected_jurisdictions"))
        If isNewRecord Then
            rs("affected_jurisdictions") = Left(affectedJurisdictions, 500)
        Else
            existingRs("affected_jurisdictions") = Left(affectedJurisdictions, 500)
        End If
    End If
    
    ' === GROUP 4: ECONOMIC TARGETING ===
    If intervention.Exists("targeted_products") Then
        targetedProducts = ExtractProductCodes(intervention("targeted_products"))
        If isNewRecord Then
            rs("targeted_products_hs6") = Left(targetedProducts, 1000)
        Else
            existingRs("targeted_products_hs6") = Left(targetedProducts, 1000)
        End If
    End If
    
    If intervention.Exists("targeted_sectors") Then
        targetedSectors = ExtractSectorCodes(intervention("targeted_sectors"))
        If isNewRecord Then
            rs("targeted_sectors_cpc3") = Left(targetedSectors, 500)
        Else
            existingRs("targeted_sectors_cpc3") = Left(targetedSectors, 500)
        End If
    End If
    
    ' === GROUP 5: ADMINISTRATIVE ===
    If isNewRecord Then
        rs("sync_source") = "SGEPT_API"
    Else
        existingRs("sync_source") = "SGEPT_API"
    End If
    
    ' Requires full API access (available upon request for trial purposes)
    If intervention.Exists("source") Then
        If isNewRecord Then
            rs("source") = Left(CStr(intervention("source")), 500)
        Else
            existingRs("source") = Left(CStr(intervention("source")), 500)
        End If
    End If
    
    ' === SAVE CHANGES ===
    If isNewRecord Then
        rs.Update
    Else
        existingRs.Update
    End If
    
    existingRs.Close
    Set existingRs = Nothing
    
    Exit Sub
    
ErrHandler:
    If Not existingRs Is Nothing Then
        existingRs.Close
        Set existingRs = Nothing
    End If
    LogMessage "InsertInterventionRecord", "Error processing intervention: " & Err.Description
End Sub

''
' Check if intervention data has changed compared to existing record
'
' @method RecordHasChanges
' @param {Object} existingRs - Existing recordset positioned on current record
' @param {Object} intervention - JSON intervention object from API
' @return {Boolean} True if changes detected, False if identical
''
Private Function RecordHasChanges(ByRef existingRs As Object, ByVal intervention As Object) As Boolean
    On Error GoTo ErrHandler
    
    Dim newValue As String
    Dim existingValue As String
    
    RecordHasChanges = False
    
    ' === CHECK CORE INTERVENTION INFORMATION ===
    ' Check state_act_title
    If intervention.Exists("state_act_title") Then
        newValue = Left(CStr(intervention("state_act_title")), 255)
        existingValue = Nz(existingRs("state_act_title"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check intervention_type
    If intervention.Exists("intervention_type") Then
        newValue = Left(CStr(intervention("intervention_type")), 100)
        existingValue = Nz(existingRs("intervention_type"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check gta_evaluation
    If intervention.Exists("gta_evaluation") Then
        newValue = Left(CStr(intervention("gta_evaluation")), 50)
        existingValue = Nz(existingRs("gta_evaluation"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check intervention_description (if available with full API access)
    If intervention.Exists("intervention_description") Then
        newValue = Left(CStr(intervention("intervention_description")), 1000)
        existingValue = Nz(existingRs("intervention_description"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' === CHECK KEY DATES ===
    ' Check implementation_date
    If intervention.Exists("implementation_date") Then
        If IsNull(existingRs("implementation_date")) Then
            RecordHasChanges = True
            Exit Function
        ElseIf CDate(intervention("implementation_date")) <> existingRs("implementation_date") Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check removal_date
    If intervention.Exists("removal_date") Then
        If IsNull(existingRs("removal_date")) Then
            RecordHasChanges = True
            Exit Function
        ElseIf CDate(intervention("removal_date")) <> existingRs("removal_date") Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' === CHECK GEOGRAPHIC SCOPE ===
    ' Check implementing_jurisdictions
    If intervention.Exists("implementing_jurisdictions") Then
        newValue = Left(ExtractJurisdictionNames(intervention("implementing_jurisdictions")), 255)
        existingValue = Nz(existingRs("implementing_jurisdiction_name"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check affected_jurisdictions
    If intervention.Exists("affected_jurisdictions") Then
        newValue = Left(ExtractJurisdictionNames(intervention("affected_jurisdictions")), 500)
        existingValue = Nz(existingRs("affected_jurisdictions"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' === CHECK ECONOMIC TARGETING ===
    ' Check targeted_products
    If intervention.Exists("targeted_products") Then
        newValue = Left(ExtractProductCodes(intervention("targeted_products")), 1000)
        existingValue = Nz(existingRs("targeted_products_hs6"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check targeted_sectors
    If intervention.Exists("targeted_sectors") Then
        newValue = Left(ExtractSectorCodes(intervention("targeted_sectors")), 500)
        existingValue = Nz(existingRs("targeted_sectors_cpc3"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' === CHECK ADMINISTRATIVE ===
    ' Check source (if available with full API access)
    If intervention.Exists("source") Then
        newValue = Left(CStr(intervention("source")), 500)
        existingValue = Nz(existingRs("source"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' If we reach here, no changes were detected
    RecordHasChanges = False
    
    Exit Function
    
ErrHandler:
    ' On error, assume changes exist to be safe
    RecordHasChanges = True
End Function

''
' Extract jurisdiction names from implementing_jurisdictions array
'
' @method ExtractJurisdictionNames
' @param {Object} jurisdictions - Collection of jurisdiction objects
' @return {String} Comma-separated jurisdiction names
''
Private Function ExtractJurisdictionNames(ByVal jurisdictions As Object) As String
    On Error GoTo ErrHandler
    
    Dim jurisdiction As Object
    Dim names As String
    Dim i As Long
    
    ExtractJurisdictionNames = ""
    
    If TypeName(jurisdictions) = "Collection" Then
        For i = 1 To jurisdictions.Count
            Set jurisdiction = jurisdictions(i)
            
            If jurisdiction.Exists("name") Then
                If Len(names) > 0 Then names = names & ", "
                names = names & CStr(jurisdiction("name"))
            End If
        Next i
    End If
    
    ExtractJurisdictionNames = names
    
    Exit Function
    
ErrHandler:
    ExtractJurisdictionNames = "Error extracting names"
End Function

''
' Extract HS 6-digit product codes from targeted_products array
'
' @method ExtractProductCodes
' @param {Object} products - Collection of product objects
' @return {String} Comma-separated HS codes
''
Private Function ExtractProductCodes(ByVal products As Object) As String
    On Error GoTo ErrHandler
    
    Dim product As Object
    Dim codes As String
    Dim i As Long
    
    ExtractProductCodes = ""
    
    If TypeName(products) = "Collection" Then
        For i = 1 To products.Count
            Set product = products(i)
            
            If product.Exists("hs_code") Then
                If Len(codes) > 0 Then codes = codes & ", "
                codes = codes & CStr(product("hs_code"))
            ElseIf product.Exists("code") Then
                If Len(codes) > 0 Then codes = codes & ", "
                codes = codes & CStr(product("code"))
            End If
        Next i
    End If
    
    ExtractProductCodes = codes
    
    Exit Function
    
ErrHandler:
    ExtractProductCodes = "Error extracting product codes"
End Function

''
' Extract CPC 3-digit sector codes from targeted_sectors array
'
' @method ExtractSectorCodes
' @param {Object} sectors - Collection of sector objects
' @return {String} Comma-separated CPC codes
''
Private Function ExtractSectorCodes(ByVal sectors As Object) As String
    On Error GoTo ErrHandler
    
    Dim sector As Object
    Dim codes As String
    Dim i As Long
    
    ExtractSectorCodes = ""
    
    If TypeName(sectors) = "Collection" Then
        For i = 1 To sectors.Count
            Set sector = sectors(i)
            
            If sector.Exists("cpc_code") Then
                If Len(codes) > 0 Then codes = codes & ", "
                codes = codes & CStr(sector("cpc_code"))
            ElseIf sector.Exists("code") Then
                If Len(codes) > 0 Then codes = codes & ", "
                codes = codes & CStr(sector("code"))
            End If
        Next i
    End If
    
    ExtractSectorCodes = codes
    
    Exit Function
    
ErrHandler:
    ExtractSectorCodes = "Error extracting sector codes"
End Function

''
' Log messages for debugging and audit trail
'
' @method LogMessage
' @param {String} source - Source function/module
' @param {String} message - Log message
''
Private Sub LogMessage(ByVal source As String, ByVal message As String)
    On Error GoTo ErrHandler
    
    Dim logRs As Object
    Dim logEntry As String
    
    ' Create formatted log entry
    logEntry = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " [" & source & "] " & message
    
    ' Always output to immediate window for debugging
    Debug.Print logEntry
    
    ' Try to write to database log table (if it exists)
    Set logRs = CurrentDb.OpenRecordset("SELECT * FROM " & SYNC_LOG_TABLE & " WHERE 1=0")
    
    logRs.AddNew
    logRs("log_timestamp") = Now()
    logRs("source_function") = Left(source, 50)
    logRs("log_level") = DetermineLogLevel(message)
    logRs("message") = Left(message, 500)
    logRs("session_id") = GetCurrentSessionId()
    logRs.Update
    
    logRs.Close
    Set logRs = Nothing
    
    Exit Sub
    
ErrHandler:
    ' If log table doesn't exist or other error, just continue
    ' This ensures backwards compatibility and doesn't break the sync
    If Not logRs Is Nothing Then
        logRs.Close
        Set logRs = Nothing
    End If
    ' Still output to debug window as fallback
    Debug.Print logEntry
End Sub

''
' Determine log level based on message content
'
' @method DetermineLogLevel  
' @param {String} message - Log message
' @return {String} Log level (INFO, WARNING, ERROR, SUCCESS)
''
Private Function DetermineLogLevel(ByVal message As String) As String
    Dim upperMsg As String
    upperMsg = UCase(message)
    
    If InStr(upperMsg, "ERROR") > 0 Or InStr(upperMsg, "FAILED") > 0 Then
        DetermineLogLevel = "ERROR"
    ElseIf InStr(upperMsg, "WARNING") > 0 Or InStr(upperMsg, "MISSING") > 0 Then
        DetermineLogLevel = "WARNING"
    ElseIf InStr(upperMsg, "COMPLETED") > 0 Or InStr(upperMsg, "SUCCESS") > 0 Or InStr(upperMsg, "INSERTED") > 0 Or InStr(upperMsg, "UPDATED") > 0 Then
        DetermineLogLevel = "SUCCESS"
    Else
        DetermineLogLevel = "INFO"
    End If
End Function

''
' Get current session identifier for grouping related log entries
'
' @method GetCurrentSessionId
' @return {String} Session identifier
''
Private Function GetCurrentSessionId() As String
    Static sessionId As String
    
    ' Generate session ID once per VBA session
    If Len(sessionId) = 0 Then
        sessionId = "SYNC_" & Format(Now(), "yyyymmdd_hhnnss") & "_" & Int(Rnd() * 1000)
    End If
    
    GetCurrentSessionId = sessionId
End Function

''
' Log messages with intervention ID for specific record tracking
'
' @method LogMessageWithId
' @param {String} source - Source function/module  
' @param {String} message - Log message
' @param {Long} interventionId - Intervention ID for tracking
''
Private Sub LogMessageWithId(ByVal source As String, ByVal message As String, ByVal interventionId As Long)
    On Error GoTo ErrHandler
    
    Dim logRs As Object
    Dim logEntry As String
    
    ' Create formatted log entry
    logEntry = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " [" & source & "] " & message
    
    ' Always output to immediate window for debugging
    Debug.Print logEntry
    
    ' Try to write to database log table (if it exists)
    Set logRs = CurrentDb.OpenRecordset("SELECT * FROM " & SYNC_LOG_TABLE & " WHERE 1=0")
    
    logRs.AddNew
    logRs("log_timestamp") = Now()
    logRs("source_function") = Left(source, 50)
    logRs("log_level") = DetermineLogLevel(message)
    logRs("message") = Left(message, 500)
    logRs("session_id") = GetCurrentSessionId()
    logRs("intervention_id") = interventionId
    logRs.Update
    
    logRs.Close
    Set logRs = Nothing
    
    Exit Sub
    
ErrHandler:
    ' If log table doesn't exist or other error, just continue
    ' This ensures backwards compatibility and doesn't break the sync
    If Not logRs Is Nothing Then
        logRs.Close
        Set logRs = Nothing
    End If
    ' Still output to debug window as fallback
    Debug.Print logEntry
End Sub 