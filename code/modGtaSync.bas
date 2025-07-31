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
' @version 1.1.0
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

Public Function SyncGTA(Optional ByVal pageSize As Long = DEFAULT_PAGE_SIZE) As Boolean
    On Error GoTo ErrHandler
    Dim startTime As Double
    Dim recordsProcessed As Long
    Dim apiKey As String
    Dim jsonResponse As Object
    Dim httpStatus As Long
    Dim requestDetails As String
    
    startTime = Timer
    SyncGTA = False
    
    Debug.Print "[" & Now() & "] ====== STARTING SYNCGTA ======"
    Debug.Print "[" & Now() & "] Initial PageSize parameter: " & pageSize
    
    If pageSize <= 0 Or pageSize > 1000 Then
        pageSize = DEFAULT_PAGE_SIZE
        Debug.Print "[" & Now() & "] PageSize adjusted to default: " & pageSize
    End If
    
    LogMessage "SyncGTA", "Starting GTA data synchronization (PageSize: " & pageSize & ")"
    Debug.Print "[" & Now() & "] Logged operation start"
    
    ' Step 1: Retrieve API key
    Debug.Print "[" & Now() & "] Retrieving API key..."
    apiKey = GetApiKeyFromSettings()
    Debug.Print "[" & Now() & "] Retrieved API key (first 10 chars): " & Left(apiKey, 10) & "... (Length: " & Len(apiKey) & ")"
    
    If Len(apiKey) = 0 Then
        Debug.Print "[" & Now() & "] ERROR: Empty API key detected"
        Err.Raise vbObjectError + 1001, "SyncGTA", "API Key not found in settings table. Please configure APIKey in " & SETTINGS_TABLE
    End If
    
    ' Step 2: Make API request
    Debug.Print "[" & Now() & "] Preparing API request..."
    Debug.Print "[" & Now() & "] API Base URL: " & API_BASE_URL
    Debug.Print "[" & Now() & "] API Endpoint: " & API_ENDPOINT
    
    LogMessage "SyncGTA", "Making API request to " & API_BASE_URL & API_ENDPOINT
    
    requestDetails = "API Key: " & Left(apiKey, 5) & "..." & vbCrLf & _
                    "PageSize: " & pageSize & vbCrLf & _
                    "URL: " & API_BASE_URL & API_ENDPOINT
                    
    Debug.Print "[" & Now() & "] Request Details:" & vbCrLf & requestDetails
    
    Debug.Print "[" & Now() & "] Calling MakeApiRequest..."
    Set jsonResponse = MakeApiRequest(apiKey, pageSize, httpStatus)
    Debug.Print "[" & Now() & "] API Response Status: " & httpStatus
    
    If Not jsonResponse Is Nothing Then
        On Error Resume Next
        Debug.Print "[" & Now() & "] API Response (partial): " & Left(jsonResponse.ToString, 200)
        On Error GoTo ErrHandler
    Else
        Debug.Print "[" & Now() & "] WARNING: jsonResponse is Nothing"
    End If
    
    If httpStatus <> 200 Then
        Debug.Print "[" & Now() & "] API ERROR DETAILS:"
        Debug.Print "Status Code: " & httpStatus
        Err.Raise vbObjectError + 1002, "SyncGTA", "API request failed with HTTP status: " & httpStatus & _
                 vbCrLf & "Request Details:" & vbCrLf & requestDetails
    End If
    
    ' Step 3: Process response
    Debug.Print "[" & Now() & "] Processing API response..."
    recordsProcessed = ProcessApiResponse(jsonResponse)
    Debug.Print "[" & Now() & "] Records processed: " & recordsProcessed
    
    ' Step 4: Show success message
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    Debug.Print "[" & Now() & "] Sync completed successfully"
    Debug.Print "[" & Now() & "] Time elapsed: " & Format(elapsedTime, "0.0") & " seconds"
    
    MsgBox "GTA Sync completed successfully!" & vbCrLf & _
           "Records processed: " & recordsProcessed & vbCrLf & _
           "Time elapsed: " & Format(elapsedTime, "0.0") & " seconds", _
           vbInformation, "SGEPT API Sync"
           
    LogMessage "SyncGTA", "Sync completed successfully. Records: " & recordsProcessed & ", Time: " & Format(elapsedTime, "0.0") & "s"
    SyncGTA = True
    Debug.Print "[" & Now() & "] ====== SYNCGTA COMPLETED SUCCESSFULLY ======"
    
    Exit Function
    
ErrHandler:
    Dim errorMsg As String
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    
    Debug.Print "[" & Now() & "] ====== ERROR IN SYNCGTA ======"
    Debug.Print "[" & Now() & "] Error Number: " & Err.Number
    Debug.Print "[" & Now() & "] Error Description: " & Err.Description
    Debug.Print "[" & Now() & "] Error Source: " & Err.source
    
    If httpStatus = 403 Then
        Debug.Print "[" & Now() & "] 403 Forbidden - Detailed Analysis:"
        Debug.Print "1. Verify API key is correct (first 5 chars: " & Left(apiKey, 5) & "...)"
    End If
    
    LogMessage "SyncGTA", "ERROR - " & errorMsg
    
    MsgBox "GTA Sync failed!" & vbCrLf & vbCrLf & _
           errorMsg & vbCrLf & vbCrLf & _
           "Please check your API key and internet connection.", _
           vbCritical, "SGEPT API Sync Error"
           
    SyncGTA = False
    Debug.Print "[" & Now() & "] ====== SYNCGTA ENDED WITH ERRORS ======"
End Function

' ============================================= '
' Private Helper Functions
' ============================================= '

Private Function GetApiKeyFromSettings() As String
    On Error GoTo ErrHandler
    Dim rs As Object
    Dim sql As String
    Dim rawKey As String
    
    Debug.Print "[" & Now() & "] Entering GetApiKeyFromSettings()"
    GetApiKeyFromSettings = ""
    rawKey = ""
    
    sql = "SELECT setting_value FROM tblSettings WHERE setting_name = 'APIKey'"
    Debug.Print "[" & Now() & "] Executing SQL: " & sql
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        rawKey = Nz(rs("setting_value"), "")
        GetApiKeyFromSettings = "APIKey " & rawKey
        Debug.Print "[" & Now() & "] Retrieved raw API key (length): " & Len(rawKey) & " characters"
    Else
        Debug.Print "[" & Now() & "] No records found for API key setting"
    End If
    
    rs.Close
    Set rs = Nothing
    
    Debug.Print "[" & Now() & "] Exiting GetApiKeyFromSettings() with value: " & Left(GetApiKeyFromSettings, 10) & "..."
    Exit Function
    
ErrHandler:
    Debug.Print "[" & Now() & "] ERROR in GetApiKeyFromSettings:"
    Debug.Print "   Error Number: " & Err.Number
    Debug.Print "   Error Description: " & Err.Description
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    GetApiKeyFromSettings = ""
    Debug.Print "[" & Now() & "] Exiting after error with empty string"
End Function

Private Function MakeApiRequest(ByVal apiKey As String, ByVal pageSize As Long, ByRef httpStatus As Long) As Object
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonPayload As String
    Dim responseText As String
    Dim startTime As Double
    
    startTime = Timer
    Debug.Print "[" & Now() & "] ====== STARTING API REQUEST ======"
    Debug.Print "[" & Now() & "] API Key (first 5 chars): " & Left(apiKey, 5) & "..."
    Debug.Print "[" & Now() & "] PageSize: " & pageSize
    
    Set MakeApiRequest = Nothing
    httpStatus = 0
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    jsonPayload = "{" & _
                  """limit"":" & pageSize & "," & _
                  """sorting"":""-date_announced""," & _
                  """request_data"":{" & _
                  """mast_chapters"":[" & MAST_CHAPTER_D & "]" & _
                  "}" & _
                  "}"
    
    Debug.Print "[" & Now() & "] Request Payload: " & jsonPayload
    
    With http
        Dim apiUrl As String
        apiUrl = API_BASE_URL & API_ENDPOINT
        Debug.Print "[" & Now() & "] API Endpoint: " & apiUrl
        
        .Open "POST", apiUrl, False
        
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", apiKey
        .setRequestHeader "User-Agent", "SGEPT-Access-Integration/1.0"
        
        Debug.Print "[" & Now() & "] Sending request..."
        .Send jsonPayload
        
        DoEvents
        
        httpStatus = .Status
        responseText = .responseText
        
        Debug.Print "[" & Now() & "] Response Status: " & httpStatus
        Debug.Print "[" & Now() & "] Response Time: " & Format(Timer - startTime, "0.00") & " seconds"
        
        If Len(responseText) > 0 Then
            Debug.Print "[" & Now() & "] Response (first 200 chars): " & Left(responseText, 200)
        Else
            Debug.Print "[" & Now() & "] Empty response received"
        End If
    End With
    
    If httpStatus = 200 Then
        If Len(responseText) > 0 Then
            Set MakeApiRequest = JsonConverter.ParseJson(responseText)
            Debug.Print "[" & Now() & "] Successfully parsed JSON response"
        Else
            Debug.Print "[" & Now() & "] ERROR: Empty response from API"
            Err.Raise vbObjectError + 1003, "MakeApiRequest", "Empty response from API"
        End If
    Else
        Debug.Print "[" & Now() & "] API ERROR DETAILS:"
        Debug.Print "Status Code: " & httpStatus
    End If
    
    Debug.Print "[" & Now() & "] ====== API REQUEST COMPLETED ======"
    Set http = Nothing
    Exit Function
    
ErrHandler:
    Debug.Print "[" & Now() & "] ERROR in MakeApiRequest:"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    
    If Not http Is Nothing Then
        Debug.Print "HTTP Status: " & http.Status
        Debug.Print "Response Text: " & Left(http.responseText, 200)
        Set http = Nothing
    End If
    
    Debug.Print "[" & Now() & "] ====== API REQUEST FAILED ======"
    Err.Raise Err.Number, "MakeApiRequest", Err.Description
End Function

Private Function ProcessApiResponse(ByVal jsonResponse As Object) As Long
    On Error GoTo ErrHandler
    
    Dim interventionsArray As Object
    Dim intervention As Object
    Dim recordCount As Long
    Dim i As Long
    Dim rs As Object
    Dim sql As String
    
    ProcessApiResponse = 0
    
    ' Check response structure
    If TypeName(jsonResponse) = "Collection" Then
        Set interventionsArray = jsonResponse
    ElseIf jsonResponse.Exists("results") Then
        Set interventionsArray = jsonResponse("results")
    Else
        Err.Raise vbObjectError + 1004, "ProcessApiResponse", "API response structure not recognized"
    End If
    
    recordCount = interventionsArray.Count
    
    If recordCount = 0 Then
        LogMessage "ProcessApiResponse", "No interventions returned from API"
        Exit Function
    End If
    
    ' DEBUG: Print first item structure
    Debug.Print "===== FIRST INTERVENTION STRUCTURE ====="
    PrintJsonStructure interventionsArray(1), 0
    Debug.Print "======================================="
    
    ' Open recordset for insertion
    sql = "SELECT * FROM " & INTERVENTIONS_TABLE & " WHERE 1=0"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    ' Process each intervention
    For i = 1 To recordCount
        Set intervention = interventionsArray(i)
        
        If i Mod 10 = 0 Then DoEvents
        
        InsertInterventionRecord rs, intervention
        
        ProcessApiResponse = ProcessApiResponse + 1
    Next i
    
    rs.Close
    Set rs = Nothing
    Set interventionsArray = Nothing
    
    LogMessage "ProcessApiResponse", "Processed " & ProcessApiResponse & " intervention records"
    
    Exit Function
    
ErrHandler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Err.Raise Err.Number, "ProcessApiResponse", Err.Description
End Function

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
    
    ' Check if intervention object is valid
    If intervention Is Nothing Then
        LogMessage "InsertInterventionRecord", "Warning: Null intervention object received"
        Exit Sub
    End If
    
    ' Extract intervention ID
    If Not intervention.Exists("intervention_id") Then
        LogMessage "InsertInterventionRecord", "Warning: Intervention missing ID, skipping"
        Exit Sub
    End If
    
    interventionId = CLng(intervention("intervention_id"))
    
    ' Check if record exists
    Set existingRs = CurrentDb.OpenRecordset("SELECT * FROM " & INTERVENTIONS_TABLE & " WHERE intervention_id = " & interventionId)
    
    isNewRecord = existingRs.EOF
    hasChanges = False
    
    If isNewRecord Then
        rs.AddNew
        hasChanges = True
        LogMessageWithId "InsertInterventionRecord", "Creating new intervention ID: " & interventionId, interventionId
    Else
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
    
    ' === UPDATE FIELDS ===
    If isNewRecord Then
        rs("intervention_id") = interventionId
    End If
    
    ' Basic information
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
    
    ' Description (available with full API key)
    If intervention.Exists("intervention_description") Then
        If isNewRecord Then
            rs("intervention_description") = Left(CStr(intervention("intervention_description")), 1000)
        Else
            existingRs("intervention_description") = Left(CStr(intervention("intervention_description")), 1000)
        End If
    End If
    
    ' Dates
    If intervention.Exists("date_announced") Then
        If isNewRecord Then
            rs("date_announced") = CDate(intervention("date_announced"))
        Else
            existingRs("date_announced") = CDate(intervention("date_announced"))
        End If
    End If
    
    If intervention.Exists("date_implemented") Then
        If Not IsNull(intervention("date_implemented")) Then
            If isNewRecord Then
                rs("implementation_date") = CDate(intervention("date_implemented"))
            Else
                existingRs("implementation_date") = CDate(intervention("date_implemented"))
            End If
        End If
    End If
    
    If intervention.Exists("date_removed") Then
        If Not IsNull(intervention("date_removed")) Then
            If isNewRecord Then
                rs("removal_date") = CDate(intervention("date_removed"))
            Else
                existingRs("removal_date") = CDate(intervention("date_removed"))
            End If
        End If
    End If
    
    ' Update last updated timestamp
    If isNewRecord Then
        rs("last_updated") = Now()
    Else
        existingRs("last_updated") = Now()
    End If
    
    ' Jurisdictions
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
    
    ' Products and sectors
    If intervention.Exists("affected_products") Then
        targetedProducts = ExtractProductCodes(intervention("affected_products"))
        If isNewRecord Then
            rs("targeted_products_hs6") = Left(targetedProducts, 1000)
        Else
            existingRs("targeted_products_hs6") = Left(targetedProducts, 1000)
        End If
    End If
    
    If intervention.Exists("affected_sectors") Then
        targetedSectors = ExtractSectorCodes(intervention("affected_sectors"))
        If isNewRecord Then
            rs("targeted_sectors_cpc3") = Left(targetedSectors, 500)
        Else
            existingRs("targeted_sectors_cpc3") = Left(targetedSectors, 500)
        End If
    End If
    
    ' Administrative information
    If isNewRecord Then
        rs("sync_source") = "SGEPT_API"
    Else
        existingRs("sync_source") = "SGEPT_API"
    End If
    
    ' Source (available with full API key)
    If intervention.Exists("source") Then
        If isNewRecord Then
            rs("source") = Left(CStr(intervention("source")), 500)
        Else
            existingRs("source") = Left(CStr(intervention("source")), 500)
        End If
    End If
    
    ' Save changes
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

Private Function RecordHasChanges(ByRef existingRs As Object, ByVal intervention As Object) As Boolean
    On Error GoTo ErrHandler
    
    Dim newValue As String
    Dim existingValue As String
    
    RecordHasChanges = False
    
    ' Check changes in main fields
    If intervention.Exists("state_act_title") Then
        newValue = Left(CStr(intervention("state_act_title")), 255)
        existingValue = Nz(existingRs("state_act_title"), "")
        If newValue <> existingValue Then
            RecordHasChanges = True
            Exit Function
        End If
    End If
    
    ' Check other changes...
    ' (Rest of function remains as is)
    
    Exit Function
    
ErrHandler:
    RecordHasChanges = True
End Function

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

Private Sub LogMessage(ByVal source As String, ByVal message As String)
    On Error GoTo ErrHandler
    
    Dim logRs As Object
    Dim logEntry As String
    
    logEntry = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " [" & source & "] " & message
    Debug.Print logEntry
    
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
    If Not logRs Is Nothing Then
        logRs.Close
        Set logRs = Nothing
    End If
    Debug.Print logEntry
End Sub

Private Function DetermineLogLevel(ByVal message As String) As String
    Dim upperMsg As String
    upperMsg = UCase(message)
    
    If InStr(upperMsg, "ERROR") > 0 Or InStr(upperMsg, "FAILED") > 0 Then
        DetermineLogLevel = "ERROR"
    ElseIf InStr(upperMsg, "WARNING") > 0 Or InStr(upperMsg, "MISSING") > 0 Then
        DetermineLogLevel = "WARNING"
    ElseIf InStr(upperMsg, "COMPLETED") > 0 Or InStr(upperMsg, "SUCCESS") > 0 Then
        DetermineLogLevel = "SUCCESS"
    Else
        DetermineLogLevel = "INFO"
    End If
End Function

Private Function GetCurrentSessionId() As String
    Static sessionId As String
    
    If Len(sessionId) = 0 Then
        sessionId = "SYNC_" & Format(Now(), "yyyymmdd_hhnnss") & "_" & Int(Rnd() * 1000)
    End If
    
    GetCurrentSessionId = sessionId
End Function

Private Sub LogMessageWithId(ByVal source As String, ByVal message As String, ByVal interventionId As Long)
    On Error GoTo ErrHandler
    
    Dim logRs As Object
    Dim logEntry As String
    
    logEntry = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " [" & source & "] " & message
    Debug.Print logEntry
    
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
    If Not logRs Is Nothing Then
        logRs.Close
        Set logRs = Nothing
    End If
    Debug.Print logEntry
End Sub

Private Sub PrintJsonStructure(obj As Object, level As Integer)
    Dim key As Variant
    Dim indent As String
    Dim i As Integer
    
    indent = Space(level * 2)
    
    If TypeName(obj) = "Dictionary" Then
        For Each key In obj.Keys
            Debug.Print indent & key & ": " & TypeName(obj(key))
            If TypeName(obj(key)) = "Dictionary" Or TypeName(obj(key)) = "Collection" Then
                PrintJsonStructure obj(key), level + 1
            End If
        Next key
    ElseIf TypeName(obj) = "Collection" Then
        For i = 1 To obj.Count
            Debug.Print indent & "[" & i & "]: " & TypeName(obj(i))
            If TypeName(obj(i)) = "Dictionary" Or TypeName(obj(i)) = "Collection" Then
                PrintJsonStructure obj(i), level + 1
            End If
        Next i
    End If
End Sub

