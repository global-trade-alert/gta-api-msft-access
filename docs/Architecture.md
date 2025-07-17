# SGEPT API Integration Architecture

## System Overview

The SGEPT API integration provides automated synchronization of Global Trade Alert intervention data into Microsoft Access databases. The architecture follows a layered approach with clear separation between data access, business logic, and presentation components.

## Components

### External Systems

**🌐 SGEPT API (api.globaltradealert.org)**
- RESTful API endpoint: `/api/v1/data/`
- Authentication: APIKey header
- Data format: JSON request/response
- Rate limits: Managed by intelligent paging
- Focus: MAST Chapter D trade interventions

**⏰ Windows Task Scheduler (Optional)**
- Enables automated synchronization
- Configurable schedule (daily, weekly, etc.)
- Runs PowerShell scripts or Access macros
- Background execution without user interaction

### Application Layer

**📄 JsonConverter.bas (VBA-JSON Library)**
- Third-party library by Tim Hall (MIT License)
- Handles JSON parsing and serialization
- Converts API responses to VBA objects
- Industry-standard, well-tested implementation

**📄 modGtaSync.bas (Main Integration Logic)**
- Core synchronization orchestration
- API request/response handling
- Smart change detection and UPSERT logic
- Error handling and logging
- Public interface: `SyncGTA()` function

### Data Layer

**🗃️ Microsoft Access Database (GtaDemo.accdb)**

*⚙️ tblSettings (Configuration)*
- API key storage
- Sync preferences (page size, enabled/disabled)
- Last sync timestamp tracking
- User-configurable parameters

*📊 tblGTAInterventions (Main Data)*
- 15+ fields organized in logical groups:
  - Core Information (ID, title, type, evaluation)
  - Key Dates (announced, implementation, removal)
  - Geographic Scope (implementing/affected jurisdictions)
  - Economic Targeting (HS6 products, CPC3 sectors)
  - Administrative (source URLs, sync metadata)

*📝 tblSyncLog (Audit Trail)*
- Complete operation history
- Session-based grouping
- Log levels (SUCCESS, INFO, WARNING, ERROR)
- Performance monitoring data
- User-accessible change tracking

### Presentation Layer

**📈 Pre-Built Queries**
- Ready-to-use analytical reports
- Common business scenarios covered
- Performance optimized with proper indexes
- Easy customization for specific needs

## Data Flow Process

### 1. Sync Initiation
```
User executes: SyncGTA
OR
Task Scheduler triggers automated run
```

### 2. Configuration Reading
```
modGtaSync → tblSettings
• Retrieves API key
• Gets page size preference
• Checks sync enabled status
```

### 3. API Request
```
modGtaSync → SGEPT API
• HTTP POST to /api/v1/data/
• JSON payload with MAST Chapter D filter
• APIKey authentication header
• Configurable result limit (default: 50)
```

### 4. Response Processing
```
SGEPT API → modGtaSync
• JSON response with intervention array
• Error handling for HTTP failures
• Rate limit compliance
```

### 5. Data Parsing
```
modGtaSync → JsonConverter.bas
• Parse JSON response to VBA objects
• Extract intervention records
• Handle missing or malformed fields
```

### 6. Smart Data Management
```
JsonConverter.bas → modGtaSync → tblGTAInterventions
• NEW RECORDS: Insert with full data
• EXISTING RECORDS: Compare all fields
• CHANGES DETECTED: Update with new values
• NO CHANGES: Skip to optimize performance
```

### 7. Audit Logging
```
modGtaSync → tblSyncLog
• Log all operations (insert/update/skip)
• Session-based grouping
• Performance metrics
• Error details
```

### 8. User Access
```
tblGTAInterventions + tblSyncLog → Pre-Built Queries
• Analytical reports
• Change summaries
• Error diagnostics
• Performance monitoring
```

## Key Design Principles

### Reliability
- **Comprehensive error handling** at every integration point
- **Graceful degradation** when optional components fail
- **Transaction safety** with proper cleanup
- **Logging** for troubleshooting and audit

### Performance
- **Smart change detection** prevents unnecessary database writes
- **Configurable page sizes** balance speed vs. completeness
- **Database indexes** on frequently queried fields
- **Session-based logging** for efficient grouping

### Maintainability
- **Modular design** with clear component boundaries
- **Extensive documentation** within VBA code
- **Standardized naming conventions** throughout
- **Separation of concerns** between API, business logic, and data

### User Experience
- **One-command operation** (`SyncGTA`)
- **Clear progress feedback** via logging
- **Transparent change tracking** shows what happened
- **Ready-to-use queries** for immediate value

### Security
- **API key protection** in encrypted Access database
- **Input validation** on all API responses
- **SQL injection prevention** through parameterized queries
- **Error message sanitization** to prevent information leakage

## Deployment Considerations

### Development Environment
- Microsoft Access 2016+ with VBA enabled
- Internet connectivity for API access
- Administrative rights for VBA reference installation
- SGEPT API key (demo or full access)

### Production Environment
- Stable internet connection with HTTPS support
- Regular database backup procedures
- Monitoring of sync log for operational issues
- Consider automation via Task Scheduler

### Scalability
- Current design optimized for single-user/small team usage
- API rate limits naturally constrain maximum sync frequency
- Database size management through archival procedures
- Horizontal scaling possible through multiple database instances

## Technology Stack

| Component | Technology | Version | License |
|-----------|------------|---------|---------|
| Database | Microsoft Access | 2016+ | Commercial |
| Runtime | VBA (Visual Basic for Applications) | Built-in | Commercial |
| JSON Library | VBA-JSON by Tim Hall | 2.3.1 | MIT |
| API | SGEPT RESTful API | v1 | Terms of Service |
| Automation | Windows Task Scheduler | Built-in | Commercial |
| Documentation | Markdown + Mermaid | Latest | Open Source |

## Future Enhancements

### Planned Features
- **Multi-chapter support** beyond MAST Chapter D
- **Real-time notifications** for high-priority interventions
- **Advanced filtering** by jurisdiction or product category
- **Export capabilities** to Excel and CSV formats

### Integration Opportunities
- **Power BI connectivity** for advanced analytics
- **SharePoint integration** for team collaboration
- **API webhook support** for push notifications
- **Mobile dashboard** for executive reporting

---

*This architecture supports the business requirement to "pull the 50 most recent GTA interventions in MAST chapter D from the SGEPT REST API into a local Access database and schedule that sync automatically."* 