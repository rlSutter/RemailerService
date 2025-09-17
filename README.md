# ReMailerService

A Windows service application that processes email messages from a SQL Server database and sends them via a web service. The service provides reliable email delivery with retry mechanisms, dead letter handling, and comprehensive logging.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Database Schema](#database-schema)
- [Usage](#usage)
- [Monitoring](#monitoring)
- [Troubleshooting](#troubleshooting)
- [Development](#development)
- [License](#license)

## Overview

ReMailerService is a Windows service built in VB.NET that acts as a reliable email processing system. It continuously monitors a SQL Server database for pending email messages, processes them through a web service, and handles delivery failures with retry logic and dead letter queuing.

### Key Components

- **Windows Service**: Runs continuously in the background
- **Database Integration**: Connects to SQL Server for message queuing
- **Web Service Client**: Sends emails via external web service
- **Locking Mechanism**: Prevents duplicate message processing
- **Retry Logic**: Handles temporary failures with configurable retry attempts
- **Dead Letter Queue**: Stores permanently failed messages
- **Comprehensive Logging**: Uses log4net for detailed logging

## Features

- **Reliable Message Processing**: Uses database locking to prevent duplicate processing
- **Configurable Retry Logic**: Automatically retries failed messages up to 3 times
- **Dead Letter Handling**: Moves permanently failed messages to a separate table
- **Multi-format Support**: Handles both plain text and HTML emails
- **Attachment Support**: Processes email attachments with document management integration
- **Deferred Sending**: Supports scheduled email delivery
- **Machine-specific Processing**: Supports multiple service instances
- **Comprehensive Logging**: File-based and syslog logging with configurable levels
- **Registry Configuration**: Stores configuration in Windows Registry
- **Nagios Integration**: Provides monitoring output for system monitoring

## System Requirements

### Software Requirements
- Windows Server 2008 R2 or later / Windows 7 or later
- .NET Framework 4.0 or later
- SQL Server 2008 or later
- IIS (for web service endpoint)

### Hardware Requirements
- Minimum 512 MB RAM
- 100 MB free disk space
- Network connectivity to database and web service

### Dependencies
- log4net 2.0.3
- System.Data.SqlClient
- System.ServiceProcess
- System.Web

## Installation

### 1. Database Setup

First, create the required database schema:

```sql
-- Run the provided SQL script
sqlcmd -S [SERVER_NAME] -d [DATABASE_NAME] -i ReMailerService_Database_Schema.sql
```

### 2. Service Installation

1. **Build the Service**:
   ```cmd
   msbuild ReMailerService.sln /p:Configuration=Release
   ```

2. **Install the Service**:
   ```cmd
   InstallUtil.exe ReMailerService.exe
   ```

3. **Alternative Installation** (using provided batch file):
   ```cmd
   "Install Remailer Service 2.bat"
   ```

### 3. Service Configuration

The service will automatically create default registry entries on first run. Configure the service using the Windows Registry:

**Registry Path**: `HKEY_LOCAL_MACHINE\SOFTWARE\ReMailerService`

| Value Name | Default | Description |
|------------|---------|-------------|
| MyInterval | 60 | Timer interval in seconds |
| UserName | SCANNER | Database username |
| Password | SCANNER | Database password |
| DBName | scanner | Database name |
| DBServer | (empty) | Database server name |
| Debug | Y | Enable debug logging |
| Logging | Y | Enable general logging |

### 4. Application Configuration

Edit `app.config` to configure:

```xml
<appSettings>
    <add key="cloudsvc_farm" value="[PROXY_SERVER_IP]"/>
</appSettings>
```

## Configuration

### Registry Configuration

The service reads configuration from the Windows Registry. Use the following registry path:

```
HKEY_LOCAL_MACHINE\SOFTWARE\ReMailerService
```

#### Configuration Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| MyInterval | String | "60" | Timer interval in seconds between processing cycles |
| UserName | String | "SCANNER" | SQL Server username for database connection |
| Password | String | "SCANNER" | SQL Server password for database connection |
| DBName | String | "scanner" | Name of the database containing message tables |
| DBServer | String | "" | SQL Server instance name or IP address |
| Debug | String | "Y" | Enable debug logging ("Y" or "N") |
| Logging | String | "Y" | Enable general logging ("Y" or "N") |

### Application Configuration

The `app.config` file contains additional settings:

```xml
<configuration>
    <appSettings>
        <!-- Proxy server for web service calls -->
        <add key="cloudsvc_farm" value="192.168.7.21"/>
    </appSettings>
    
    <!-- Web service endpoint configuration -->
    <applicationSettings>
        <ReMailerService.My.MySettings>
            <setting name="ReMailerService_com_certegrity_cloudsvc_Service"
                serializeAs="String">
                <value>http://cloudsvc.certegrity.com/basic/service.asmx</value>
            </setting>
        </ReMailerService.My.MySettings>
    </applicationSettings>
</configuration>
```

### Logging Configuration

The service uses log4net for logging with the following configuration:

- **Remote Syslog**: Sends logs to syslog server (192.168.5.37)
- **File Logging**: Writes to `C:\Logs\ReMailerService.log`
- **Log Rotation**: 2MB file size with 3 backup files
- **Log Levels**: DEBUG, INFO, WARN, ERROR

## Database Schema

The service uses three main tables:

### MESSAGES Table
Primary table storing email messages to be sent.

| Column | Type | Description |
|--------|------|-------------|
| MS_IDENT | int (PK) | Message identifier (auto-increment) |
| SEND_TO | nvarchar(255) | Recipient email address |
| SEND_FROM | nvarchar(255) | Sender email address |
| SUBJECT | nvarchar(500) | Email subject |
| BODY | ntext | Email body content |
| SENT_FLG | char(1) | Send status ('N'=Not sent, 'Y'=Sent, 'E'=Error) |
| CREATED | datetime | Message creation timestamp |
| SENT | datetime | Message sent timestamp |
| ATTACHMENT | nvarchar(500) | Attachment file path |
| CC | nvarchar(255) | CC recipients |
| BCC | nvarchar(255) | BCC recipients |
| HTML | char(1) | HTML format flag ('Y' or 'N') |
| TO_ID | nvarchar(50) | Recipient ID |
| SRC_TYPE | nvarchar(50) | Source system type |
| DEFER_UNTIL | datetime | Deferred send time |
| EXPIRES | datetime | Message expiration time |
| ATTACH_DOC_ID | nvarchar(50) | Document management system ID |
| READ_FLG | char(1) | Read status flag |
| READ_DT | datetime | Read timestamp |
| FROM_NM | nvarchar(255) | Sender display name |
| FROM_ID | nvarchar(50) | Sender ID |
| SRC_ID | nvarchar(50) | Source system ID |
| ATTACH_TYPE | nvarchar(20) | Attachment type ('dms', 'reports', etc.) |

### MSGS_LOCK Table
Temporary table for processing messages (prevents duplicate processing).

Contains all MESSAGES columns plus:
- MACHINE: nvarchar(50) - Processing machine name
- ATTEMPTS: int - Number of processing attempts

### MSGS_DEAD Table
Dead letter table for permanently failed messages.

Contains all MESSAGES columns plus:
- MACHINE: nvarchar(50) - Processing machine name
- DEAD_DATE: datetime - Date moved to dead letter queue

## Usage

### Starting the Service

```cmd
net start ReMailerService
```

### Stopping the Service

```cmd
net stop ReMailerService
```

### Service Status

```cmd
sc query ReMailerService
```

### Adding Messages to Queue

Insert messages into the MESSAGES table:

```sql
INSERT INTO scanner.dbo.MESSAGES 
(SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, HTML, FROM_NM)
VALUES 
('recipient@example.com', 'sender@example.com', 'Test Email', 
 'This is a test email message.', 'N', 'N', 'Test System');
```

### Processing Flow

1. **Timer Fires**: Service checks for new messages every configured interval
2. **Lock Messages**: Selects pending messages and locks them for processing
3. **Process Messages**: Sends each message via web service
4. **Update Status**: Marks successful messages as sent, retries failed ones
5. **Dead Letter**: Moves permanently failed messages to dead letter queue
6. **Cleanup**: Removes processed messages from lock table

## Monitoring

### Log Files

- **Service Log**: `C:\Logs\ReMailerService.log`
- **Nagios Output**: `C:\ReMailerService.nagios`
- **Windows Event Log**: Application log with source "ReMailerService"

### Database Views

Use the provided views for monitoring:

```sql
-- View pending messages
SELECT * FROM vw_PendingMessages;

-- View locked messages
SELECT * FROM vw_LockedMessages;

-- View dead messages
SELECT * FROM vw_DeadMessages;

-- Get service statistics
EXEC sp_GetServiceStats;
```

### Nagios Integration

The service creates a Nagios-compatible status file at `C:\ReMailerService.nagios` containing:
- Success message with timestamp
- Error message with details if failures occur

### Performance Monitoring

Monitor these key metrics:
- Message processing rate
- Failed message count
- Dead letter queue size
- Service uptime
- Database connection health

## Troubleshooting

### Common Issues

#### Service Won't Start
1. Check Windows Event Log for error details
2. Verify database connection settings in registry
3. Ensure SQL Server is accessible
4. Check file permissions for log directory

#### Messages Not Processing
1. Verify messages exist in MESSAGES table with SENT_FLG='N'
2. Check DEFER_UNTIL field (messages may be deferred)
3. Ensure BODY, SEND_TO, and SEND_FROM are not null/empty
4. Check for messages locked by other service instances

#### Web Service Errors
1. Verify web service endpoint is accessible
2. Check proxy server configuration
3. Review network connectivity
4. Check web service logs

#### Database Connection Issues
1. Verify SQL Server is running
2. Check connection string parameters
3. Ensure database user has proper permissions
4. Test connection with SQL Server Management Studio

### Debug Mode

Enable debug logging by setting registry value:
```
HKEY_LOCAL_MACHINE\SOFTWARE\ReMailerService\Debug = "Y"
```

This will provide detailed logging including:
- SQL queries executed
- Web service requests/responses
- Processing details for each message
- Error details and stack traces

### Log Analysis

Key log entries to monitor:
- `Service Starting/Stopping`: Service lifecycle events
- `Message ID: X sent to Y. Sent: Success/Error`: Message processing results
- `Unable to open database connection`: Database connectivity issues
- `Error executing SendMail web service`: Web service failures

## Development

### Project Structure

```
ReMailerService/
├── ReMailerService.vb          # Main service class
├── CRegistry.vb                # Registry access utilities
├── ProjectInstaller.vb         # Service installer
├── AssemblyInfo.vb             # Assembly information
├── app.config                  # Application configuration
├── ReMailerService.vbproj      # Project file
├── packages.config             # NuGet packages
└── ServiceTestQueries.sql      # Test SQL queries
```

### Building from Source

1. **Prerequisites**:
   - Visual Studio 2010 or later
   - .NET Framework 4.0 SDK

2. **Build**:
   ```cmd
   msbuild ReMailerService.sln /p:Configuration=Release
   ```

3. **Dependencies**:
   - log4net 2.0.3 (included in packages/)
   - System references (included with .NET Framework)

### Code Architecture

- **ServiceBase**: Inherits from System.ServiceProcess.ServiceBase
- **Timer-based Processing**: Uses System.Timers.Timer for periodic execution
- **Database Access**: Uses System.Data.SqlClient for SQL Server connectivity
- **Web Service**: Uses System.Net.HttpWebRequest for HTTP calls
- **Logging**: Uses log4net for comprehensive logging
- **Configuration**: Uses Windows Registry for persistent configuration

### Testing

Use the provided test queries in `ServiceTestQueries.sql`:

```sql
-- Test message insertion
INSERT scanner.dbo.MESSAGES (SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG)
VALUES ('test@example.com', 'noreply@example.com', 'Test', 'Test message', 'N');

-- Check processing status
SELECT * FROM vw_PendingMessages;
SELECT * FROM vw_LockedMessages;
```

## License

This project is proprietary software. All rights reserved.

## Support

For technical support or questions:
1. Check the troubleshooting section above
2. Review log files for error details
3. Contact the development team with specific error messages and log excerpts

---

**Version**: 1.1.0.0  
**Last Updated**: [Current Date]  
**Author**: Development Team
