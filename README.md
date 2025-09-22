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

## Overview

ReMailerService is a Windows service built in VB.NET that acts as a reliable email processing system. It continuously monitors a SQL Server database for pending email messages, processes them through a web service, and handles delivery failures with retry logic and dead letter queuing.

### System Architecture

This service is part of a three-component application architecture designed for scalability and reliability:

1. **ReMailerService**: A Windows service application that runs on designated servers and monitors the message queue, wraps messages in XML documents, and invokes SendMail using HTTP POST with SOAP calls.

2. **SendMail**: A web service that runs on the cloud service cluster. It accepts connections, receives messages in XML format, transmits them via SMTP to the configured SMTP server, and updates the message queue when directed.

3. **ServicesController**: A Windows form-based application that runs in the taskbar and allows the ReMailerService to be stopped and started, with default value configuration capabilities.

The purpose of this architecture is to provide scalability and improve reliability by decoupling queue management from message transmission functions, allowing multiple transmission services to handle the queue simultaneously.

### Key Components

- **Windows Service**: Runs continuously in the background
- **Database Integration**: Connects to SQL Server for message queuing
- **Web Service Client**: Sends emails via external web service
- **Locking Mechanism**: Prevents duplicate message processing using elastic record locking
- **Retry Logic**: Handles temporary failures with configurable retry attempts
- **Dead Letter Queue**: Stores permanently failed messages
- **Comprehensive Logging**: Uses log4net for detailed logging
- **Multi-instance Support**: Supports multiple service instances for load distribution

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

## Deployment

### Deployment Process

The application is deployed by copying the contents of the build output folder to the installation folder on target servers. The deployment includes the following files and directories:

#### Required Files for Deployment:
- `Installservice.bat` - Service installation script
- `ReMailerService.application` - Application manifest
- `ReMailerService.exe` - Main executable
- `ReMailerService.exe.config` - Application configuration
- `ReMailerService.exe.manifest` - Application manifest
- `ReMailerService.pdb` - Debug symbols
- `ReMailerService.vshost.application` - Visual Studio host application
- `ReMailerService.vshost.exe` - Visual Studio host executable
- `ReMailerService.vshost.exe.config` - Host configuration
- `ReMailerService.vshost.exe.manifest` - Host manifest
- `ReMailerService.xml` - Configuration file
- `ServiceInstaller.exe` - Service installer utility
- `XMLDB.dll` - Database XML library

#### Directories:
- `ReMailerService.publish/` - Published application files
- `Xml/` - XML configuration files

### Installation Steps

1. **Stop the Service**: Turn off the ReMailerService Windows service from the Services control panel
2. **Close Services Applet**: Close the Services control panel
3. **Run Installation Script**: Execute `C:\attachments\ReMailerService\Installservice.bat`
4. **Verify Installation**: Open the Services applet and check that the ReMailerService Startup Type is not marked "disabled"
5. **Retry if Needed**: If installation was unsuccessful, wait about a minute and rerun the installation script
6. **Confirm Status**: The service should show "Started" status when installation is complete

### Multi-Server Deployment

The service can be installed on multiple servers by:
1. Creating the folder `C:\attachments` on the target server
2. Copying the installation folder to that server
3. Executing the installation script on each server

Each server will have its own registry configuration and can process messages independently using the locking mechanism.

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

### Service Execution

The service is controlled through the Windows Services applet where it appears as "ReMailerService". It can be started and stopped using the control buttons or from the command line:

#### Starting the Service
```cmd
net start ReMailerService
```

#### Stopping the Service
```cmd
net stop ReMailerService
```

#### Service Status
```cmd
sc query ReMailerService
```

### Execution Parameters

The service execution parameters are stored as Registry entries in the key `HKEY_LOCAL_MACHINE\SOFTWARE\ReMailerService`:

| Value | Purpose |
|-------|---------|
| DBName | The name of the database used for the connection |
| DBServer | The database server that contains the MESSAGES table |
| Debug | "Y" to turn on debug logging mode |
| Logging | "Y" to turn on operational logging mode |
| MyInterval | How frequently, in seconds, the service checks the MESSAGES queue |
| Password | The database server password |
| UserName | The database server username |

**Note**: The URL of the SendMail web service cannot be changed without modifying the program. The current endpoint is configured in the application settings.

### Service Operation

When the service is executed, it:
1. Starts a timer that checks the queue with the specified frequency
2. Sets a flag to prevent concurrent queue checks during processing
3. Creates log files and Windows Event entries when messages are processed (depending on logging configuration)
4. Processes messages in batches based on the interval size (formula: interval size × 2)

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

## Scheduling

This service is not scheduled in the traditional sense. Instead, it runs as a background service that continuously checks the queue at the interval specified in the registry configuration. The service:

- Runs continuously in the background
- Checks the MESSAGES queue at the configured interval (default: 60 seconds)
- Processes messages in batches (batch size = interval × 2)
- Uses elastic record locking to prevent duplicate processing across multiple service instances
- Automatically handles queue management without external scheduling dependencies

## Monitoring

### Log Files

- **Service Log**: `C:\Logs\ReMailerService.log`
- **Nagios Output**: `C:\ReMailerService.nagios`
- **Windows Event Log**: Application log with source "ReMailerService"

### Database Monitoring Queries

#### View Current Queue Status
```sql
-- View pending messages (assuming 10-second interval)
SELECT TOP 20 SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, 
ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES, 
ATTACH_DOC_ID, READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT 
FROM scanner.dbo.MESSAGES 
WHERE SENT_FLG='N' 
AND (DEFER_UNTIL IS NULL OR GETDATE()>DEFER_UNTIL) 
AND BODY IS NOT NULL 
AND SEND_TO IS NOT NULL AND SEND_TO<>''
AND SEND_FROM IS NOT NULL AND SEND_FROM<>'' 
ORDER BY CREATED DESC
```

#### Count Unsent Messages
```sql
SELECT COUNT(*)
FROM scanner.dbo.MESSAGES 
WHERE SENT_FLG='N' 
AND (DEFER_UNTIL IS NULL OR GETDATE()>DEFER_UNTIL) 
AND BODY IS NOT NULL 
AND SEND_TO IS NOT NULL AND SEND_TO<>''
AND SEND_FROM IS NOT NULL AND SEND_FROM<>''
```

#### View Locked Messages
```sql
SELECT * FROM scanner.dbo.MSGS_LOCK WHERE MACHINE='[MACHINE_NAME]'
```

#### View Dead Messages
```sql
SELECT * FROM scanner.dbo.MSGS_DEAD ORDER BY CREATED DESC
```

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

Every time the program is executed, it generates the file `ReMailerService.nagios` with either "Success" or "Failure", followed by the execution date and time. The exact text format is:

```
Success on 10/17/2007 12:46:40 PM
```

When reporting a failure, the exact error appears after the date/time of execution. This file is located at `C:\ReMailerService.nagios` and is used for system monitoring integration.

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

## Operational Notes

### Email Processing Flow

When someone sends an email message through the web site, the following process occurs:

1. **Message Insertion**: A query is used to put the message into the reports.MESSAGES table
2. **Queue Processing**: ReMailerService checks the queue of messages and pulls out a batch to be sent
3. **Message Locking**: Messages are locked in the queue and sent one at a time to the SendMail web service
4. **Web Service Processing**: Messages are directed by SendMail to the SMTP server, which uses SMTP to send them to the destination server

### Troubleshooting Email Delivery Issues

If email isn't going through, check the following in order:

1. **Check MESSAGES Table**: Verify if it's being blocked by a malformed email
2. **Check Service Logs**: Review `C:\ReMailService.log` on service servers - has anything been added recently? If no, restart the service
3. **Check Web Service Logs**: Review `C:\inetpub\basic\SendMail.log` on cloudsvc servers - has anything been added recently? If no, restart the web service
4. **Check SMTP Server**: Verify if the SMTP server is down

### Locked Message Issues

Sometimes messages will be "trapped" in the MSGS_LOCK table and not processed by any instance of the service. This can occur when a malformed message is "clogging" the queue. To resolve:

#### Identify the Problem Message
```sql
SELECT TOP 20 MS_IDENT, SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT,  
ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES,  ATTACH_DOC_ID, 
READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT,'SIEBEL'  
FROM scanner.dbo.MESSAGES M WHERE SENT_FLG='N'  AND 
(DEFER_UNTIL IS NULL OR GETDATE()>DEFER_UNTIL)  AND BODY IS NOT NULL  AND 
SEND_TO IS NOT NULL AND SEND_TO<>'' AND SEND_FROM IS NOT NULL AND SEND_FROM<>''  
AND NOT EXISTS 
 ( SELECT MS_IDENT FROM scanner.dbo.MSGS_LOCK WHERE MS_IDENT=M.MS_IDENT ) 
ORDER BY CREATED DESC
```

#### Remove Malformed Message
```sql
UPDATE scanner.dbo.[MESSAGES]
SET SENT_FLG='E'
WHERE MS_IDENT='[MS_IDENT FROM ABOVE QUERY]'
```

#### Clear the Lock Table
```sql
TRUNCATE TABLE scanner.dbo.MSGS_LOCK
```

### Service Configuration Notes

- The service creates registry keys in `[HKEY_LOCAL_MACHINE][SOFTWARE][ReMailerService]` with default values on first run
- Default values include database connection settings, logging configuration, and processing intervals
- The web service URL is hardcoded in the application and cannot be changed without code modification
- Multiple service instances can run simultaneously using the elastic record locking mechanism

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

The application can generate two different types of logs, both stored in `C:\ReMailerService.log`:

#### Debug Logging
Used to monitor the internals of the application for debugging queries and transmission issues. Example debug log:

```
ReMailerService Trace Log Started 10/1/2019 2:44:35 PM
PARAMETERS
  Debug: Y
  Logging: Y
  cloudsvc_farm: 192.168.7.21
  MyInterval: 60
  UserName: SCANNER
  PassWord: SCANNER
  DBName: scanner
  DBServer: <database server>

Opened connection to con 
Opened connection to ucon

Running on DATAFLUXAPP1

QUERY: Remove extraneous lock table entries: 
DELETE FROM scanner.dbo.MSGS_LOCK WHERE MACHINE='<server>'

Checking for 60 new messages at 10/1/2019 2:45:34 PM on <server>

QUERY: Get count of email messages in lock table: 
SELECT COUNT(*) FROM scanner.dbo.MSGS_LOCK WHERE MACHINE='<server>'

   ...Lock table message count: 0

Message # 1 Started
  > mSEND_FROM: user@example.com
  > mSEND_TO: recipient@example.com
  > mSUBJECT: Test Email
  > mTO_ID: NULL
  > mATTACH_DOC_ID: NULL
  > mATTACH_TYPE: NULL

Msg #: 1  Sending: http://<server>/basic/service.asmx/SendMail?sXML=...
  > results <?xml version="1.0" encoding="utf-8"?>
<boolean xmlns="http://<server>/basic/">true</boolean>
Message #  1 ID: 35209772  Sent: Success
```

#### Operational Logging
Simpler logging that reports each email sent. Both types of logs also generate entries in the Windows Event log, and basic transactions and critical errors are logged to SysLog.

#### Key Log Entries to Monitor:
- `Service Starting/Stopping`: Service lifecycle events
- `Message ID: X sent to Y. Sent: Success/Error`: Message processing results
- `Unable to open database connection`: Database connectivity issues
- `Error executing SendMail web service`: Web service failures
- `ReMailerService : Service Starting/Stopping`: SysLog entries

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

### Key Programming Components

#### Main Service Class (ReMailerService.vb)
The main service class contains:
- Timer management for periodic queue checking
- Database connection handling with connection pooling
- Message processing logic with retry mechanisms
- Web service communication via HTTP POST
- Comprehensive error handling and logging

#### Registry Management (CRegistry.vb)
Custom class for reading and writing Windows Registry values, overcoming restrictions of GetSetting/SaveSetting which only allow access to HKEY_CURRENT_USER\Software\VB and VBA.

#### HTTP Proxy Class
Custom `simplehttp` class for reliable HTTP communication with proxy support:
- GET and POST request handling
- Proxy server configuration
- Timeout and error handling
- Cookie management

#### Database Functions
- `OpenDBConnection()`: Opens database connections with extreme error handling
- `CloseDBConnection()`: Safely closes database connections and clears connection pools
- Connection pooling management to prevent resource leaks

#### Utility Functions
- `CleanString()`: Removes extraneous spaces and bad characters from email addresses
- `EmailAddressCheck()`: Validates email address format using regex
- `ChkString()`: Creates SQL-safe strings for INSERT statements
- `ClearString()`: Removes CRLFs from strings

### Testing

Use the provided test queries in `ServiceTestQueries.sql`:

```sql
-- Test message insertion
INSERT <database>.dbo.MESSAGES (SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG)
VALUES ('test@example.com', 'noreply@example.com', 'Test', 'Test message', 'N');

-- Check processing status
SELECT * FROM vw_PendingMessages;
SELECT * FROM vw_LockedMessages;
```

## Update History

### Version History and Major Updates

#### 10/1/2019 - Load Balancer Migration Support
- Updated to support LoadBalancer migration by putting the cloudsvc farm address into the "cloudsvc_farm" app setting
- Fixed logging output formatting

#### 10/4/2016 - Attachment Support Improvements
- Updated to improve logging functionality
- Fixed support for attachments to support "reports" instead of "report"

#### 9/30/2016 - Attachment Type Support
- Modified to support the ATTACH_TYPE field per the adjusted data model in SendMail
- Enhanced attachment handling capabilities

#### 10/5/2015 - Version Reporting and Connection Management
- Updated to report the version number in the event log when starting
- Implemented standard function for closing database connections
- Added automatic close and reopen of database connections if the application can't reach the table

#### 3/6/2014 - HTTP Reliability Improvements
- Updated to reduce routine logging when no messages are found
- Implemented HTTP proxy class instead of calling web services through normal calls
- Improved HTTP reliability and error handling

#### 1/13/2014 - Logging Integration
- Updated to integrate with SysLog
- Implemented log4net to manage local debug logs
- Updated to .NET 4 framework

#### 7/14/2011 - Error Handling Fixes
- Fixed inaccessible code for updating the MSGS_DEAD table
- Corrected attempt counting for failures
- Added CleanString function to remove extraneous spaces and bad characters from email addresses

#### 7/12/2011 - Queue Management Improvements
- Updated to check the queue for existing records before populating it
- Resolved queue management problem by checking for existing records before loading more
- Implemented attempt tracking and removal only after 3 retry attempts
- Previously marked messages as "Error" status after one attempt only

#### 11/29/2010 - Web Service Migration
- Modified to change web reference from com.gettips.siebel to com.certegrity.cloudsvc
- Part of migration from siebel.hq.local to cloudsvc.certegrity.com

#### 10/30/2008 - Database Connection Optimization
- Moved database connection opening and closing to OnStart and OnStop events
- Reduced database connections by opening once per service start/stop
- Changed default configuration to use SCANNER database login instead of "sa"
- Application now empties work queue (MSGS_LOCK) on startup
- Improved debug logic with actual Windows error messages

#### 10/14/2008 - Record Locking Implementation
- Adapted service to perform record locking for multiple instance support
- Created MSGS_LOCK table for temporary message processing
- Implemented MSGS_DEAD table for permanently failed messages
- Fixed issues with messages that couldn't be sent due to incorrect body information
- Improved communication speed between Windows service and web service

#### 1/21/2008 - Queue Processing Optimization
- Modified service to sort messages by date created descending
- Most recently created records are now addressed first
- Changed queue slice to twice the check interval (instead of same as interval)
- Improved reliability and prevented system resource overuse
- Reduced network service demand and bandwidth usage


