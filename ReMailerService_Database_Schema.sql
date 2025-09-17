-- =============================================
-- ReMailerService Database Schema
-- SQL Server DDL Script
-- =============================================

-- Create database (uncomment if needed)
-- CREATE DATABASE [scanner];
-- GO

USE [scanner];
GO

-- =============================================
-- Table: MESSAGES
-- Main table for storing email messages to be sent
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MESSAGES]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[MESSAGES](
        [MS_IDENT] [int] IDENTITY(1,1) NOT NULL,
        [SEND_TO] [nvarchar](255) NULL,
        [SEND_FROM] [nvarchar](255) NULL,
        [SUBJECT] [nvarchar](500) NULL,
        [BODY] [ntext] NULL,
        [SENT_FLG] [char](1) NULL DEFAULT ('N'),
        [CREATED] [datetime] NULL DEFAULT (getdate()),
        [SENT] [datetime] NULL,
        [ATTACHMENT] [nvarchar](500) NULL,
        [CC] [nvarchar](255) NULL,
        [BCC] [nvarchar](255) NULL,
        [HTML] [char](1) NULL DEFAULT ('N'),
        [TO_ID] [nvarchar](50) NULL,
        [SRC_TYPE] [nvarchar](50) NULL,
        [DEFER_UNTIL] [datetime] NULL,
        [EXPIRES] [datetime] NULL,
        [ATTACH_DOC_ID] [nvarchar](50) NULL,
        [READ_FLG] [char](1) NULL DEFAULT ('N'),
        [READ_DT] [datetime] NULL,
        [FROM_NM] [nvarchar](255) NULL,
        [FROM_ID] [nvarchar](50) NULL,
        [SRC_ID] [nvarchar](50) NULL,
        [ATTACH_TYPE] [nvarchar](20) NULL,
        CONSTRAINT [PK_MESSAGES] PRIMARY KEY CLUSTERED ([MS_IDENT] ASC)
    );
END
GO

-- =============================================
-- Table: MSGS_LOCK
-- Lock table for processing messages (prevents duplicate processing)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MSGS_LOCK]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[MSGS_LOCK](
        [MS_IDENT] [int] NOT NULL,
        [SEND_TO] [nvarchar](255) NULL,
        [SEND_FROM] [nvarchar](255) NULL,
        [SUBJECT] [nvarchar](500) NULL,
        [BODY] [ntext] NULL,
        [SENT_FLG] [char](1) NULL,
        [CREATED] [datetime] NULL,
        [SENT] [datetime] NULL,
        [ATTACHMENT] [nvarchar](500) NULL,
        [CC] [nvarchar](255) NULL,
        [BCC] [nvarchar](255) NULL,
        [HTML] [char](1) NULL,
        [TO_ID] [nvarchar](50) NULL,
        [SRC_TYPE] [nvarchar](50) NULL,
        [DEFER_UNTIL] [datetime] NULL,
        [EXPIRES] [datetime] NULL,
        [ATTACH_DOC_ID] [nvarchar](50) NULL,
        [READ_FLG] [char](1) NULL,
        [READ_DT] [datetime] NULL,
        [FROM_NM] [nvarchar](255) NULL,
        [FROM_ID] [nvarchar](50) NULL,
        [SRC_ID] [nvarchar](50) NULL,
        [MACHINE] [nvarchar](50) NULL,
        [ATTACH_TYPE] [nvarchar](20) NULL,
        [ATTEMPTS] [int] NULL DEFAULT (1),
        CONSTRAINT [PK_MSGS_LOCK] PRIMARY KEY CLUSTERED ([MS_IDENT] ASC)
    );
END
GO

-- =============================================
-- Table: MSGS_DEAD
-- Dead letter table for failed messages
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MSGS_DEAD]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[MSGS_DEAD](
        [MS_IDENT] [int] NOT NULL,
        [SEND_TO] [nvarchar](255) NULL,
        [SEND_FROM] [nvarchar](255) NULL,
        [SUBJECT] [nvarchar](500) NULL,
        [BODY] [ntext] NULL,
        [SENT_FLG] [char](1) NULL,
        [CREATED] [datetime] NULL,
        [SENT] [datetime] NULL,
        [ATTACHMENT] [nvarchar](500) NULL,
        [CC] [nvarchar](255) NULL,
        [BCC] [nvarchar](255) NULL,
        [HTML] [char](1) NULL,
        [TO_ID] [nvarchar](50) NULL,
        [SRC_TYPE] [nvarchar](50) NULL,
        [DEFER_UNTIL] [datetime] NULL,
        [EXPIRES] [datetime] NULL,
        [ATTACH_DOC_ID] [nvarchar](50) NULL,
        [READ_FLG] [char](1) NULL,
        [READ_DT] [datetime] NULL,
        [FROM_NM] [nvarchar](255) NULL,
        [FROM_ID] [nvarchar](50) NULL,
        [SRC_ID] [nvarchar](50) NULL,
        [MACHINE] [nvarchar](50) NULL,
        [ATTACH_TYPE] [nvarchar](20) NULL,
        [DEAD_DATE] [datetime] NULL DEFAULT (getdate()),
        CONSTRAINT [PK_MSGS_DEAD] PRIMARY KEY CLUSTERED ([MS_IDENT] ASC)
    );
END
GO

-- =============================================
-- Indexes for Performance
-- =============================================

-- Index on MESSAGES table for common queries
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MESSAGES]') AND name = N'IX_MESSAGES_SENT_FLG_CREATED')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MESSAGES_SENT_FLG_CREATED] ON [dbo].[MESSAGES]
    (
        [SENT_FLG] ASC,
        [CREATED] DESC
    );
END
GO

-- Index on MESSAGES table for defer processing
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MESSAGES]') AND name = N'IX_MESSAGES_DEFER_UNTIL')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MESSAGES_DEFER_UNTIL] ON [dbo].[MESSAGES]
    (
        [DEFER_UNTIL] ASC
    );
END
GO

-- Index on MSGS_LOCK table for machine-based queries
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MSGS_LOCK]') AND name = N'IX_MSGS_LOCK_MACHINE')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MSGS_LOCK_MACHINE] ON [dbo].[MSGS_LOCK]
    (
        [MACHINE] ASC
    );
END
GO

-- Index on MSGS_LOCK table for attempts tracking
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[MSGS_LOCK]') AND name = N'IX_MSGS_LOCK_ATTEMPTS')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_MSGS_LOCK_ATTEMPTS] ON [dbo].[MSGS_LOCK]
    (
        [ATTEMPTS] ASC
    );
END
GO

-- =============================================
-- Sample Data (Optional - for testing)
-- =============================================

-- Uncomment the following section to insert sample test data
/*
INSERT INTO [dbo].[MESSAGES] 
([SEND_TO], [SEND_FROM], [SUBJECT], [BODY], [SENT_FLG], [HTML], [FROM_NM])
VALUES 
('test@example.com', 'noreply@example.com', 'Test Email', 'This is a test email message.', 'N', 'N', 'Test System');

INSERT INTO [dbo].[MESSAGES] 
([SEND_TO], [SEND_FROM], [SUBJECT], [BODY], [SENT_FLG], [HTML], [FROM_NM])
VALUES 
('test2@example.com', 'noreply@example.com', 'Test HTML Email', '<html><body><h1>Test HTML Email</h1><p>This is a test HTML email message.</p></body></html>', 'N', 'Y', 'Test System');
*/

-- =============================================
-- Views for Monitoring
-- =============================================

-- View for pending messages
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vw_PendingMessages]'))
BEGIN
    EXEC dbo.sp_executesql @statement = N'
    CREATE VIEW [dbo].[vw_PendingMessages]
    AS
    SELECT 
        MS_IDENT,
        SEND_TO,
        SEND_FROM,
        SUBJECT,
        CREATED,
        DEFER_UNTIL,
        FROM_NM,
        FROM_ID
    FROM [dbo].[MESSAGES]
    WHERE SENT_FLG = ''N''
        AND (DEFER_UNTIL IS NULL OR GETDATE() > DEFER_UNTIL)
        AND BODY IS NOT NULL
        AND SEND_TO IS NOT NULL AND SEND_TO <> ''''
        AND SEND_FROM IS NOT NULL AND SEND_FROM <> ''''
    ';
END
GO

-- View for locked messages
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vw_LockedMessages]'))
BEGIN
    EXEC dbo.sp_executesql @statement = N'
    CREATE VIEW [dbo].[vw_LockedMessages]
    AS
    SELECT 
        MS_IDENT,
        SEND_TO,
        SEND_FROM,
        SUBJECT,
        CREATED,
        MACHINE,
        ATTEMPTS,
        FROM_NM,
        FROM_ID
    FROM [dbo].[MSGS_LOCK]
    ';
END
GO

-- View for dead messages
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vw_DeadMessages]'))
BEGIN
    EXEC dbo.sp_executesql @statement = N'
    CREATE VIEW [dbo].[vw_DeadMessages]
    AS
    SELECT 
        MS_IDENT,
        SEND_TO,
        SEND_FROM,
        SUBJECT,
        CREATED,
        DEAD_DATE,
        MACHINE,
        FROM_NM,
        FROM_ID
    FROM [dbo].[MSGS_DEAD]
    ';
END
GO

-- =============================================
-- Stored Procedures for Maintenance
-- =============================================

-- Procedure to clean up old processed messages
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_CleanupOldMessages]') AND type in (N'P', N'PC'))
BEGIN
    EXEC dbo.sp_executesql @statement = N'
    CREATE PROCEDURE [dbo].[sp_CleanupOldMessages]
        @DaysOld INT = 30
    AS
    BEGIN
        SET NOCOUNT ON;
        
        -- Delete old sent messages
        DELETE FROM [dbo].[MESSAGES] 
        WHERE SENT_FLG = ''Y'' 
            AND SENT < DATEADD(day, -@DaysOld, GETDATE());
        
        -- Delete old dead messages
        DELETE FROM [dbo].[MSGS_DEAD] 
        WHERE DEAD_DATE < DATEADD(day, -@DaysOld, GETDATE());
        
        SELECT @@ROWCOUNT AS ''RecordsDeleted'';
    END
    ';
END
GO

-- Procedure to get service statistics
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_GetServiceStats]') AND type in (N'P', N'PC'))
BEGIN
    EXEC dbo.sp_executesql @statement = N'
    CREATE PROCEDURE [dbo].[sp_GetServiceStats]
    AS
    BEGIN
        SET NOCOUNT ON;
        
        SELECT 
            ''Pending Messages'' AS StatType,
            COUNT(*) AS Count
        FROM [dbo].[MESSAGES]
        WHERE SENT_FLG = ''N''
            AND (DEFER_UNTIL IS NULL OR GETDATE() > DEFER_UNTIL)
            AND BODY IS NOT NULL
            AND SEND_TO IS NOT NULL AND SEND_TO <> ''''
            AND SEND_FROM IS NOT NULL AND SEND_FROM <> ''''
        
        UNION ALL
        
        SELECT 
            ''Locked Messages'' AS StatType,
            COUNT(*) AS Count
        FROM [dbo].[MSGS_LOCK]
        
        UNION ALL
        
        SELECT 
            ''Dead Messages'' AS StatType,
            COUNT(*) AS Count
        FROM [dbo].[MSGS_DEAD]
        
        UNION ALL
        
        SELECT 
            ''Sent Today'' AS StatType,
            COUNT(*) AS Count
        FROM [dbo].[MESSAGES]
        WHERE SENT_FLG = ''Y''
            AND CAST(SENT AS DATE) = CAST(GETDATE() AS DATE);
    END
    ';
END
GO

PRINT 'ReMailerService Database Schema created successfully!';
PRINT 'Tables created: MESSAGES, MSGS_LOCK, MSGS_DEAD';
PRINT 'Indexes created for performance optimization';
PRINT 'Views created for monitoring: vw_PendingMessages, vw_LockedMessages, vw_DeadMessages';
PRINT 'Stored procedures created: sp_CleanupOldMessages, sp_GetServiceStats';
