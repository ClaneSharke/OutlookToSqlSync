/* ============================================================================
   CreateTable.sql
   Destination table for the OutlookToSqlSync VBA macro. Same shape as the
   MailboxReaderClr / EwsToSqlSync projects' tables on purpose, so all three
   capture methods can point at compatible-shaped data if you ever want to
   mix and match. MessageId (Outlook's EntryID) is the primary key -- the
   macro checks it before inserting, and this constraint is a second line of
   defense against duplicate rows.
   ============================================================================ */

CREATE TABLE dbo.MailboxMessages (
    MessageId         NVARCHAR(400)  NOT NULL PRIMARY KEY,
    Subject           NVARCHAR(1000) NULL,
    FromAddress       NVARCHAR(320)  NULL,
    FromName          NVARCHAR(200)  NULL,
    ReceivedDateTime  DATETIME2      NULL,
    IsRead            BIT            NULL,
    HasAttachments    BIT            NULL,
    BodyPreview       NVARCHAR(4000) NULL,
    LoadedAtUtc       DATETIME2      NOT NULL DEFAULT SYSUTCDATETIME()
);
GO

-- Grant the account Outlook/this macro connects as just what it needs
-- (adjust the login name, or skip if using Windows/Integrated auth as the
-- signed-in user):
-- GRANT SELECT, INSERT ON dbo.MailboxMessages TO your_sql_login;
