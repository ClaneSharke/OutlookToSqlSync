# OutlookToSqlSync

An Outlook VBA macro that writes each new mail item straight into a SQL
Server table the moment it arrives -- no server-side component, no
scheduled task, nothing installed on SQL Server itself.

## Read this before you set it up

This captures mail **only while Outlook is open, running, and signed into
the mailbox you want captured**, on whatever Windows machine you set it up
on. If that machine reboots, Outlook isn't running, nobody's signed in, or
the mailbox is one nobody actively keeps an open desktop session for, this
silently stops capturing -- no error, no alert, it just misses mail until
Outlook is next open and watching.

It's also **event-driven, not a poll/backfill**: it only reacts to mail
that arrives while it's actively watching. It does not go back and pick up
messages that arrived while Outlook was closed, the way a scheduled
poll-based sync would.

If you need mail capture that works reliably and unattended regardless of
whether any desktop client is open -- which matters if this is a dedicated
service-account mailbox rather than something on your own daily-use desktop
-- use a server-side approach instead:
[MailboxReaderClr](https://github.com/ClaneSharke/MailboxReaderClr) (a SQL
Server CLR proc, or its standalone `EwsToSqlSync` companion tool) polls
independently of any desktop client being open at all.

This project is the right fit when you're already signed into the mailbox
in Outlook most of the day anyway, or for a quick prototype -- not for
unattended capture on a mailbox nobody's actively logged into.

## Requirements

- **Classic desktop Outlook** (the Win32 client with VBA support). The
  newer "new Outlook" (the Monarch-based client Microsoft has been rolling
  out) has no VBA support at all -- this won't work there.
- Windows.
- A SQL Server table to write into (see `CreateTable.sql`).

## Setup

### 1. Create the destination table

Run `CreateTable.sql` against your SQL Server database.

### 2. Allow macros to run

File > Options > Trust Center > Trust Center Settings > Macro Settings.
Set it to **"Notifications for all macros"** (or lower) -- if it's set to
disable all macros, the code below will never run, silently.

### 3. Open the VBA editor

Alt+F11 (or Developer tab > Visual Basic -- enable the Developer tab first
via File > Options > Customize Ribbon if you don't see it).

### 4. Add the code

In the Project Explorer, expand **Project1 (VbaProject.OTM) > Microsoft
Outlook Objects**, then double-click **ThisOutlookSession**. Paste in
everything from `ThisOutlookSession.cls` in this repo, starting at
`Option Explicit` (skip the `VERSION 1.0 CLASS` / `Attribute` header lines
at the top of the file -- those are only needed if you use VBE's File >
Import File instead of copy-paste, and ThisOutlookSession already exists as
a module, so importing isn't the normal path here).

Then: right-click **Project1 (VbaProject.OTM) > Insert > Module**, and
paste the entire contents of `modSqlSync.bas` into the new module.

### 5. Configure it

At the top of `modSqlSync.bas` (now pasted into your VBA project), edit:

```vb
Private Const SQL_SERVER As String = "YOUR_SQL_SERVER"
Private Const SQL_DATABASE As String = "YourDatabase"
Private Const SQL_TABLE As String = "dbo.MailboxMessages"
```

Then in `GetConnectionString()`, pick the option that matches what you
have: **Option A** (default, uncommented) uses the legacy `SQLOLEDB`
provider, which ships in-box on every Windows install -- no extra driver
needed, just edit in your SQL login and password. Options B (modern
`MSOLEDBSQL` driver, if installed) and C (Windows/Integrated auth instead
of a SQL login) are included as comments -- swap one in if it fits your
environment better.

### 6. Save and restart Outlook

Ctrl+S to save the VBA project (stored in `VbaProject.OTM`), then fully
close and reopen Outlook so `Application_Startup` runs and starts watching.

### 7. Test it

Send yourself a test email. Check `dbo.MailboxMessages` for a new row. If
nothing shows up, check `%USERPROFILE%\OutlookToSqlSync.log` -- errors
(bad connection string, permission denied, etc.) are logged there since
there's no console to print to.

## Watching a different mailbox or folder

By default this watches the **default Inbox** of the signed-in profile. To
watch a different folder -- a shared/secondary mailbox already added to
this Outlook profile, or a subfolder rather than the Inbox -- change this
line in `Application_Startup`:

```vb
Set InboxItems = ns.GetDefaultFolder(olFolderInbox).Items
```

For a secondary mailbox added to the profile, something like:

```vb
Dim sharedInbox As Outlook.Folder
Set sharedInbox = ns.Folders("shared-mailbox-display-name").Folders("Inbox")
Set InboxItems = sharedInbox.Items
```

## Notes on the data captured

- **`MessageId`** is Outlook's own `EntryID` for the item. It's stable
  while the item stays in the same folder/store, but Outlook's own
  documentation notes an `EntryID` *can* change if an item is later moved
  between folders or stores -- fine for this use case (capturing on
  arrival, not tracking items as they get filed later), just worth knowing.
- **`FromAddress`** resolves correctly for both external SMTP senders and
  internal Exchange/GAL senders (the latter would otherwise show an
  unreadable X.500 directory name instead of an email address -- handled
  via `GetExchangeUser().PrimarySmtpAddress`).
- **`BodyPreview`** is the first 4000 characters of the plain-text body.

## Troubleshooting

- **Nothing gets synced, no log file appears** -- macros likely aren't
  enabled at all (step 2), or `Application_Startup` never ran because
  Outlook wasn't restarted after pasting the code (step 6).
- **Log file shows a connection error** -- check the provider/driver you
  picked in `GetConnectionString()` is actually available on this machine,
  and that the server name, database, and credentials are correct. Test
  the same connection string format from a small standalone script if you
  want to isolate whether it's a VBA problem or a SQL Server / network
  problem.
- **Log file shows a permission error** -- the SQL login (or Windows
  account, if using Option C) needs `INSERT`/`SELECT` on the destination
  table.

## License

MIT -- see `LICENSE`.
