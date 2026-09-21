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

By default this watches the **default Inbox** of the signed-in profile.
Both the live watcher (`Application_Startup` in `ThisOutlookSession.cls`)
and the manual backfill (`SyncAllMailboxMessages`, below) get their folder
from the same place -- `GetWatchedFolder()` in `modSqlSync.bas` -- so there
is exactly one line to change, and the two can never end up pointed at
different folders. Edit it there:

```vb
Public Function GetWatchedFolder() As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Set GetWatchedFolder = ns.GetDefaultFolder(olFolderInbox)
End Function
```

For a shared/secondary mailbox added to the profile, replace the last line
with something like:

```vb
Set GetWatchedFolder = ns.Folders("shared-mailbox-display-name").Folders("Inbox")
```

The exact display name is whatever that mailbox shows as in Outlook's
folder pane -- if you're not sure, run this in the VBA Immediate Window
(Alt+F11, then Ctrl+G) to list every store and its top-level folders:

```vb
Sub ListAllFolders()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Dim st As Outlook.Store
    Dim rootFolder As Outlook.Folder
    Dim topFolder As Outlook.Folder
    Dim subFolder As Outlook.Folder

    For Each st In ns.Stores
        Debug.Print "STORE: " & st.DisplayName
        On Error Resume Next
        Set rootFolder = Nothing
        Set rootFolder = st.GetRootFolder
        On Error GoTo 0

        If Not rootFolder Is Nothing Then
            On Error Resume Next
            For Each topFolder In rootFolder.Folders
                Debug.Print "  " & topFolder.Name
                For Each subFolder In topFolder.Folders
                    Debug.Print "    " & subFolder.Name
                Next subFolder
            Next topFolder
            On Error GoTo 0
        End If
    Next st
End Sub
```

## Running it manually on demand (backfill)

`modSqlSync.bas` includes `SyncAllMailboxMessages`, a separate `Sub` that
loops over every item currently sitting in the watched folder and syncs each
one -- not just new arrivals. It's safe to run any time: `SyncMailItemToSql`
already skips anything whose `MessageId` is already in the table, so
re-running it never creates duplicates. Useful for catching up mail that
arrived before this macro was set up, or while Outlook was closed (remember,
this project is event-driven only -- see the caveat at the top of this
README).

It targets whatever `GetWatchedFolder()` (in `modSqlSync.bas`) returns --
the same function `Application_Startup` uses for the live watcher -- so
there's no separate folder setting to keep in sync; change
`GetWatchedFolder()` once and both the live sync and this backfill follow.

To run it as a one-click button instead of from the VBA editor: **File >
Options > Quick Access Toolbar**, set "Choose commands from" to **Macros**,
select `SyncAllMailboxMessages`, click **Add >>**, then **OK**. It now shows
as an icon in Outlook's Quick Access Toolbar -- one click runs the backfill
and shows a summary (items found / synced / failed) when it's done. The same
"Macros" source is available under **File > Options > Customize Ribbon** if
you'd rather have a labeled ribbon button instead of a small toolbar icon.

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
- **It's writing your own mail instead of the mailbox you actually care
  about** -- `GetWatchedFolder()` is returning your own default Inbox
  (the out-of-the-box default), not the shared/secondary mailbox you meant.
  See "Watching a different mailbox or folder" above -- you need to point it
  at that mailbox's folder explicitly; it's never inferred automatically.
- **`ListAllFolders` (or a similar recursive folder-walk macro) throws an
  error partway through** -- a hidden system folder, search folder, or a
  shared-mailbox subfolder you don't have full rights to can error out
  mid-loop. Wrap folder access in `On Error Resume Next` / `On Error GoTo 0`
  per folder (as the `ListAllFolders` version in this README already does)
  rather than letting one bad folder abort the whole listing.

## License

MIT -- see `LICENSE`.
