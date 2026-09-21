VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================================================
' ThisOutlookSession
'
' Watches a mailbox folder for new mail and syncs each one to SQL Server via
' SyncMailItemToSql (in modSqlSync.bas). This only runs -- and only catches
' mail -- while Outlook is open and this profile is signed into the mailbox
' you want captured. See README.md for the "is this the right tool for this"
' caveat before relying on it.
'
' The call below is unqualified (SyncMailItemToSql, not
' modSqlSync.SyncMailItemToSql) on purpose: pasting modSqlSync.bas's
' contents into a module via Insert > Module does NOT rename that module to
' "modSqlSync" just because the pasted text has an "Attribute VB_Name" line
' -- that line only takes effect via VBE's File > Import File. If you leave
' the module as the default "Module1" (very common), a qualified
' modSqlSync.SyncMailItemToSql call fails to compile. Calling it unqualified
' works regardless of what the module ends up named, as long as there's
' exactly one Public Sub with that name in the project (there is).
'
' This file is exported/importable as-is via VBE's File > Import File, but
' the usual path is simpler: open the real ThisOutlookSession module inside
' Outlook's VBA editor (Alt+F11) and paste the code below (from "Private
' WithEvents" down) into it directly -- see README.md step 4.
' ============================================================================

Private WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    ' GetWatchedFolder (in modSqlSync.bas) is the single source of truth for
    ' which folder this project watches -- the manual backfill
    ' (SyncAllMailboxMessages) reads the same function, so the two can never
    ' point at different folders. To watch a different folder -- e.g. a
    ' shared/secondary mailbox added to this profile, or a subfolder rather
    ' than the Inbox itself -- edit GetWatchedFolder, not this line. See
    ' README.md, "Watching a different mailbox or folder".
    Set InboxItems = GetWatchedFolder().Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    ' The Inbox can receive non-mail items too (meeting requests are
    ' MailItem-derived and are included on purpose; skip anything else).
    If TypeOf Item Is Outlook.MailItem Then
        SyncMailItemToSql Item
    End If
End Sub
