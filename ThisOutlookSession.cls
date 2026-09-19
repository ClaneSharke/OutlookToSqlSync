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
' Watches the Inbox for new mail and syncs each one to SQL Server via
' modSqlSync.SyncMailItemToSql. This only runs -- and only catches mail --
' while Outlook is open and this profile is signed into the mailbox you
' want captured. See README.md for the "is this the right tool for this"
' caveat before relying on it.
'
' This file is exported/importable as-is via VBE's File > Import File, but
' the usual path is simpler: open the real ThisOutlookSession module inside
' Outlook's VBA editor (Alt+F11) and paste the code below (from "Private
' WithEvents" down) into it directly -- see README.md step 4.
' ============================================================================

Private WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    ' Watches the default Inbox of this Outlook profile. To watch a
    ' different folder -- e.g. a shared/secondary mailbox added to this
    ' profile, or a subfolder rather than the Inbox itself -- change this
    ' line. See README.md for an example.
    Set InboxItems = ns.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    ' The Inbox can receive non-mail items too (meeting requests are
    ' MailItem-derived and are included on purpose; skip anything else).
    If TypeOf Item Is Outlook.MailItem Then
        modSqlSync.SyncMailItemToSql Item
    End If
End Sub
