Attribute VB_Name = "modSqlSync"
Option Explicit

' ============================================================================
' modSqlSync
'
' Writes one row per new mail item to a SQL Server table via ADODB. Uses
' late binding (CreateObject) instead of a VBA project reference to the ADO
' library, so this works as-is without adding a Tools > References entry --
' any ADO version installed (which is effectively always, on Windows) is
' enough.
'
' EDIT THE CONSTANTS BELOW, and pick a connection-string option inside
' GetConnectionString(), before using this. See README.md for full setup
' steps.
' ============================================================================

Private Const SQL_SERVER As String = "YOUR_SQL_SERVER"
Private Const SQL_DATABASE As String = "YourDatabase"
Private Const SQL_TABLE As String = "dbo.MailboxMessages"

Private Const LOG_FILE_NAME As String = "OutlookToSqlSync.log" ' written to %USERPROFILE%

' ADO enum values, defined manually because late binding (CreateObject)
' doesn't expose the ADODB type library's named constants.
Private Const adVarWChar As Long = 202
Private Const adDate As Long = 7
Private Const adBoolean As Long = 11
Private Const adParamInput As Long = 1
Private Const adStateOpen As Long = 1

Public Sub SyncMailItemToSql(ByVal Item As Outlook.MailItem)
    On Error GoTo ErrHandler

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.Open GetConnectionString()

    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = _
        "IF NOT EXISTS (SELECT 1 FROM " & SQL_TABLE & " WHERE MessageId = ?) " & _
        "INSERT INTO " & SQL_TABLE & _
        " (MessageId, Subject, FromAddress, FromName, ReceivedDateTime, IsRead, HasAttachments, BodyPreview) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"

    Dim messageId As String
    messageId = Item.EntryID

    ' Parameters must be appended in the exact order the '?' placeholders
    ' appear above: the EXISTS check first, then the 8 INSERT values.
    cmd.Parameters.Append cmd.CreateParameter("p_check", adVarWChar, adParamInput, 400, messageId)
    cmd.Parameters.Append cmd.CreateParameter("p_id", adVarWChar, adParamInput, 400, messageId)
    cmd.Parameters.Append cmd.CreateParameter("p_subject", adVarWChar, adParamInput, 1000, Left$(NzStr(Item.Subject), 1000))
    cmd.Parameters.Append cmd.CreateParameter("p_from_addr", adVarWChar, adParamInput, 320, Left$(GetSenderSmtpAddress(Item), 320))
    cmd.Parameters.Append cmd.CreateParameter("p_from_name", adVarWChar, adParamInput, 200, Left$(NzStr(Item.SenderName), 200))
    cmd.Parameters.Append cmd.CreateParameter("p_received", adDate, adParamInput, , Item.ReceivedTime)
    cmd.Parameters.Append cmd.CreateParameter("p_isread", adBoolean, adParamInput, , CBool(Not Item.UnRead))
    cmd.Parameters.Append cmd.CreateParameter("p_hasattach", adBoolean, adParamInput, , CBool(Item.Attachments.Count > 0))
    cmd.Parameters.Append cmd.CreateParameter("p_preview", adVarWChar, adParamInput, 4000, Left$(NzStr(Item.Body), 4000))

    cmd.Execute

    conn.Close
    Exit Sub

ErrHandler:
    LogError "SyncMailItemToSql", Err.Number, Err.Description
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
End Sub

' --------------------------------------------------------------------------
' Pick ONE connection string. Leave exactly one of these assignments
' uncommented. See README.md for which option fits what you have installed.
' --------------------------------------------------------------------------
Private Function GetConnectionString() As String

    ' Option A (default): legacy SQLOLEDB provider. Ships in-box on every
    ' Windows install, so nothing extra to install -- deprecated by
    ' Microsoft but still functional. SQL login auth.
    GetConnectionString = "Provider=SQLOLEDB;Data Source=" & SQL_SERVER & _
        ";Initial Catalog=" & SQL_DATABASE & _
        ";User ID=YOUR_SQL_LOGIN;Password=YOUR_SQL_PASSWORD;"

    ' Option B: modern "Microsoft OLE DB Driver for SQL Server" (MSOLEDBSQL),
    ' if you've installed it: https://learn.microsoft.com/sql/connect/oledb
    ' GetConnectionString = "Provider=MSOLEDBSQL;Server=" & SQL_SERVER & _
    '     ";Database=" & SQL_DATABASE & _
    '     ";UID=YOUR_SQL_LOGIN;PWD=YOUR_SQL_PASSWORD;Encrypt=yes;TrustServerCertificate=yes;"

    ' Option C: Windows auth via SQLOLEDB -- the Windows account Outlook is
    ' running as needs SQL Server access; no SQL login/password needed.
    ' GetConnectionString = "Provider=SQLOLEDB;Data Source=" & SQL_SERVER & _
    '     ";Initial Catalog=" & SQL_DATABASE & ";Integrated Security=SSPI;"

End Function

' Outlook's Item.SenderEmailAddress returns an X.500 directory name (not a
' usable SMTP address) when the sender resolved to an Exchange/GAL contact
' rather than an external SMTP sender. Resolve through GetExchangeUser() in
' that case; fall back to SenderEmailAddress as-is otherwise.
Private Function GetSenderSmtpAddress(ByVal Item As Outlook.MailItem) As String
    On Error GoTo Fallback
    If Item.SenderEmailType = "EX" Then
        Dim exUser As Outlook.ExchangeUser
        Set exUser = Item.Sender.GetExchangeUser()
        If Not exUser Is Nothing Then
            GetSenderSmtpAddress = exUser.PrimarySmtpAddress
            Exit Function
        End If
    End If
Fallback:
    GetSenderSmtpAddress = NzStr(Item.SenderEmailAddress)
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Private Sub LogError(ByVal procName As String, ByVal errNum As Long, ByVal errDesc As String)
    On Error Resume Next
    Dim logPath As String
    logPath = Environ$("USERPROFILE") & "\" & LOG_FILE_NAME
    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Append As #fnum
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ERROR in " & procName & ": (" & errNum & ") " & errDesc
    Close #fnum
End Sub

' --------------------------------------------------------------------------
' Single source of truth for which folder this project watches. Both the
' live ItemAdd watcher (ThisOutlookSession.Application_Startup) and the
' manual backfill (SyncAllMailboxMessages below) call THIS SAME function, so
' there is exactly one place to edit -- change it here and both the live
' sync and the backfill point at the new folder together, with no way for
' them to drift out of sync with each other.
'
' Default: your own Inbox. For a shared/secondary mailbox added to this
' profile, comment the line below out and use something like:
'   Set GetWatchedFolder = ns.Folders("shared-mailbox-display-name").Folders("Inbox")
' See README.md, "Watching a different mailbox or folder", for how to find
' the exact display name.
' --------------------------------------------------------------------------
Public Function GetWatchedFolder() As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Set GetWatchedFolder = ns.GetDefaultFolder(olFolderInbox)
End Function

' --------------------------------------------------------------------------
' Manual backfill: syncs every mail item CURRENTLY sitting in the watched
' folder (GetWatchedFolder, above -- the same folder the live watcher uses),
' not just new arrivals. Safe to re-run any time -- SyncMailItemToSql
' already skips anything whose MessageId is already in the table, so this
' never creates duplicates. Useful for catching up mail that arrived before
' this macro was set up, or while Outlook was closed (this project is
' event-driven only -- see README.md's opening caveat).
'
' To turn this into an actual clickable button instead of running it from
' the VBA editor: File > Options > Quick Access Toolbar > "Choose commands
' from" > Macros > select SyncAllMailboxMessages > Add >> > OK. It then
' shows as a one-click icon in Outlook's Quick Access Toolbar.
' --------------------------------------------------------------------------
Public Sub SyncAllMailboxMessages()
    Dim targetFolder As Outlook.Folder
    Set targetFolder = GetWatchedFolder()

    Dim total As Long, synced As Long, failed As Long
    total = targetFolder.Items.Count

    Dim i As Long
    Dim itm As Object
    For i = 1 To total
        Set itm = targetFolder.Items(i)
        If TypeOf itm Is Outlook.MailItem Then
            On Error Resume Next
            Err.Clear
            SyncMailItemToSql itm
            If Err.Number <> 0 Then
                failed = failed + 1
            Else
                synced = synced + 1
            End If
            On Error GoTo 0
        End If
    Next i

    MsgBox "Backfill complete." & vbCrLf & _
           "Mail items in folder: " & total & vbCrLf & _
           "Synced OK (new or already present): " & synced & vbCrLf & _
           "Failed: " & failed & vbCrLf & vbCrLf & _
           "Check " & Environ$("USERPROFILE") & "\OutlookToSqlSync.log for failure details.", _
           vbInformation, "OutlookToSqlSync - Manual Backfill"
End Sub
