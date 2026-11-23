# Export-GraphMail

PowerShell script that archives Microsoft 365 / Exchange Online email as `.eml`
files using Microsoft Graph. Messages are marked with the Outlook category
`Filed`; messages with that category older than the configured cutoff are moved
to the mailbox's Deleted Items folder.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7 or later.
- A signed-in Microsoft 365 user who can access each target mailbox.
- The Microsoft Graph PowerShell modules `Microsoft.Graph.Authentication` and
	`Microsoft.Graph.Mail`:

	```powershell
	Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Mail -Scope CurrentUser
	```

- Microsoft Graph delegated permission `Mail.ReadWrite`. With `-SharedMailbox`,
	the script also requests `Mail.ReadWrite.Shared`. The signed-in user must have
	appropriate mailbox access; tenant consent policies may require an
	administrator to approve these permissions.
- A local destination with enough free space for the exported messages.

The first run opens the Microsoft Graph sign-in flow. The script connects for
the current PowerShell process and does not save credentials.

## Usage

Run the script from PowerShell, supplying one or more exact mail folder names,
the destination path, and one or more mailbox IDs (user principal names or
Microsoft Entra object IDs):

```powershell
./Export-GraphMail.ps1 `
	-FolderNames 'Inbox,Projects' `
	-DestinationPath 'D:\MailArchive' `
	-MailboxIds 'alex@contoso.com','support@contoso.com'
```

Folder names are comma-separated and matched exactly against root-level
folders. Subfolders are not searched. For example, `Inbox/Projects` is not a
valid nested-folder path. A matching folder name is processed in every
specified mailbox where it exists.

To archive a shared mailbox, include its mailbox ID and add
`-SharedMailbox`:

```powershell
./Export-GraphMail.ps1 `
	-FolderNames 'Inbox' `
	-DestinationPath 'D:\MailArchive' `
	-MailboxIds 'shared@contoso.com' `
	-SharedMailbox
```

The default cutoff is one year before the run date. Set it explicitly with
`-OlderThanDate` when you want a different retention boundary:

```powershell
./Export-GraphMail.ps1 `
	-FolderNames 'Inbox' `
	-DestinationPath 'D:\MailArchive' `
	-MailboxIds 'alex@contoso.com' `
	-OlderThanDate (Get-Date '2024-01-01')
```

Preview planned exports and moves before running against live mail:

```powershell
./Export-GraphMail.ps1 `
	-FolderNames 'Inbox' `
	-DestinationPath 'D:\MailArchive' `
	-MailboxIds 'alex@contoso.com' `
	-WhatIf -Verbose
```

`-WhatIf` still signs in and reads mailbox data, but does not create the
destination directory, download messages, change categories, or move messages.
Use `-Confirm` to request confirmation for each proposed change.

## What the Script Does

For each matching folder, the script performs two passes:

1. Finds messages without the `Filed` category, downloads their MIME content as
	 `.eml`, and adds the `Filed` category after a successful download.
2. Finds messages sent before `OlderThanDate` and moves those already marked
	 `Filed` to Deleted Items. This includes old messages archived on the first
	 pass of the same run.

Exports are organized by the message's sent date using the machine's local
time zone:

```text
<DestinationPath>/<year>/<month>/<timestamp> <subject>.eml
```

The script sanitizes characters that are invalid in filenames. It does not
overwrite an existing file; a name collision is reported as a processing
error. Temporary downloads are removed when an export fails.

## Operational Notes

- Moving to Deleted Items is not permanent deletion. Mailbox retention policies
	may later remove those messages. Review your organization's retention rules
	before using this as an archival workflow.
- `Filed` is the script's marker that a message has been exported. Messages
	already marked `Filed` are eligible for the move pass even if this run did
	not export them. Keep the local archive protected and backed up before
	enabling cleanup.
- An export is written before the category is updated. If the category update
	fails, the `.eml` may exist while the message remains unfiled; a later run
	will refuse to overwrite a file with the same generated name. Check the
	reported output path and mailbox category before retrying or manually
	resolving that message.
- The destination can contain sensitive email. Apply appropriate access
	controls, encryption, and backup/retention practices.
- Do not run multiple instances against the same mail folders and destination
	at the same time.
- If Microsoft Graph rejects its server-side category filter, the script falls
	back to reading the folder and verifies categories locally. This may take
	longer for large folders.
- The script prints a final summary and throws an error at the end if any
	individual messages failed. Review the error details and summary even when
	the command reports failure; other messages may have completed successfully.

## License

See [LICENSE](LICENSE).
