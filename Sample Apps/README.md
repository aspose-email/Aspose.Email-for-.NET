# Aspose.Email for .NET — Sample Apps

Simple apps that demonstrate the capabilities of [Aspose.Email for .NET](https://products.aspose.com/email/net).
Each app is a standalone console project with its own solution — open the `.sln`, or run it from the
command line.

Looking for short, single-feature snippets instead? See the [Examples](../Examples) project.

## Quick start

```bash
cd "Sample Apps/MailboxExtractor"
dotnet run --project MailboxExtractor -- "mailbox.pst" "out"
```

Requires the [.NET SDK](https://dotnet.microsoft.com/download). Target frameworks range from .NET 5.0
to .NET 8.0 depending on the app; see each project file.

## What is where

| App | What it does |
|---|---|
| [`ConversationThread`](ConversationThread) | Groups messages from a PST by conversation thread and saves each thread to its own folder. |
| [`EWSModernAuthenticationApp`](EWSModernAuthenticationApp) | Makes a modern-authenticated EWS request using application-only authentication. |
| [`EWSModernAuthenticationDelegated`](EWSModernAuthenticationDelegated) | Makes a modern-authenticated EWS request using delegated authentication. |
| [`EWSModernAuthenticationImapSmtp`](EWSModernAuthenticationImapSmtp) | Makes modern-authenticated IMAP and SMTP requests using delegated authentication. |
| [`GraphApp`](GraphApp) | Reads the folder hierarchy and Inbox messages of a mailbox through the Microsoft Graph API. |
| [`MailboxExtractor`](MailboxExtractor) | Extracts messages from PST/OST, MBOX, OLM and TGZ storages into a directory. |
| [`PstPasswordManager`](PstPasswordManager) | Validates, adds, changes and removes the password of a PST file. |
| [`ReplaceAttach`](ReplaceAttach) | Replaces the attachments of a `MapiMessage` by index. |

## Which apps run offline

`ConversationThread`, `MailboxExtractor`, `PstPasswordManager` and `ReplaceAttach` work on local
files — point them at your own storage file and run.

`EWSModernAuthenticationApp`, `EWSModernAuthenticationDelegated`, `EWSModernAuthenticationImapSmtp`
and `GraphApp` talk to Microsoft 365. They need an application registered in Azure AD; put its
tenant, client and user details in the `appsettings.json` of the app before running it.

## Licensing

Without a license Aspose.Email runs in evaluation mode: output is watermarked and truncated, and
some operations are capped (for example `ReplaceAttach` handles no more than three attachments).
Each app sets the license at start-up — replace the placeholder path in its `Program.cs` with the
path to your `.lic` file.

Get a free 30-day temporary license from the
[Aspose site](https://purchase.aspose.com/temporary-license/) if you do not have one.

> Keep your license file out of version control.

## Support

- [Documentation](https://docs.aspose.com/email/net/)
- [API reference](https://reference.aspose.com/email/net/)
- [Free support forum](https://forum.aspose.com/c/email)
- [Paid support helpdesk](https://helpdesk.aspose.com/)
