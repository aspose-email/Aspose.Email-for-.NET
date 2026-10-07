# Aspose.Email for .NET — Examples

Runnable C# examples for [Aspose.Email for .NET](https://products.aspose.com/email/net).
Each example is a small, self-contained program that does one thing and prints what it did.

## Quick start

```bash
git clone <this-repository>
cd Examples
dotnet run --project Aspose.Email.Examples
```

That opens an interactive menu: filter by keyword or category, pick a number, see the output.

To run one example directly, pass its name:

```bash
dotnet run --project Aspose.Email.Examples -- CreateNewEmail
```

Requires the [.NET SDK](https://dotnet.microsoft.com/download) 8.0 or later. The project also
targets .NET Framework 4.8, which is what the Windows Forms and Gmail examples need.

## Licensing

Without a license Aspose.Email runs in evaluation mode: output is watermarked and truncated.
To apply your license, do either of:

- drop your `.lic` file into this folder (next to the `.sln`), or
- set the `ASPOSE_EMAIL_LICENSE` environment variable to its full path.

The environment variable wins if both are present. Get a free 30-day temporary license from the
[Aspose site](https://purchase.aspose.com/temporary-license/) if you do not have one.

> Keep your license file out of version control — `.gitignore` already excludes `*.lic`.

## What is where

File formats and storages:

| Folder | What it covers |
|---|---|
| [`Email`](Aspose.Email.Examples/Email) | MIME messages: create, load, convert, attachments, headers, MHTML/HTML, iCalendar, TNEF |
| [`MAPI`](Aspose.Email.Examples/MAPI) | Outlook items and PST storages: messages, contacts, tasks, notes, appointments, recurrences |
| [`PST`](Aspose.Email.Examples/PST) | Personal storage files: open from a file or stream, enumerate folders, read asynchronously, convert OST → PST |
| [`OLM`](Aspose.Email.Examples/OLM) | Outlook for Mac storages: find folders, filter and extract messages, handle damaged files |
| [`MBOX`](Aspose.Email.Examples/MBOX) | Thunderbird / mbox storages: read sequentially, asynchronously or by page, filter, write |

Mail servers and services:

| Folder | What it covers |
|---|---|
| [`Graph`](Aspose.Email.Examples/Graph) | Microsoft 365 through Microsoft Graph: folders, messages, attachments, sending, calendars, contacts, To Do tasks, categories, inbox rules, OData queries, paging, throttling, asynchronous client |
| [`EWS`](Aspose.Email.Examples/EWS) | Exchange Web Services |
| [`IMAP`](Aspose.Email.Examples/IMAP) | IMAP client: connection and security, folders, messages and attachments, search, threading and sorting, CONDSTORE sync, quotas, folder monitoring (IDLE), backup and restore, task-based API |
| [`POP3`](Aspose.Email.Examples/POP3) | POP3 client |
| [`SMTP`](Aspose.Email.Examples/SMTP) | SMTP client: plain-text, HTML and alternate-view messages, bulk, multi-connection and queued sending, forwarding, mail merge, meeting requests, S/MIME, TNEF, pickup directory, delivery notifications, error handling, OAuth, TLS and proxies, task-based API |
| [`Gmail`](Aspose.Email.Examples/Gmail) | Gmail via Google API (.NET Framework 4.8 only) |

Other:

| Folder | What it covers |
|---|---|
| [`Licensing`](Aspose.Email.Examples/Licensing) | Metered licensing |
| [`Tools`](Aspose.Email.Examples/Tools) | Utilities used to regenerate the sample data — not examples |

Supporting folders:

- `Data/` — input files the examples read. Treat as read-only: examples never modify them.
- `Out/` — everything the examples write. Emptied at the start of every run.

## Which examples run offline

`Email`, `MAPI`, `PST`, `OLM`, `MBOX` and `Licensing` work straight away against the files in
`Data/` — no server, no account, nothing to configure.

`Graph`, `EWS`, `IMAP`, `POP3`, `SMTP` and `Gmail` talk to a real mail server or service and need
an account — see [Connecting to a mail server](#connecting-to-a-mail-server).

Some examples do useful work either way: they build their result offline, save it to `Out/`, and
only send it when SMTP is configured. That covers a few examples in the offline folders and,
in `SMTP`, mail merge, meeting requests, S/MIME signing and the iCalendar, EML and pickup-directory
examples.

## Connecting to a mail server

Most server examples take their connection details from
[`clientsettings.json`](Aspose.Email.Examples/clientsettings.json), one section per service:

| Section | Used by | How it signs in | What to fill in |
|---|---|---|---|
| `Graph` | `Graph` examples | OAuth 2.0 with application permissions (client secret) | `TenantId`, `ClientId`, `ClientSecret`, `MailboxId` |
| `Ews` | `EWS` examples | OAuth 2.0 with application permissions (client secret) | `TenantId`, `ClientId`, `ClientSecret`, `UserName` |
| `Imap` | `IMAP` examples | OAuth 2.0 with delegated permissions — you sign in in the browser | `HostName`, `Port`, `UserName`, `TenantId`, `ClientId` |
| `Smtp` | `SMTP` examples, and offline examples that can mail their result | user name and password | `HostName`, `Port`, `UserName`, `Password` |

The OAuth sections assume an application registered in
[Microsoft Entra ID](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app):

- **Graph** needs application permissions with admin consent for what the examples touch —
  `Mail.ReadWrite`, `Mail.Send`, `Contacts.ReadWrite`, `Calendars.ReadWrite`, `Tasks.ReadWrite`.
  With application permissions there is no signed-in user, so `MailboxId` names the mailbox to
  work on (its user principal name). `EndPoint` only needs changing for a national cloud.
  Until the section is filled in, the Graph examples print what is missing instead of failing.
- **IMAP** needs the delegated `IMAP.AccessAsUser.All` permission and `http://localhost` as a
  redirect URI for a public client.

For **SMTP**, `UserName` must be an e-mail address: the `SMTP` examples send their test messages
to it, so nothing leaves your own mailbox. Port 587 means STARTTLS, port 465 implicit TLS. Until
`HostName` is filled in, the `SMTP` examples print what is missing instead of failing.

Examples that change the mailbox — append, move, flag, delete — work in a temporary folder with a
unique name (`Aspose-<guid>`) and delete it at the end, so existing mail is only read.

Some examples keep values in the code on purpose, because those values are what the example is
about: OAuth tokens and Entra ID app ids, NTLM, certificate validation and allowed sign-in
mechanisms, and the proxy addresses in the proxy examples. Older examples, including all of
`POP3` and `Gmail`, still carry their connection details as placeholders in the code (`<HOST>`,
`username`, `password` and the like). Replace such values before running these examples.

## Writing your own

Copy any example as a starting point. The conventions are:

- One class per file, with a `public static void Run()` — that is all the runner needs to find it.
- Input paths come from `Data.<Category>`, output paths from `Data.Out`, for example:

  ```csharp
  var msg = MapiMessage.Load(Data.Mapi/"message.msg");
  msg.Save(Data.Out/"MyExample_out.msg");
  ```

- Server clients come from `ClientBuilder`, which reads `clientsettings.json`:

  ```csharp
  using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
  {
      client.SelectFolder(ImapFolderInfo.InBox);
  }
  ```

  `ClientBuilder.Smtp`, `ClientBuilder.Graph` / `ClientBuilder.GraphAsync` and `ClientBuilder.Ews`
  work the same way. Check `ClientBuilder.IsSmtpConfigured` or `ClientBuilder.IsGraphConfigured`
  first, so that the example explains what is missing when it runs without a server.
- Asynchronous examples keep the synchronous `Run()` and call
  `RunAsync().GetAwaiter().GetResult()` from it.
- The class name is how the example is invoked, so keep it descriptive.

## Support

- [Documentation](https://docs.aspose.com/email/net/)
- [API reference](https://reference.aspose.com/email/net/)
- [Free support forum](https://forum.aspose.com/c/email)
- [Paid support helpdesk](https://helpdesk.aspose.com/)
