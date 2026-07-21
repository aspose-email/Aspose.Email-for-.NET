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

| Folder | What it covers |
|---|---|
| [`Email`](Aspose.Email.Examples/Email) | MIME messages: create, load, convert, attachments, headers, MHTML/HTML, iCalendar, TNEF |
| [`MAPI`](Aspose.Email.Examples/MAPI) | Outlook items and PST storages: messages, contacts, tasks, notes, appointments, recurrences |
| [`PST`](Aspose.Email.Examples/PST) | Personal storage files: open, inspect, convert OST → PST |
| [`OLM`](Aspose.Email.Examples/OLM) | Outlook for Mac storages |
| [`MBOX`](Aspose.Email.Examples/MBOX) | Thunderbird / mbox storages |
| [`EWS`](Aspose.Email.Examples/EWS) | Exchange Web Services |
| [`IMAP`](Aspose.Email.Examples/IMAP), [`POP3`](Aspose.Email.Examples/POP3), [`SMTP`](Aspose.Email.Examples/SMTP) | Mail protocol clients |
| [`Gmail`](Aspose.Email.Examples/Gmail) | Gmail via Google API |
| [`Licensing`](Aspose.Email.Examples/Licensing) | Metered licensing |
| [`Tools`](Aspose.Email.Examples/Tools) | Utilities used to regenerate the sample data — not examples |

Supporting folders:

- `Data/` — input files the examples read. Treat as read-only: examples never modify them.
- `Out/` — everything the examples write. Emptied at the start of every run.

## Which examples run offline

`Email`, `MAPI`, `PST`, `OLM`, `MBOX` and `Licensing` work straight away against the files in
`Data/` — no server, no account, nothing to configure.

`EWS`, `IMAP`, `POP3`, `SMTP` and `Gmail` talk to a real mail server. Put your server details in
[`clientsettings.json`](Aspose.Email.Examples/clientsettings.json) before running them. A few
examples in the offline folders can also send their result by mail; they save it to `Out/` either
way and only send when SMTP is configured.

## Writing your own

Copy any example as a starting point. The conventions are:

- One class per file, with a `public static void Run()` — that is all the runner needs to find it.
- Input paths come from `Data.<Category>`, output paths from `Data.Out`, for example:

  ```csharp
  var msg = MapiMessage.Load(Data.Mapi/"message.msg");
  msg.Save(Data.Out/"MyExample_out.msg");
  ```

- The class name is how the example is invoked, so keep it descriptive.

## Support

- [Documentation](https://docs.aspose.com/email/net/)
- [API reference](https://reference.aspose.com/email/net/)
- [Free support forum](https://forum.aspose.com/c/email)
- [Paid support helpdesk](https://helpdesk.aspose.com/)
