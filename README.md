# .NET Email API

[Aspose.Email for .NET](https://products.aspose.com/email/net) allows you to work with MIME messages, appointments, Microsoft Outlook® items, Outlook storage files, various clients & protocols ([SMTP](https://docs.aspose.com/email/net/connecting-to-smtp-server/), [POP3](https://docs.aspose.com/email/net/connect-to-pop3-server/), [IMAP](https://docs.aspose.com/email/net/connecting-to-imap-server/), [Exchange EWS](https://docs.aspose.com/email/net/connecting-to-exchange-server/), [Exchange WebDav](https://docs.aspose.com/email/net/connecting-to-exchange-server-using-webdav/), [Gmail](https://docs.aspose.com/email/net/gmail-utility-features/), [Thunderbird](https://docs.aspose.com/email/net/programming-with-thunderbird/), [Zimbra](https://docs.aspose.com/email/net/working-with-zimbra/), [IBM Notes](https://docs.aspose.com/email/net/working-with-ibm-notes/) & AMP HTML emails and more.

<p align="center">
<a title="Download complete Aspose.Email for .NET source code" href="https://github.com/aspose-email/Aspose.Email-for-.NET/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

Directory | Description
--------- | -----------
[Demos](Demos)  | Source code for the live demos hosted at https://products.aspose.app/email/family.
[Examples](Examples)  | A collection of .NET examples that help you learn the product features.
[Plugins](Plugins)  | Visual Studio Plugins related to Aspose.Email for .NET.


## Email Creation, Conversion & Transactional API

- Open or save emails in Microsoft Outlook & other formats.
- Conversion of email files to various formats.
- Parse, read & save MS Outlook emails, PST & OST files.
- [Comprehensive support for MIME messages](https://docs.aspose.com/email/net/creating-and-setting-contents-of-emails/).
- Send & receive emails via POP3, IMAP, Microsoft Exchange Server.
- Send emails in bulk while performing mail merge via various types of data sources.
- Send iCalendar compliant messages.
- Consume and produce recurrence patterns in the iCalendar (RFC 2445) format.
- Tools to verify email addresses, email syntax, email domain, mail server & MX records.
- Extract objects from various mail storage formats as well as [create email storage files from scratch](https://docs.aspose.com/email/net/create-new-pst-file-and-add-subfolders/).

## Read & Write Email Formats

**Microsoft Outlook:** MSG, PST, OST, OFT\
**Email:** EML, EMLX, MBOX\
**Others:** ICS, VCF, HTML, MHTML

## Read Email Formats

**Mac Outlook:** OLM

## Platform Independence

Aspose.Email for .NET is implemented using Managed C# and can be used with any .NET based languages including C#, VB.NET, J# and so on. Aspose.Email for .NET can be integrated with any kind of application on Windows, Linux, Mac OS X and Xamarin Platforms.

## Get Started with Aspose.Email for .NET

Are you ready to give Aspose.Email for .NET a try? Simply execute `Install-Package Aspose.Email` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Aspose.Email for .NET and want to upgrade the version, please execute `Update-Package Aspose.Email` to get the latest version.

## Create an Email Message in MSG Format

```csharp
// create an instance of the MailMessage class
var message = new MailMessage();
// set from, to, subject and body properties
message.From = "sender@domain.com";
message.To = "receiver@domain.com";
message.Subject = "This is test message";
mamessageilMsg.Body = "This is test body";
// create an instance of the MapiMessage class and pass object of MailMessage as argument
var msg = MapiMessage.FromMailMessage(message);
// save file on disc
msg.Save(dir + "output.msg");
```

## Send Bulk Emails via SMTP

```csharp
// create SmtpClient as client and specify server, port, user name and password
SmtpClient client = new SmtpClient("mail.server.com", 25, "Username", "Password");

// create instances of MailMessage class and Specify To, From, Subject and Message
MailMessage message1 = new MailMessage("msg1@from.com", "msg1@to.com", "Subject1", "message1, how are you?");
MailMessage message2 = new MailMessage("msg1@from.com", "msg2@to.com", "Subject2", "message2, how are you?");
MailMessage message3 = new MailMessage("msg1@from.com", "msg3@to.com", "Subject3", "message3, how are you?");

// create an instance of MailMessageCollection class
MailMessageCollection manyMsg = new MailMessageCollection();
manyMsg.Add(message1);
manyMsg.Add(message2);
manyMsg.Add(message3);

// send emails in bulk
client.Send(manyMsg);
```

## Create an ICS Appointment via .NET

```csharp
// create and initialize an instance of the Appointment class
var appointment = new Appointment(
	"Meeting Room 3 at Office Headquarters",// Location
	"Monthly Meeting",                      // Summary
	"Please confirm your availability.",    // Description
	new DateTime(2015, 2, 8, 13, 0, 0),     // Start date
	new DateTime(2015, 2, 8, 14, 0, 0),     // End date
	"from@domain.com",                      // Organizer
	"attendees@domain.com");                // Attendees

// save the appointment to disk in ICS format
appointment.Save("output.ics", AppointmentSaveFormat.Ics);
```

[Home](https://www.aspose.com/) | [Product Page](https://products.aspose.com/email/net) | [Docs](https://docs.aspose.com/email/net/) | [Demos](https://products.aspose.app/email/family) | [API Reference](https://apireference.aspose.com/email/net) | [Examples](https://github.com/aspose-email/Aspose.Email-for-.NET) | [Blog](https://blog.aspose.com/category/email/) | [Free Support](https://forum.aspose.com/c/email) | [Temporary License](https://purchase.aspose.com/temporary-license)
