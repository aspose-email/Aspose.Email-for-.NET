// Demonstrates how to load an OFT template, fill in sender/recipient/body placeholders, and save as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadAndWritingOutlookTemplateFile
    {
        public static void Run()
        {
            // An OFT is loaded with the MSG options - the two share the same format.
            var message = MailMessage.Load(Data.Mapi/"sample.oft", new MsgLoadOptions());

            const string senderDisplayName = "John";
            const string senderEmailAddress = "john@abc.com";
            const string recipientDisplayName = "William";
            const string recipientEmailAddress = "william@xzy.com";

            message.Sender = new MailAddress(senderEmailAddress, senderDisplayName);
            message.To.Add(new MailAddress(recipientEmailAddress, recipientDisplayName));

            // The template's body holds plain placeholder words; filling the template in
            // is a matter of replacing them.
            message.HtmlBody = message.HtmlBody.Replace("DisplayName", $"<b>{recipientDisplayName}</b>");
            message.HtmlBody = message.HtmlBody.Replace("MeetingPlace", "<u>Hall 1, Convention Center, New York, USA</u>");
            message.HtmlBody = message.HtmlBody.Replace("MeetingTime", "<u>Monday, June 28, 2010</u>");

            var msg = MapiMessage.FromMailMessage(message);

            // Save as a draft, so Outlook opens the filled-in template for editing.
            msg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT);

            var outputPath = Data.Out/"ReadAndWritingOutlookTemplateFile_out.msg";
            msg.Save(outputPath);

            Console.WriteLine($"Template: {message.Subject}");
            Console.WriteLine($"Filled in for {recipientDisplayName} <{recipientEmailAddress}>");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
