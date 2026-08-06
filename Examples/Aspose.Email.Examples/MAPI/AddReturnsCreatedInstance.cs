// Demonstrates that Recipients.Add and Attachments.Add hand back the object they
// just created, so it can be configured straight away instead of being fetched back
// out of the collection by index.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddReturnsCreatedInstance
    {
        public static void Run()
        {
            var message = new MapiMessage
            {
                Subject = "Invoice 2026-001",
                Body = "The invoice is attached."
            };

            // The return value replaces the old "add, then index back in" pattern:
            //   message.Recipients.Add(...);
            //   message.Recipients[message.Recipients.Count - 1].DisplayName = ...;
            var recipient = message.Recipients.Add(
                "alice.johnson@example.com", "SMTP", "Alice Johnson", MapiRecipientType.MAPI_TO);
            recipient.DisplayName = "Alice Johnson (Accounts)";

            var attachment = message.Attachments.Add("invoice.pdf", File.ReadAllBytes(Data.Email/"1.pdf"));
            attachment.DisplayName = "Invoice #2026-001.pdf";

            Console.WriteLine($"Recipient: {recipient.DisplayName} <{recipient.EmailAddress}>");
            Console.WriteLine($"Attachment: {attachment.DisplayName} ({attachment.LongFileName})");

            // Adding an already-built object still returns void - there is nothing new
            // to hand back in that case.
            message.Attachments.Add(MapiAttachment.LoadFromTnef(BuildTnef()));
            Console.WriteLine($"Attachments in total: {message.Attachments.Count}");

            var outputPath = Data.Out/"AddReturnsCreatedInstance_out.msg";
            message.Save(outputPath);
            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static string BuildTnef()
        {
            var source = MapiMessage.Load(Data.Mapi/"MsgWithAtt.msg");
            var tnefPath = Data.Out/"AddReturnsCreatedInstance_winmail.dat";
            source.Attachments[0].SaveToTnef(tnefPath);
            return tnefPath;
        }
    }
}
