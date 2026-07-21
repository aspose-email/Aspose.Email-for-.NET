// Demonstrates how to convert an EML file to MSG format while keeping an embedded
// Outlook message embedded, instead of flattening it into a plain attachment.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class PreservingEmbeddedMsgFormat
    {
        public static void Run()
        {
            // sample.eml is TNEF-encoded and carries a real embedded Outlook message.
            // PreserveEmbeddedMessageFormat keeps that item in its original MSG format
            // rather than re-encoding it as rfc822 while loading.
            var eml = MailMessage.Load(Data.Email/"sample.eml",
                new EmlLoadOptions { PreserveEmbeddedMessageFormat = true });

            var options = new MapiConversionOptions
            {
                Format = OutlookMessageFormat.Unicode,
                PreserveEmbeddedMessageFormat = true
            };

            var msg = MapiMessage.FromMailMessage(eml, options);

            var outputPath = Data.Out/"PreservingEmbeddedMsgFormat_out.msg";
            msg.Save(outputPath);

            Console.WriteLine($"Converted: {eml.Subject}");
            Console.WriteLine($"Saved to:  {outputPath}\n");

            // Read the result back to show the attachment is still an embedded message.
            var saved = MapiMessage.Load(outputPath);
            foreach (var attachment in saved.Attachments)
            {
                var embedded = attachment.ObjectData != null && attachment.ObjectData.IsOutlookMessage
                    ? MapiMessage.FromProperties(attachment.ObjectData.Properties)
                    : null;

                Console.WriteLine(embedded != null
                    ? $"Attachment '{attachment.DisplayName}' is an embedded Outlook message: \"{embedded.Subject}\""
                    : $"Attachment '{attachment.DisplayName}' is a plain file attachment.");
            }
        }
    }
}
