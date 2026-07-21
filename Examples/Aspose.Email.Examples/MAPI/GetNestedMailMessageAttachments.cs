// Demonstrates how to extract a message embedded in an MSG file and save it as EML.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetNestedMailMessageAttachments
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"WithEmbeddedMsg.msg");
            var saved = 0;

            foreach (var attachment in message.Attachments)
            {
                // Only an embedded message exposes ObjectData; ordinary file attachments
                // carry their bytes in BinaryData instead and are skipped here.
                if (attachment.ObjectData == null || !attachment.ObjectData.IsOutlookMessage)
                {
                    Console.WriteLine($"Skipping file attachment: {attachment.DisplayName}");
                    continue;
                }

                var embedded = MapiMessage.FromProperties(attachment.ObjectData.Properties);
                var mailMessage = embedded.ToMailMessage(new MailConversionOptions());

                var outputPath = Data.Out/"NestedMailMessageAttachments_out.eml";
                mailMessage.Save(outputPath, SaveOptions.DefaultEml);
                saved++;

                Console.WriteLine($"Extracted embedded message \"{embedded.Subject}\" to {outputPath}");
            }

            Console.WriteLine($"\nExtracted {saved} embedded message(s).");
        }
    }
}
