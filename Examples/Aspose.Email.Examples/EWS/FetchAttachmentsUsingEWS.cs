// Demonstrates downloading a single attachment without the rest of the message.
//
// Listing with ExchangeListMessagesOptions.FetchAttachmentInformation fills
// ExchangeMessageInfo.Attachments with each attachment's name, size and URI.
// FetchAttachment then downloads just that attachment - useful when a message carries
// several large files and only one is needed.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchAttachmentsUsingEWS
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("EwsAttachments");

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var infos = client.ListMessages(client.MailboxInfo.InboxUri, 50,
                    ExchangeListMessagesOptions.FetchAttachmentInformation);

                var message = infos.FirstOrDefault(info => info.Attachments.Count > 0);
                if (message == null)
                {
                    Console.WriteLine("None of the first 50 messages in the Inbox has attachments.");
                    return;
                }

                Console.WriteLine($"'{message.Subject}' has {message.Attachments.Count} attachment(s):");
                foreach (var info in message.Attachments)
                    Console.WriteLine($"  {info.Name} ({info.Size} bytes)");

                var smallest = message.Attachments.OrderBy(info => info.Size).First();
                var attachment = client.FetchAttachment(smallest.AttachmentUri);

                var outputPath = outputDir/attachment.Name;
                attachment.Save(outputPath);
                Console.WriteLine($"\nDownloaded only {attachment.Name} to {outputPath}");
            }
        }
    }
}
