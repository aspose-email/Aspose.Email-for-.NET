// Demonstrates working with the attachments of a message that lives on the server:
// adding one, listing them, downloading one and removing it.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphAttachments
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            var outputDir = Data.OutSub("GraphAttachments");
            var from = ClientBuilder.GraphMailboxId;
            const string subject = "Graph attachment example";

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                var workFolder = client.CreateFolder("Aspose.Email attachments");

                var message = new MapiMessage(from, "recipient@example.com", subject,
                    "This message gets an attachment added on the server.", OutlookMessageFormat.Unicode);
                client.CreateMessage(workFolder.ItemId, message);

                var info = client.ListMessages(workFolder.ItemId, null)
                    .Cast<MessageInfo>()
                    .FirstOrDefault(m => m.Subject == subject);

                if (info == null)
                {
                    Console.WriteLine("The message was not found after creating it.");
                    return;
                }

                // MapiAttachment has no public constructor: one is made by adding it to
                // a message, and since 26.7 Add hands the created instance straight back.
                var carrier = new MapiMessage();
                var attachment = carrier.Attachments.Add("1.txt", File.ReadAllBytes(Data.Mapi/"1.txt"));

                // Attach it to the message that is already on the server.
                var created = client.CreateAttachment(info.ItemId, attachment);
                Console.WriteLine($"Attached {created.LongFileName}");

                var attachments = client.ListAttachments(info.ItemId, null);
                Console.WriteLine($"{attachments.Count} attachment(s) on the message:");

                foreach (var item in attachments)
                {
                    var size = item.BinaryData == null ? 0 : item.BinaryData.Length;
                    Console.WriteLine($"  {item.LongFileName} ({size:N0} bytes), id: {item.ItemId}");
                }

                // The server-side id lives on ItemId, and that is what the fetch and
                // delete calls take - not the file name.
                var downloaded = client.FetchAttachment(created.ItemId);
                if (downloaded != null && downloaded.BinaryData != null)
                {
                    var outputPath = outputDir/downloaded.LongFileName;
                    File.WriteAllBytes(outputPath, downloaded.BinaryData);
                    Console.WriteLine($"Downloaded to {outputPath}");
                }

                client.DeleteAttachment(created.ItemId);
                Console.WriteLine("Removed the attachment.");

                client.Delete(workFolder.ItemId);
                Console.WriteLine($"Deleted {workFolder.DisplayName}");
            }
        }
    }
}
