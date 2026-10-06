// Demonstrates how to see what a message has attached and download a single
// attachment, leaving the rest of the message on the server.
//
// ListAttachments reads only the message structure, so it stays cheap even for huge
// messages; FetchAttachment then pulls one attachment by name.

using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListAndFetchAttachments
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var outputDir = Data.OutSub("ImapAttachments");

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    // A test message with a small text file and a large binary file.
                    var message = new MailMessage("from@example.com", "to@example.com",
                        "Site survey", "Notes and photos from the visit.");
                    message.Attachments.Add(new Attachment(
                        new MemoryStream(Encoding.UTF8.GetBytes("Roof needs repair.")), "notes.txt", "text/plain"));
                    message.Attachments.Add(new Attachment(
                        new MemoryStream(new byte[1024 * 1024]), "photos.zip", "application/zip"));
                    client.AppendMessage(folderName, message);

                    client.SelectFolder(folderName);
                    var sequenceNumber = client.ListMessages()[0].SequenceNumber;

                    var attachments = client.ListAttachments(sequenceNumber);
                    Console.WriteLine($"{attachments.Count} attachment(s):");
                    foreach (var info in attachments)
                        Console.WriteLine($"  {info.Name,-12} {info.MediaType,-26} {info.Size} bytes");

                    // Download only the smallest one.
                    var smallest = attachments.OrderBy(info => info.Size).First();
                    var attachment = client.FetchAttachment(sequenceNumber, smallest.Name);

                    var outputPath = outputDir/attachment.Name;
                    attachment.Save(outputPath);
                    Console.WriteLine($"\nSaved {attachment.Name} to {outputPath}");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
