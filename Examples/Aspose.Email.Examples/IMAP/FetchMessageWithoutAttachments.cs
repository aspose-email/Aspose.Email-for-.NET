// Demonstrates how to fetch a message without downloading its attachments.
//
// FetchMessage(sequenceNumber, ignoreAttachments: true) retrieves the headers, the body
// and the list of attachments, but not the attachment content - a large saving when you
// only need to show or index the text of a message.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class FetchMessageWithoutAttachments
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    // A test message with a 2 MB attachment.
                    var message = new MailMessage("from@example.com", "to@example.com",
                        "Quarterly figures", "The figures are attached.");
                    message.Attachments.Add(new Attachment(new MemoryStream(new byte[2 * 1024 * 1024]),
                        "figures.bin", "application/octet-stream"));
                    client.AppendMessage(folderName, message);

                    client.SelectFolder(folderName);
                    var sequenceNumber = client.ListMessages()[0].SequenceNumber;

                    var watch = Stopwatch.StartNew();
                    var light = client.FetchMessage(sequenceNumber, true);
                    Console.WriteLine($"Without attachment content: {watch.ElapsedMilliseconds} ms");
                    Console.WriteLine($"  subject: {light.Subject}");
                    Console.WriteLine($"  body:    {light.Body.Trim()}");
                    foreach (var attachment in light.Attachments)
                        Console.WriteLine($"  attachment listed: {attachment.Name}");

                    watch.Restart();
                    var full = client.FetchMessage(sequenceNumber);
                    Console.WriteLine($"\nWith attachment content:    {watch.ElapsedMilliseconds} ms");
                    foreach (var attachment in full.Attachments)
                        Console.WriteLine($"  attachment: {attachment.Name}, {attachment.ContentStream.Length} bytes");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
