// Demonstrates getting message summaries for known item URIs.
//
// When you already hold item URIs - saved from AppendMessage, a previous listing or
// your own database - ListMessages(ids) returns their current summaries in one
// request, without listing whole folders. The example works in a uniquely named folder
// and deletes it at the end.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetExchangeMessageInfoFromMessageURI
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 3)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Message {i}", "Body"));
                    var uris = client.AppendMessages(folder.Uri, messages);

                    foreach (var info in client.ListMessages(uris))
                        Console.WriteLine($"{info.Subject}: {info.Size} bytes, read: {info.IsRead}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
