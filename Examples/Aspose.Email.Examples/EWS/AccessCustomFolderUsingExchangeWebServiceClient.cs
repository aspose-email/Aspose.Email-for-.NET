// Demonstrates finding a folder by name and reading the messages in it.
//
// Folders are addressed by URI, not by name. FolderExists looks a folder up by its
// display name under a parent folder and, through the out parameter, returns its
// ExchangeFolderInfo with the URI to work with. The example creates a uniquely named
// folder with two messages, finds it again, and deletes it at the end.

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class AccessCustomFolderUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inboxUri = client.MailboxInfo.InboxUri;
                var created = client.CreateFolder(inboxUri, folderName);

                try
                {
                    client.AppendMessages(created.Uri,
                        new MailMessage("sender@example.com", "receiver@example.com", "First", "Body"),
                        new MailMessage("sender@example.com", "receiver@example.com", "Second", "Body"));

                    ExchangeFolderInfo folder;
                    if (!client.FolderExists(inboxUri, folderName, out folder))
                    {
                        Console.WriteLine($"No folder named '{folderName}' under the Inbox.");
                        return;
                    }

                    Console.WriteLine($"Found '{folder.DisplayName}' ({folder.TotalCount} item(s)):");
                    foreach (var info in client.ListMessages(folder.Uri))
                    {
                        var message = client.FetchMessage(info.UniqueUri);
                        Console.WriteLine($"  {message.Subject}");
                    }
                }
                finally
                {
                    client.DeleteFolder(created.Uri, true);
                }
            }
        }
    }
}
