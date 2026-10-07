// Demonstrates how to download messages from the server into MailMessage objects.
//
// ListMessages returns lightweight summaries; FetchMessage downloads the complete
// message - body and attachments - by unique id, ready to read, convert or save.
// MessagesFromIMAPServerToDisk saves messages without parsing them.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class FetchEmailMessagesFromIMAPServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var newest = client.ListMessages()
                    .OrderByDescending(info => info.InternalDate)
                    .Take(5)
                    .ToList();

                Console.WriteLine($"The {newest.Count} newest message(s) in the Inbox:");

                foreach (var info in newest)
                {
                    var message = client.FetchMessage(info.UniqueId);

                    Console.WriteLine($"\n  {message.Subject}");
                    Console.WriteLine($"    from:        {message.From}");
                    Console.WriteLine($"    date:        {message.Date}");
                    Console.WriteLine($"    html body:   {message.IsBodyHtml}");
                    Console.WriteLine($"    attachments: {message.Attachments.Count}");
                }
            }
        }
    }
}
