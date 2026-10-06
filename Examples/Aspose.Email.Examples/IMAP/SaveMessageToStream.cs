// Demonstrates how to download a message into a stream instead of a file.
//
// SaveMessage writes the raw message data as the server stores it, which is what you
// want for archiving, for computing a hash, or for handing the message to another
// component without an intermediate file. MailMessage.Load parses it when needed.

using System;
using System.IO;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SaveMessageToStream
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                var messages = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 1);
                if (messages.Count == 0)
                {
                    Console.WriteLine("The Inbox is empty.");
                    return;
                }

                var uid = messages[0].UniqueId;

                using (var stream = new MemoryStream())
                {
                    client.SaveMessage(uid, stream);
                    Console.WriteLine($"Downloaded message {uid}: {stream.Length} bytes");

                    stream.Position = 0;
                    var message = MailMessage.Load(stream);

                    Console.WriteLine($"  subject:     {message.Subject}");
                    Console.WriteLine($"  from:        {message.From}");
                    Console.WriteLine($"  attachments: {message.Attachments.Count}");
                }
            }
        }
    }
}
