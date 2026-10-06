// Demonstrates how to move a single message to another folder.
//
// MoveMessage takes the message's unique id (or sequence number) and the target folder.
// On servers with the MOVE extension (RFC 6851) this is one atomic command. Elsewhere it
// is done as copy + delete, and the overloads with commitDeletions decide whether the
// original is expunged right away. MoveMultipleMessages moves several at once.
//
// The example works in two uniquely named folders and deletes both at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class MoveMessage
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var sourceFolder = "Aspose-Source-" + suffix;
            var targetFolder = "Aspose-Target-" + suffix;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(sourceFolder);
                client.CreateFolder(targetFolder);

                try
                {
                    var uid = client.AppendMessage(sourceFolder,
                        new MailMessage("from@example.com", "to@example.com", "Move me", "Body"));

                    Console.WriteLine($"MOVE extension supported: {client.MoveSupported}");
                    Console.WriteLine("\nBefore the move:");
                    PrintFolder(client, sourceFolder);
                    PrintFolder(client, targetFolder);

                    client.SelectFolder(sourceFolder);
                    client.MoveMessage(uid, targetFolder, true);

                    Console.WriteLine("\nAfter the move:");
                    PrintFolder(client, sourceFolder);
                    PrintFolder(client, targetFolder);
                }
                finally
                {
                    client.DeleteFolder(sourceFolder);
                    client.DeleteFolder(targetFolder);
                }
            }
        }

        private static void PrintFolder(ImapClient client, string folderName)
        {
            client.SelectFolder(folderName);
            var messages = client.ListMessages();

            Console.WriteLine($"  '{folderName}': {messages.Count} message(s)");
            foreach (var info in messages)
                Console.WriteLine("    " + info.Subject);
        }
    }
}
