// Demonstrates how to move several messages to another folder in one call.
//
// On servers with the MOVE extension (RFC 6851) the move is a single atomic command.
// Elsewhere it is done as copy + delete, and commitDeletions decides whether the
// originals are expunged right away; pass false to leave that to a later CommitDeletes.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class MoveMultipleMessages
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var sourceFolder = "Aspose-Source-" + suffix;
            var archiveFolder = "Aspose-Archive-" + suffix;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(sourceFolder);
                client.CreateFolder(archiveFolder);

                try
                {
                    var uids = new List<string>();
                    for (var i = 1; i <= 5; i++)
                    {
                        uids.Add(client.AppendMessage(sourceFolder,
                            new MailMessage("from@example.com", "to@example.com", $"Newsletter {i}", "Body")));
                    }

                    Console.WriteLine($"MOVE extension supported: {client.MoveSupported}");

                    // Move the three oldest ones.
                    client.SelectFolder(sourceFolder);
                    client.MoveMessages(uids.Take(3), archiveFolder, true);

                    PrintFolder(client, sourceFolder);
                    PrintFolder(client, archiveFolder);
                }
                finally
                {
                    client.DeleteFolder(sourceFolder);
                    client.DeleteFolder(archiveFolder);
                }
            }
        }

        private static void PrintFolder(ImapClient client, string folderName)
        {
            client.SelectFolder(folderName);
            var messages = client.ListMessages();

            Console.WriteLine($"\n'{folderName}': {messages.Count} message(s)");
            foreach (var info in messages)
                Console.WriteLine("  " + info.Subject);
        }
    }
}
