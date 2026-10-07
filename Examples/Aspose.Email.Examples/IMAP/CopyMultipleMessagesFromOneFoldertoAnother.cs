// Demonstrates copying several messages to another folder in one command.
//
// CopyMessages takes a list of unique ids (or sequence numbers, or message infos) from
// the selected folder and copies them to the target folder; the originals stay where
// they are. CopySingleMessage copies one message and gets the new unique id back.
// The example works in two uniquely named folders and deletes both at the end.

using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class CopyMultipleMessagesFromOneFoldertoAnother
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
                    var uids = new List<string>();
                    for (var i = 1; i <= 3; i++)
                    {
                        uids.Add(client.AppendMessage(sourceFolder,
                            new MailMessage("sender@example.com", "receiver@example.com", $"Report {i}", "Body")));
                    }

                    client.SelectFolder(sourceFolder);
                    client.CopyMessages(uids, targetFolder);

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

            Console.WriteLine($"'{folderName}': {messages.Count} message(s)");
            foreach (var info in messages)
                Console.WriteLine("  " + info.Subject);
        }
    }
}
