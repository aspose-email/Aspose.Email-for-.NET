// Demonstrates how to change the flags of many messages with one command.
//
// Every flag method has overloads that take a list of message infos, a list of unique
// ids or sequence numbers, or a first/last range. AddMessageFlags and
// RemoveMessageFlags adjust individual flags; ChangeMessageFlags replaces the whole set.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SetFlagsOnMultipleMessages
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    var uids = new List<string>();
                    for (var i = 1; i <= 5; i++)
                    {
                        uids.Add(client.AppendMessage(folderName,
                            new MailMessage("from@example.com", "to@example.com", $"Task {i}", "Body")));
                    }

                    client.SelectFolder(folderName);
                    var messages = client.ListMessages();

                    // A list of message infos: flag the first three and mark them as read.
                    client.AddMessageFlags(messages.Take(3), ImapMessageFlags.Flagged | ImapMessageFlags.IsRead);
                    PrintFlags(client, "After AddMessageFlags on the first three:");

                    // A range of unique ids: clear the read flag on all five again.
                    client.RemoveMessageFlags(uids.First(), uids.Last(), ImapMessageFlags.IsRead);
                    PrintFlags(client, "After RemoveMessageFlags on the uid range:");

                    // A list of unique ids: replace whatever the last two had with a keyword.
                    client.ChangeMessageFlags(uids.Skip(3), ImapMessageFlags.Keyword("followup"));
                    PrintFlags(client, "After ChangeMessageFlags on the last two:");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }

        private static void PrintFlags(ImapClient client, string title)
        {
            Console.WriteLine(title);
            foreach (var info in client.ListMessages())
                Console.WriteLine($"  {info.Subject}: {info.Flags}");
            Console.WriteLine();
        }
    }
}
