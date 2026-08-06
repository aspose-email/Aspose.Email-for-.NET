// Demonstrates how to flip the read/unread flag of a whole batch of PST messages
// in a single call instead of updating each message separately.

using System;
using System.Collections.Generic;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetMessageReadStatusInPST
    {
        public static void Run()
        {
            var pstPath = Data.Out/"SetMessageReadStatusInPST_out.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var inbox = pst.CreatePredefinedFolder("Inbox", StandardIpmFolder.Inbox);

                for (var i = 1; i <= 5; i++)
                {
                    inbox.AddMessage(new MapiMessage(
                        "from@domain.com", "to@domain.com", $"Message {i}", $"Body of message {i}"));
                }

                // The ids and the SetReadStatus call have to come from the same
                // FolderInfo instance - each lookup builds its own index.
                var idsForProcessing = new List<string>();
                foreach (var entryId in inbox.EnumerateMessagesEntryId())
                    idsForProcessing.Add(entryId);

                Console.WriteLine($"Messages in Inbox: {idsForProcessing.Count}");
                Console.WriteLine($"Unread at start:   {UnreadCount(pst)}");

                // One call updates the whole batch.
                inbox.SetReadStatus(idsForProcessing, false);
                Console.WriteLine($"Unread after marking them unread: {UnreadCount(pst)}");

                inbox.SetReadStatus(idsForProcessing, true);
                Console.WriteLine($"Unread after marking them read:   {UnreadCount(pst)}");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }

        // ContentUnreadCount is read when the FolderInfo is obtained, so the folder
        // has to be looked up again to see the updated value.
        private static int UnreadCount(PersonalStorage pst) =>
            pst.RootFolder.GetSubFolder("Inbox").ContentUnreadCount;
    }
}
