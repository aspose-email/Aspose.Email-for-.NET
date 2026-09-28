// Demonstrates incremental synchronization with CONDSTORE (RFC 7162).
//
// Every change to a message - a new flag, a new message - raises the folder's
// modification sequence (MODSEQ). Remember HighestModSequence after a sync; next time
// ask only for messages whose MODSEQ is higher instead of re-reading the whole folder.
// ListMessages(long) does exactly that; ImapQueryBuilder.ModSeq expresses the same
// condition as part of a search.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SyncChangesWithModSequence
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.CondstoreSupported)
                {
                    Console.WriteLine("The server does not support CONDSTORE.");
                    return;
                }

                // Makes the server report MODSEQ values from now on.
                if (client.EnableSupported)
                    client.ClientCapabilities("CONDSTORE");

                client.CreateFolder(folderName);

                try
                {
                    string secondUid = null;
                    for (var i = 1; i <= 3; i++)
                    {
                        var uid = client.AppendMessage(folderName,
                            new MailMessage("from@example.com", "to@example.com", $"Report {i}", "Body"));
                        if (i == 2)
                            secondUid = uid;
                    }

                    // The "last sync": remember where the folder stands.
                    client.SelectFolder(folderName);
                    var lastSyncModSeq = client.CurrentFolder.HighestModSequence;
                    Console.WriteLine($"HIGHESTMODSEQ at the last sync: {lastSyncModSeq}");

                    // Somebody flags one message in the meantime.
                    client.AddMessageFlags(secondUid, ImapMessageFlags.Flagged);

                    Console.WriteLine("\nChanged since the last sync (ListMessages):");
                    foreach (var info in client.ListMessages(lastSyncModSeq))
                        Console.WriteLine($"  {info.Subject}: MODSEQ {info.ModificationSequence}, flags {info.Flags}");

                    var builder = new ImapQueryBuilder();
                    builder.ModSeq.GreaterOrEqualTo(lastSyncModSeq + 1);

                    Console.WriteLine("\nChanged since the last sync (search):");
                    foreach (var info in client.ListMessages(builder.GetQuery()))
                        Console.WriteLine($"  {info.Subject}");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
