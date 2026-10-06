// Demonstrates a conditional flag update (STORE ... UNCHANGEDSINCE, RFC 7162).
//
// Two clients working on the same mailbox can overwrite each other's changes. Passing
// the MODSEQ you last saw makes the server apply the change only if nobody has touched
// the message since; otherwise the update is refused, and you can re-read the message
// and decide again. The example prints the outcome of both attempts.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class UpdateFlagsIfUnchanged
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

                if (client.EnableSupported)
                    client.ClientCapabilities("CONDSTORE");

                client.CreateFolder(folderName);

                try
                {
                    var uid = client.AppendMessage(folderName,
                        new MailMessage("from@example.com", "to@example.com", "Shared task", "Body"));

                    client.SelectFolder(folderName);
                    var seenModSeq = client.ListMessage(uid).ModificationSequence;
                    Console.WriteLine($"MODSEQ when we read the message: {seenModSeq}");

                    // First update: nothing has changed since we read it, so it goes through.
                    client.AddMessageFlags(uid, ImapMessageFlags.Flagged, seenModSeq);
                    Report(client, uid, "After the first update: ");

                    // Second update still quotes the old MODSEQ, which is now out of date.
                    try
                    {
                        client.AddMessageFlags(uid, ImapMessageFlags.IsRead, seenModSeq);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Second update rejected: {ex.Message}");
                    }
                    Report(client, uid, "After the second update:");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }

        private static void Report(ImapClient client, string uid, string title)
        {
            var info = client.ListMessage(uid);
            Console.WriteLine($"{title} flags {info.Flags}, MODSEQ {info.ModificationSequence}");
        }
    }
}
