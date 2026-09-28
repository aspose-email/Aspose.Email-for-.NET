// Demonstrates that deleting in IMAP takes two steps, and how to take back the first.
//
// DeleteMessage only sets the \Deleted flag; the message stays in the folder until the
// deletions are committed (EXPUNGE). Until then UndeleteMessage clears the flag again.
// AutoCommit is switched off so the client does not commit deletions by itself when the
// folder changes or the connection closes.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class UndeleteMarkedMessage
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.AutoCommit = false;
                client.CreateFolder(folderName);

                try
                {
                    var uid = client.AppendMessage(folderName,
                        new MailMessage("from@example.com", "to@example.com", "Keep me", "Body"));

                    client.SelectFolder(folderName);
                    Console.WriteLine($"Appended:              {DescribeState(client, uid)}");

                    client.DeleteMessage(uid);
                    Console.WriteLine($"After DeleteMessage:   {DescribeState(client, uid)}");

                    client.UndeleteMessage(uid);
                    Console.WriteLine($"After UndeleteMessage: {DescribeState(client, uid)}");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }

        private static string DescribeState(ImapClient client, string uid)
        {
            var info = client.ListMessage(uid);
            if (info == null)
                return "not listed";

            return info.Deleted ? $"'{info.Subject}' is marked as deleted" : $"'{info.Subject}' is not deleted";
        }
    }
}
