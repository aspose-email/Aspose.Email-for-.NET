// Demonstrates clearing message flags - for example marking a message as unread again.
//
// RemoveMessageFlags clears only the given flags and leaves the rest in place. Removing
// IsRead (\Seen) makes a message show as unread in mail programs again.
// The example works in a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RemovingMessageFlags
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    var uid = client.AppendMessage(folderName,
                        new MailMessage("sender@example.com", "receiver@example.com", "Read me later", "Body"));

                    client.SelectFolder(folderName);
                    client.AddMessageFlags(uid, ImapMessageFlags.IsRead | ImapMessageFlags.Flagged);
                    Report(client, uid, "Read and flagged:   ");

                    client.RemoveMessageFlags(uid, ImapMessageFlags.IsRead);
                    Report(client, uid, "After Remove(IsRead):");
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
            Console.WriteLine($"{title} {info.Flags}  (read: {info.IsRead}, flagged: {info.Flagged})");
        }
    }
}
