// Demonstrates setting message flags, and the difference between adding and replacing.
//
// AddMessageFlags sets the given flags and keeps the others; ChangeMessageFlags replaces
// the message's whole flag set with the given one. Flags are combined with |. The
// standard ones are \Seen (IsRead), \Answered, \Flagged, \Deleted and \Draft.
// The example works in a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SettingMessageFlags
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
                        new MailMessage("sender@example.com", "receiver@example.com", "Flag me", "Body"));

                    client.SelectFolder(folderName);
                    Report(client, uid, "Appended:               ");

                    client.AddMessageFlags(uid, ImapMessageFlags.IsRead);
                    Report(client, uid, "After Add(IsRead):      ");

                    client.AddMessageFlags(uid, ImapMessageFlags.Flagged);
                    Report(client, uid, "After Add(Flagged):     ");

                    client.ChangeMessageFlags(uid, ImapMessageFlags.Answered);
                    Report(client, uid, "After Change(Answered): ");
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
            Console.WriteLine($"{title}{info.Flags}  " +
                              $"(read: {info.IsRead}, flagged: {info.Flagged}, answered: {info.Answered})");
        }
    }
}
