// Demonstrates how to find the Sent, Drafts, Trash and other special folders.
//
// Folder names differ between servers and languages ("Sent Items", "[Gmail]/Sent Mail",
// "Gesendet"...). Servers with SPECIAL-USE (RFC 6154) mark the folders by role instead,
// and MailboxInfo exposes them by that role. A property is null when the server does
// not mark a folder for it.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapSpecialUseMailboxes
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var mailbox = client.MailboxInfo;

                Console.WriteLine($"SPECIAL-USE supported: {client.SpecialUseSupported}\n");
                Show("Inbox", mailbox.Inbox);
                Show("Drafts", mailbox.DraftMessages);
                Show("Sent", mailbox.SentMessages);
                Show("Junk", mailbox.JunkMessages);
                Show("Trash", mailbox.Trash);
                Show("Archive", mailbox.ArchivedMessages);
                Show("All mail", mailbox.AllMessages);
                Show("Flagged", mailbox.FlaggedMessages);
                Show("Important", mailbox.Important);
            }
        }

        private static void Show(string role, ImapFolderInfo folder)
        {
            Console.WriteLine($"  {role,-10} {folder?.Name ?? "(not marked)"}");
        }
    }
}
