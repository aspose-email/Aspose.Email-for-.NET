// Demonstrates the CHECK command, exposed as RequestCheckpoint.
//
// CHECK asks the server to do any housekeeping for the selected folder, such as writing
// its state to disk. Most modern servers do that continuously and treat CHECK like
// NOOP, but it is harmless to send after a batch of changes to a folder.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RequestMailboxCheckpoint
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                // CHECK applies to the selected folder, so select one first.
                client.SelectFolder(ImapFolderInfo.InBox);

                client.RequestCheckpoint();
                Console.WriteLine($"Checkpoint requested for '{client.CurrentFolder.Name}'.");
            }
        }
    }
}
