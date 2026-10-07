// Demonstrates listing the messages of a folder together with all its subfolders.
//
// ListMessages(folderName, retrieveRecursively: true) walks the folder tree below the
// given folder and returns one combined list; ParentFolder tells which folder each
// message came from. ReadMessagesRecursively walks the tree itself and downloads the
// messages.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListingMessagesRecursively
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var messages = client.ListMessages(ImapFolderInfo.InBox, true);

                Console.WriteLine($"{messages.Count} message(s) in the Inbox and its subfolders:");

                foreach (var folder in messages.GroupBy(info => info.ParentFolder))
                    Console.WriteLine($"  {folder.Key}: {folder.Count()}");
            }
        }
    }
}
