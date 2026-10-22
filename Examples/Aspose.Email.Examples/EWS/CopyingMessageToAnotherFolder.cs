// Demonstrates copying a message to another folder.
//
// CopyItem copies the item to the destination folder and returns the URI of the copy;
// the original stays where it was. The example works in two uniquely named folders and
// deletes both at the end.

using System;
using System.Collections.Specialized;

namespace Aspose.Email.Examples.EWS
{
    internal static class CopyingMessageToAnotherFolder
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var source = client.CreateFolder(client.MailboxInfo.InboxUri, "Aspose-Source-" + suffix);
                var target = client.CreateFolder(client.MailboxInfo.InboxUri, "Aspose-Target-" + suffix);

                try
                {
                    var uri = client.AppendMessage(source.Uri,
                        new MailMessage("from@example.com", "to@example.com", "Release notes", "Version 2.0 is out."));

                    var copyUri = client.CopyItem(uri, target.Uri);
                    Console.WriteLine($"Copied: {client.FetchMessage(copyUri).Subject}");

                    Console.WriteLine($"  '{source.DisplayName}': {client.ListMessages(source.Uri).Count} message(s)");
                    Console.WriteLine($"  '{target.DisplayName}': {client.ListMessages(target.Uri).Count} message(s)");
                }
                finally
                {
                    client.DeleteFolders(new StringCollection { source.Uri, target.Uri }, true);
                }
            }
        }
    }
}
