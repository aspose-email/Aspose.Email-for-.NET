// Demonstrates how to leave the selected folder without committing pending deletions,
// using the UNSELECT command (RFC 3691).
//
// UnselectFolder() closes the folder the classic way (CLOSE), which also expunges the
// messages marked as deleted. UnselectFolder(doNotExpunge: true) sends UNSELECT
// instead: the folder is left, and the marked messages stay where they are.
// AutoCommit is switched off so that the client does not commit deletions on its own.
// The example re-checks the message afterwards and prints what actually happened.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class UnselectFolderWithoutExpunge
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();
                if (!client.UnselectSupported)
                {
                    Console.WriteLine("The server does not support UNSELECT.");
                    return;
                }

                client.AutoCommit = false;
                client.CreateFolder(folderName);

                try
                {
                    var uid = client.AppendMessage(folderName,
                        new MailMessage("from@example.com", "to@example.com", "Marked for deletion", "Body"));

                    client.SelectFolder(folderName);
                    client.DeleteMessage(uid);
                    Console.WriteLine("Message marked as deleted.");

                    client.UnselectFolder(true);
                    Console.WriteLine("Folder unselected without expunging.");

                    client.SelectFolder(folderName);
                    var messages = client.ListMessages();
                    Console.WriteLine($"Messages still in '{folderName}': {messages.Count}");
                    foreach (var info in messages)
                        Console.WriteLine($"  {info.Subject} (deleted flag: {info.Deleted})");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
