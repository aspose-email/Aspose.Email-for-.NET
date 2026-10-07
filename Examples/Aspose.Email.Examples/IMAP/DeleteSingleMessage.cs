// Demonstrates how to delete one message.
//
// Deleting in IMAP takes two steps: DeleteMessage marks the message with the \Deleted
// flag, and CommitDeletes (EXPUNGE) removes every marked message from the selected
// folder for good. Until the commit, UndeleteMarkedMessage shows how to take it back.
// The example works in a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class DeleteSingleMessage
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
                        new MailMessage("sender@example.com", "receiver@example.com", "Message to delete", "Body"));

                    client.SelectFolder(folderName);
                    Console.WriteLine($"Before: {client.ListMessages().Count} message(s) in '{folderName}'");

                    client.DeleteMessage(uid);
                    client.CommitDeletes();

                    Console.WriteLine($"After:  {client.ListMessages().Count} message(s) in '{folderName}'");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
