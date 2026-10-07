// Demonstrates how to put a message into a folder on the server.
//
// AppendMessage uploads a MailMessage to the given folder and returns its unique id. The
// message does not travel through SMTP: it simply appears in the folder - which is how
// mail programs save drafts and copies of sent mail. SubscribeFolder adds the folder to
// the list that mail programs show.
//
// The example works in a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class AddingNewMessage
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);
                client.SubscribeFolder(folderName);

                try
                {
                    var message = new MailMessage("sender@example.com", "receiver@example.com",
                        "Saved draft", "This message was appended, not sent.");

                    var uid = client.AppendMessage(folderName, message);
                    Console.WriteLine($"Appended to '{folderName}', unique id: {uid}");

                    client.SelectFolder(folderName);
                    foreach (var info in client.ListMessages())
                        Console.WriteLine($"  {info.Subject} ({info.Size} bytes)");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
