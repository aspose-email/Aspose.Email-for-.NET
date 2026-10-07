// Demonstrates how to copy one message to another folder and learn the unique id of the
// copy. The id is returned only when the server supports UIDPLUS (RFC 4315); otherwise
// CopyMessage returns null and you have to search the target folder for the copy.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class CopySingleMessage
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var sourceFolder = "Aspose-Source-" + suffix;
            var targetFolder = "Aspose-Target-" + suffix;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(sourceFolder);
                client.CreateFolder(targetFolder);

                try
                {
                    var message = new MailMessage("from@example.com", "to@example.com",
                        "Invoice 1042", "Please find the invoice.");
                    var uid = client.AppendMessage(sourceFolder, message);

                    client.SelectFolder(sourceFolder);
                    var copyUid = client.CopyMessage(uid, targetFolder);

                    Console.WriteLine($"Original uid in '{sourceFolder}': {uid}");
                    Console.WriteLine($"Copy uid in '{targetFolder}':     {copyUid ?? "(not reported, no UIDPLUS)"}");

                    var sourceCount = client.GetFolderInfo(sourceFolder).TotalMessageCount;
                    var targetCount = client.GetFolderInfo(targetFolder).TotalMessageCount;
                    Console.WriteLine($"\n'{sourceFolder}': {sourceCount} message(s)");
                    Console.WriteLine($"'{targetFolder}': {targetCount} message(s)");
                }
                finally
                {
                    client.DeleteFolder(sourceFolder);
                    client.DeleteFolder(targetFolder);
                }
            }
        }
    }
}
