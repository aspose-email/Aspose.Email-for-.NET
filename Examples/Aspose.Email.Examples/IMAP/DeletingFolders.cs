// Demonstrates deleting a folder that holds messages and a subfolder.
//
// DeleteFolder removes the folder and its messages for good - there is no recycle bin.
// Delete subfolders first: some servers refuse to delete a folder that has children,
// others keep it as an empty placeholder that cannot hold messages.
// The example builds a uniquely named folder tree and deletes it.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class DeletingFolders
    {
        public static void Run()
        {
            var parentName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                if (!client.ExistFolder(parentName))
                    client.CreateFolder(parentName);

                var childName = parentName + client.Delimiter + "Old";
                client.CreateFolder(childName);

                client.AppendMessage(parentName,
                    new MailMessage("sender@example.com", "receiver@example.com", "Will be deleted", "Body"));

                Console.WriteLine($"'{parentName}': {client.GetFolderInfo(parentName).TotalMessageCount} message(s), " +
                                  $"{client.ListFolders(parentName).Count} subfolder(s)");

                foreach (var child in client.ListFolders(parentName))
                {
                    client.DeleteFolder(child.Name);
                    Console.WriteLine($"Deleted '{child.Name}'");
                }

                client.DeleteFolder(parentName);
                Console.WriteLine($"Deleted '{parentName}'");

                Console.WriteLine($"\n'{parentName}' still exists: {client.ExistFolder(parentName)}");
            }
        }
    }
}
