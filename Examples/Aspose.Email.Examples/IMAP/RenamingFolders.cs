// Demonstrates that renaming a folder keeps its messages.
//
// RenameFolder changes the name on the server in one step; the messages stay inside,
// so there is nothing to copy. A new name that contains the hierarchy delimiter puts
// the folder under another parent, which makes renaming a way to move folders too.
// The example uses uniquely named folders and deletes them at the end.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RenamingFolders
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var oldName = "Aspose-Projects-" + suffix;
            var newName = "Aspose-Archive-" + suffix;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(oldName);
                client.AppendMessage(oldName,
                    new MailMessage("sender@example.com", "receiver@example.com", "Kick-off notes", "Body"));

                client.RenameFolder(oldName, newName);

                try
                {
                    Console.WriteLine($"Renamed '{oldName}' to '{newName}'.");
                    Console.WriteLine($"  old name exists:    {client.ExistFolder(oldName)}");
                    Console.WriteLine($"  messages under new: {client.GetFolderInfo(newName).TotalMessageCount}");
                }
                finally
                {
                    client.DeleteFolder(newName);
                }
            }
        }
    }
}
