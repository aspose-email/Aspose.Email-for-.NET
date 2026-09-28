// Demonstrates how to move a folder, together with its messages and subfolders, under a
// new parent. Note the argument order of MoveFolder: the new parent comes first, then
// the folder to move.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class MoveFolderToNewParent
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var projectFolder = "Aspose-Project-" + suffix;
            var archiveFolder = "Aspose-Archive-" + suffix;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(projectFolder);
                client.CreateFolder(archiveFolder);

                client.AppendMessage(projectFolder,
                    new MailMessage("from@example.com", "to@example.com", "Project kickoff", "Agenda attached."));

                client.MoveFolder(archiveFolder, projectFolder);

                var movedFolder = archiveFolder + client.Delimiter + projectFolder;

                try
                {
                    Console.WriteLine($"'{projectFolder}' still at top level: {client.ExistFolder(projectFolder)}");

                    Console.WriteLine($"Subfolders of '{archiveFolder}':");
                    foreach (var folder in client.ListFolders(archiveFolder))
                    {
                        var messages = client.GetFolderInfo(folder.Name).TotalMessageCount;
                        Console.WriteLine($"  {folder.Name} ({messages} message(s))");
                    }
                }
                finally
                {
                    if (client.ExistFolder(movedFolder))
                        client.DeleteFolder(movedFolder);
                    if (client.ExistFolder(projectFolder))
                        client.DeleteFolder(projectFolder);
                    client.DeleteFolder(archiveFolder);
                }
            }
        }
    }
}
