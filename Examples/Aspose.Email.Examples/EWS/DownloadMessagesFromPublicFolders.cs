// Demonstrates reading messages from public folders.
//
// ListPublicFolders returns the top of the organisation's public folder tree;
// ListMessagesFromPublicFolder lists the messages of one public folder, and
// ListSubFolders goes one level deeper. The messages are saved as .msg files under
// Out - at most three per folder, to keep the run short. This needs public folders to
// be set up and the account to have read access.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DownloadMessagesFromPublicFolders
    {
        private const int MessagesPerFolder = 3;

        public static void Run()
        {
            var rootDir = Data.OutSub("EwsPublicFolders");

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folders = client.ListPublicFolders();
                Console.WriteLine($"{folders.Count} top-level public folder(s).");

                foreach (var folder in folders)
                    SaveFolder(client, folder, rootDir);
            }

            Console.WriteLine($"\nMessages saved under {rootDir}");
        }

        private static void SaveFolder(IEWSClient client, ExchangeFolderInfo folder, string parentDir)
        {
            var safeName = string.Concat(folder.DisplayName.Split(Path.GetInvalidFileNameChars()));
            var folderDir = Path.Combine(parentDir, safeName);
            Directory.CreateDirectory(folderDir);

            var messages = client.ListMessagesFromPublicFolder(folder).Take(MessagesPerFolder).ToList();
            for (var i = 0; i < messages.Count; i++)
            {
                var message = client.FetchMessage(messages[i].UniqueUri);
                message.Save(Path.Combine(folderDir, $"message-{i + 1}.msg"), SaveOptions.DefaultMsgUnicode);
            }

            Console.WriteLine($"{folder.DisplayName}: saved {messages.Count} message(s), " +
                              $"{folder.ChildFolderCount} subfolder(s)");

            if (folder.ChildFolderCount == 0)
                return;

            foreach (var child in client.ListSubFolders(folder))
                SaveFolder(client, child, folderDir);
        }
    }
}
