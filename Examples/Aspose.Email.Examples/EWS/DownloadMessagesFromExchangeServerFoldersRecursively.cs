// Demonstrates walking the whole folder tree of a mailbox and saving messages to disk.
//
// ListSubFolders returns the children of a folder, so a recursive walk from the root
// reaches every level. Mail folders are recognised by their folder class, and their
// messages are saved as .msg files in a matching directory tree under Out. To keep the
// run short, only the three newest messages of each folder are downloaded.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DownloadMessagesFromExchangeServerFoldersRecursively
    {
        private const int MessagesPerFolder = 3;

        public static void Run()
        {
            var rootDir = Data.OutSub("EwsMailbox");

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var folder in client.ListSubFolders(client.MailboxInfo.RootUri))
                    SaveFolder(client, folder, rootDir);
            }

            Console.WriteLine($"\nMessages saved under {rootDir}");
        }

        private static void SaveFolder(IEWSClient client, ExchangeFolderInfo folder, string parentDir)
        {
            var folderDir = Path.Combine(parentDir, SafeName(folder.DisplayName));

            if (folder.FolderClass == "IPF.Note" && folder.TotalCount > 0)
            {
                Directory.CreateDirectory(folderDir);

                var newest = client.ListMessages(folder.Uri)
                    .OrderByDescending(info => info.InternalDate)
                    .Take(MessagesPerFolder)
                    .ToList();

                for (var i = 0; i < newest.Count; i++)
                {
                    var message = client.FetchMessage(newest[i].UniqueUri);
                    message.Save(Path.Combine(folderDir, $"message-{i + 1}.msg"), SaveOptions.DefaultMsgUnicode);
                }

                Console.WriteLine($"{folder.DisplayName}: saved {newest.Count} of {folder.TotalCount}");
            }

            if (folder.ChildFolderCount == 0)
                return;

            foreach (var child in client.ListSubFolders(folder.Uri))
                SaveFolder(client, child, folderDir);
        }

        private static string SafeName(string name)
        {
            return string.Concat(name.Split(Path.GetInvalidFileNameChars()));
        }
    }
}
