// Demonstrates walking the whole folder tree of a mailbox and saving messages to disk.
//
// ListFolders(name) returns the subfolders of a folder, so a recursive walk reaches every
// level; folders that are only containers (not selectable) are skipped. The messages are
// saved as .msg files in a matching directory tree under Out. To keep the run short, only
// the three newest messages of each folder are downloaded.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ReadMessagesRecursively
    {
        private const int MessagesPerFolder = 3;

        public static void Run()
        {
            var rootDir = Data.OutSub("ImapMailbox");
            var visited = new HashSet<string>();

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                foreach (var folder in client.ListFolders())
                    SaveFolder(client, folder, rootDir, visited);
            }

            Console.WriteLine($"\n{visited.Count} folder(s) visited, messages saved under {rootDir}");
        }

        private static void SaveFolder(ImapClient client, ImapFolderInfo folder, string rootDir, HashSet<string> visited)
        {
            // ListFolders() may already include subfolders, so skip what was done before.
            if (!visited.Add(folder.Name))
                return;

            if (folder.Selectable)
            {
                var folderDir = Path.Combine(rootDir, ToRelativePath(folder.Name, client.Delimiter));
                Directory.CreateDirectory(folderDir);

                client.SelectFolder(folder.Name);
                var newest = client.ListMessages()
                    .OrderByDescending(info => info.InternalDate)
                    .Take(MessagesPerFolder)
                    .ToList();

                foreach (var info in newest)
                {
                    var message = client.FetchMessage(info.UniqueId);
                    message.Save(Path.Combine(folderDir, info.UniqueId + ".msg"), SaveOptions.DefaultMsgUnicode);
                }

                Console.WriteLine($"{folder.Name}: saved {newest.Count} of {client.CurrentFolder.TotalMessageCount}");
            }

            if (folder.NoInferiors)
                return;

            foreach (var subfolder in client.ListFolders(folder.Name))
                SaveFolder(client, subfolder, rootDir, visited);
        }

        // Turns "Inbox/Projects/2026" into a relative path, dropping characters that are
        // not allowed in file names.
        private static string ToRelativePath(string folderName, string delimiter)
        {
            var parts = string.IsNullOrEmpty(delimiter)
                ? new[] { folderName }
                : folderName.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);

            var invalid = Path.GetInvalidFileNameChars();
            var safeParts = parts.Select(part => new string(part.Select(c => invalid.Contains(c) ? '_' : c).ToArray()));

            return Path.Combine(safeParts.ToArray());
        }
    }
}
