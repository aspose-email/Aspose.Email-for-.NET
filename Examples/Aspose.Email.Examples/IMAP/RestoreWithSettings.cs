// Demonstrates restoring a PST to the mailbox with RestoreSettings.
//
// Besides recursion, RestoreSettings can reconnect and retry like BackupSettings, and
// reports progress through BeforeItemCallback, which is called for every folder and
// message before it is uploaded.
//
// RemoveNonexistentFolders and RemoveNonexistentItems make the mailbox an exact mirror
// of the PST by DELETING folders and messages that the PST does not contain. They are
// left off here on purpose - turn them on only when that is what you want.
//
// The example builds a small PST with one uniquely named folder, restores it, and then
// removes that folder from the mailbox again.

using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Common.Delegate;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RestoreWithSettings
    {
        public static void Run()
        {
            var folderName = "Aspose-Restore-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var pstPath = Data.Out/"RestoreWithSettings_in.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var folder = pst.RootFolder.AddSubFolder(folderName);
                for (var i = 1; i <= 3; i++)
                {
                    folder.AddMessage(new MapiMessage("from@example.com", "to@example.com",
                        $"Restored message {i}", "Created by the RestoreWithSettings example."));
                }
            }

            var settings = new RestoreSettings
            {
                Recursive = true,
                RestoreConnection = true,
                NumberOfAttemptsToRrepeat = 3,
                TimeoutBetweenAttempts = 5000,
                RemoveNonexistentFolders = false,
                RemoveNonexistentItems = false,
                BeforeItemCallback = ReportItem
            };

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var pst = PersonalStorage.FromFile(pstPath, false))
            {
                Console.WriteLine($"Restoring {pstPath}:");
                client.Restore(pst, settings);

                try
                {
                    ImapFolderInfo restored;
                    Console.WriteLine(client.ExistFolder(folderName, out restored)
                        ? $"\n'{folderName}' now holds {client.GetFolderInfo(folderName).TotalMessageCount} message(s)."
                        : $"\n'{folderName}' was not found at the top level of the mailbox.");
                }
                finally
                {
                    if (client.ExistFolder(folderName))
                        client.DeleteFolder(folderName);
                }
            }
        }

        private static void ReportItem(ItemCallbackArgs args)
        {
            var folder = args.Item as FolderInfo;
            var message = args.Item as MapiMessage;
            var messageInfo = args.Item as MessageInfo;

            if (folder != null)
                Console.WriteLine($"  folder:  {folder.DisplayName}");
            else if (message != null)
                Console.WriteLine($"  message: {message.Subject}");
            else if (messageInfo != null)
                Console.WriteLine($"  message: {messageInfo.Subject}");
            else
                Console.WriteLine($"  {args.Item?.GetType().Name}");
        }
    }
}
