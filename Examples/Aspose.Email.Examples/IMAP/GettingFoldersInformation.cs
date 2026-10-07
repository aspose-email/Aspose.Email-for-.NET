// Demonstrates how to list the folders of a mailbox with their message counts.
//
// ListFolders returns the folder names and attributes; GetFolderInfo asks the server for
// the status of one folder - total, new and recent message counts and whether it is
// read-only - without selecting it.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GettingFoldersInformation
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                foreach (var folder in client.ListFolders())
                {
                    if (!folder.Selectable)
                    {
                        Console.WriteLine($"{folder.Name} (container only, holds no messages)");
                        continue;
                    }

                    var info = client.GetFolderInfo(folder.Name);
                    Console.WriteLine(folder.Name);
                    Console.WriteLine($"  total: {info.TotalMessageCount}, new: {info.NewMessageCount}, " +
                                      $"recent: {info.RecentMessageCount}, read-only: {info.ReadOnly}");
                }
            }
        }
    }
}
