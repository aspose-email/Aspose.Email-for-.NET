// Demonstrates the folder operations Graph offers: create, rename, nest, copy, move
// and delete.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphFolders
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                // A folder created without a parent lands at the top level.
                var parent = client.CreateFolder("Aspose.Email examples");
                Console.WriteLine($"Created: {parent.DisplayName} ({parent.ItemId})");

                // With a parent id it becomes a subfolder.
                var child = client.CreateFolder(parent.ItemId, "Processed");
                Console.WriteLine($"Created: {parent.DisplayName}/{child.DisplayName}");

                // Renaming is an update of the display name.
                child.DisplayName = "Archived";
                var renamed = client.UpdateFolder(child);
                Console.WriteLine($"Renamed to: {renamed.DisplayName}");

                // Listing the children of a specific folder.
                var children = client.ListFolders(parent.ItemId, null);
                Console.WriteLine($"{parent.DisplayName} has {children.Count} subfolder(s)");

                // Copy and move both take the destination parent first.
                var inbox = client.GetFolder(KnownFolders.Inbox);
                var copied = client.CopyFolder(inbox.ItemId, renamed.ItemId);
                Console.WriteLine($"Copied under the Inbox as {copied.ItemId}");

                var moved = client.MoveFolder(parent.ItemId, copied.ItemId);
                Console.WriteLine($"Moved back under {parent.DisplayName} as {moved.ItemId}");

                // Reading a folder back by id.
                var reloaded = client.GetFolder(parent.ItemId);
                Console.WriteLine($"Reloaded: {reloaded.DisplayName}, " +
                                  $"{reloaded.ContentCount} item(s), subfolders: {reloaded.HasSubFolders}");

                // Delete removes the folder and everything inside it.
                client.Delete(parent.ItemId);
                Console.WriteLine($"Deleted {parent.DisplayName}");

                var remaining = client.ListFolders(null)
                    .Count(f => f.DisplayName == "Aspose.Email examples");
                Console.WriteLine($"Folders left with that name: {remaining}");
            }
        }
    }
}
