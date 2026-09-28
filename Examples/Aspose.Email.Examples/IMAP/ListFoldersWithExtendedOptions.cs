// Demonstrates LIST command extensions (RFC 5258) and special-use folders (RFC 6154).
//
// Selection options narrow down which folders are listed; return options ask the server
// for extra facts about each one - subscription state, whether it has children - in the
// same round trip. On servers with SPECIAL-USE, FolderType tells you which folder is
// Sent, Trash, Junk and so on, whatever it is called in the user's language.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListFoldersWithExtendedOptions
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.ExtendedListSupported)
                {
                    Console.WriteLine("The server does not support LIST-EXTENDED.");
                    return;
                }

                // All folders, with subscription state and child information.
                // A null parent folder lists from the top of the hierarchy.
                var folders = client.ListFolders(null, false, ListFoldersOptions.None,
                    ListFoldersReturnOptions.Subscribed | ListFoldersReturnOptions.Children);

                Console.WriteLine($"{folders.Count} folder(s):");
                foreach (var folder in folders)
                {
                    var role = folder.FolderType == ImapSpecialFolderTypes.NotSpecified ? "" : $"  [{folder.FolderType}]";
                    Console.WriteLine($"  {folder.Name}{role}");
                    Console.WriteLine($"    subscribed: {folder.Subscribed}, has children: {folder.HasChildren}, " +
                                      $"selectable: {folder.Selectable}");
                }

                // Only subscribed folders - plus the parents of subscribed subfolders, which
                // RecursiveMatch returns even when the parent itself is not subscribed.
                var subscribed = client.ListFolders(null, false,
                    ListFoldersOptions.Subscribed | ListFoldersOptions.RecursiveMatch,
                    ListFoldersReturnOptions.None);

                Console.WriteLine($"\n{subscribed.Count} subscribed folder(s) or parents of one:");
                foreach (var folder in subscribed)
                    Console.WriteLine($"  {folder.Name}{(folder.NonExistent ? "  (does not exist)" : "")}");
            }
        }
    }
}
